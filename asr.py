import collections
import datetime
import os

import webrtcvad
from langchain_mistralai import ChatMistralAI
import queue
import threading
import time
import numpy as np
import sounddevice as sd
import soundfile as sf
import win32com.client
import io
from groq import Groq

from dotenv import load_dotenv

from ListNode import ListNode
from read_file import read_file
from langchain_core.tools import tool

from langgraph.graph import START, END, StateGraph
from langgraph.checkpoint.memory import MemorySaver

from langchain_core.prompts.chat import (
    ChatPromptTemplate,
    SystemMessagePromptTemplate,
    HumanMessagePromptTemplate,
    MessagesPlaceholder
)


from pydantic import BaseModel
from typing import List
from collections import deque
from langchain_core.messages import BaseMessage

from config import Settings

load_dotenv()
settings = Settings()


os.environ["GROQ_API_KEY"] = os.getenv("GROQ_API_KEY")

frame_size = int(settings.sample_rate * settings.frame_duration_ms / 1000)
last_recognize_time = 0
block_size = frame_size
vad = webrtcvad.Vad(0)

channels = 1
dtype = "float32"
buffer = np.array([])
file_path = r"presentation.pptx"
ppAPP = None
slide_contexts = None
slide_contexts_json = None
count_slides = None
slideshow = None
presentation_settings = None
presentation = None

if not os.path.isabs(file_path):
    file_path = os.path.abspath(file_path)
file_path = file_path.replace("\\", "//")
was_audio = 0
was_audio_recognized = 0
was_LLM_recognized = 0
num_padding_frames = int(settings.padding_ms / settings.frame_duration_ms)
vad_ring = deque(maxlen=num_padding_frames)
vad_voices = []
vad_triggered = False
command_queue = queue.Queue()
client = Groq()
# model = ChatGroq(model_name=settings.llm_model_name, temperature=0)
model = ChatMistralAI(model=settings.llm_model_name, temperature=0)
history_slide = ListNode()

slide_contexts_json = read_file(file_path=file_path)
slide_contexts_str = "\n".join(
    [f"Слайд {slide['slide_number']}: {slide['slide_context']}" for slide in slide_contexts_json])
system_template = SystemMessagePromptTemplate.from_template(settings.system_instructions)


human_message = HumanMessagePromptTemplate.from_template(settings.human_template)

chat_prompt = ChatPromptTemplate.from_messages([
    system_template,
    MessagesPlaceholder("messages"),
    human_message
])

chat_prompt = chat_prompt.partial(slide_context=slide_contexts_str)


@tool
def next_slide():
    """Move to the next slide"""
    if count_slides > slideshow.View.Slide.SlideIndex:
        print("Переход на следующий слайд.")
        slideshow.View.Next()
    else:
        print("Это последний слайд.")


@tool
def prev_slide():
    """Move to 1 slide back"""
    if slideshow.View.Slide.SlideIndex > 1:
        print("Переход на предыдущий слайд.")
        slideshow.View.Previous()
    else:
        print("Это первый слайд.")


@tool
def move_to_slide(slide_number: str):
    """Move to n-th slide"""
    try:
        slide_number = int(slide_number)
        if 1 <= slide_number <= count_slides:
            print(f"Переход к {slide_number} слайду.")
            slideshow.View.GotoSlide(slide_number)
        else:
            print(f"Недопустимый номер слайда {slide_number}.")
    except ValueError:
        print("Недопустимый номер слайда. Ожидалось число.")


@tool
def close_slideshow():
    """Close slideshow"""
    if slideshow:
        print("Презентация закрыта.")
        slideshow.View.Exit()

@tool
def back_slide():
    """Return to slide what was before this"""
    global history_slide
    if history_slide.prev:
        history_slide = history_slide.prev
    print(f"Возвращение обратно на {history_slide.val} слайд.")
    move_to_slide.invoke(str(history_slide.val))


@tool
def no_move():
    """Doing nothing if not need to do something"""
    print("Ничего не делать. Выполнено!")


tools = [next_slide, prev_slide, move_to_slide, close_slideshow, back_slide, no_move]

model_with_tools = model.bind_tools(tools)

prompt_and_model = chat_prompt | model_with_tools


class PresentationState(BaseModel):
    messages: List[BaseMessage]
    recognized_speech: str
    slide_now: int


def model_node(state: PresentationState):
    input_dict = {
        "messages": state.messages,
        "recognized_speech": state.recognized_speech,
        "slide_now": state.slide_now
    }
    result = prompt_and_model.invoke(input_dict)
    all_messages = state.messages + [result]
    trimmed_messages = all_messages[-5:]
    return {"messages": trimmed_messages}


workflow = (StateGraph(PresentationState)
            .add_node("model", model_node)
            .add_edge(START, "model")
            .add_edge("model", END)
            )

memory = MemorySaver()
model_with_history = workflow.compile(checkpointer=memory)
config = {"configurable": {"thread_id": "unique"}}


def run_powerPoint(file_path_pptx=""):
    global file_path, count_slides, ppAPP, presentation, presentation_settings, slideshow
    if file_path_pptx:
        file_path = file_path_pptx
    if not os.path.exists(file_path):
        print(f"Такого файла {file_path} не существует")
        return
    try:
        ppAPP = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as e:
        ppAPP = win32com.client.Dispatch("PowerPoint.Application")
        ppAPP.Visible = True
    try:
        presentation = ppAPP.Presentations.Open(file_path)
        presentation_settings = presentation.SlideShowSettings
        presentation_settings.Run()
        for _ in range(10):
            if ppAPP.SlideShowWindows.Count:
                slideshow = ppAPP.SlideShowWindows(1)
                break
            time.sleep(1)
        count_slides = presentation.Slides.Count

    except Exception as e:
        print(f"Ошибка: {e}. Пожалуйста, укажите полный путь к презентации.")


def exit_pp():
    global ppAPP
    if ppAPP:
        print("PowerPoint закрыт")
        ppAPP.DisplayAlerts = False
        ppAPP.Quit()
    del ppAPP

def recognize(vad_frames_bytes, chunk_id):
    global last_recognize_time
    now = time.time()
    if now - last_recognize_time < settings.min_gap:
        print("Пропущено")
        return
    last_recognize_time = time.time()
    print(f"Аудио №{chunk_id} отправлено на распознавание в {datetime.datetime.now()}")
    audio_int16 = np.frombuffer(vad_frames_bytes, dtype=np.int16)
    audio_float = audio_int16.astype(np.float32) / 32767
    wav_buffer = io.BytesIO()
    sf.write(wav_buffer, audio_float, samplerate=settings.sample_rate, format="WAV", subtype="PCM_16")
    wav_buffer.seek(0)

    recognized_text = client.audio.transcriptions.create(
        file=("recording.wav", wav_buffer),
        model=settings.asr_model_name,
        language=settings.language,
        response_format="text"
    )
    print(f"Распознан текст: {recognized_text}")
    process_command(recognized_text)


def audio_callback(indata, frames, time, status):
    global was_audio, vad_ring, vad_voices, vad_triggered
    frame = indata.flatten()

    pcm = (frame * 32767).astype(np.int16).tobytes()
    is_speech = vad.is_speech(pcm, settings.sample_rate)
    if not vad_triggered:
        vad_ring.append((pcm, is_speech))
        voices = sum(1 for _, s in vad_ring if s)
        if voices > settings.ratio * vad_ring.maxlen:
            vad_triggered = True
            vad_voices = [p for p, _ in vad_ring]
            vad_ring.clear()
    else:
        vad_voices.append(pcm)
        vad_ring.append((pcm, is_speech))
        silences = sum(1 for _, s in vad_ring if not s)
        if silences > settings.ratio * vad_ring.maxlen:
            segment = b"".join(vad_voices)
            vad_triggered = False
            vad_ring.clear()
            vad_voices =[]
            was_audio += 1
            chunk_id = was_audio
            threading.Thread(target=recognize, args=(segment,chunk_id)).start()

def process_command(recognized_speech):
    try:
        global was_LLM_recognized
        was_LLM_recognized += 1
        print(f"Обработка команды №{was_LLM_recognized} начата в {datetime.datetime.now()}.")
        slide_now = history_slide.val
        state = PresentationState(
            messages=[],
            recognized_speech=recognized_speech,
            slide_now=slide_now
        )

        # current_state = model_with_history.checkpointer.get(config=config)
        # if current_state:
        #     print(f"История: {current_state}")
        # else:
        #     print("Истории нет")

        output = model_with_history.invoke(input=state, config=config)
        print(f"Команда №{was_LLM_recognized} с текстом {recognized_speech} обработана в {datetime.datetime.now()}.")
        messages = output.get("messages", [])
        if not messages or not hasattr(messages[-1], "tool_calls"):
            print("Не был вызван tools.")
            return
        command_queue.put(messages[-1].tool_calls[0])

    except Exception as e:
        print(f"Ошибка при обработке распознанной речи: {e}.")

def do_command(tool_call):
    global history_slide
    tool_name = tool_call["name"]
    tool_args = tool_call["args"]
    print(f"tool: {tool_name}, args: {tool_args}")
    for tool in tools:
        if tool.name == tool_name:
            tool.invoke(input=tool_args)
            slide_now = slideshow.View.Slide.SlideIndex
            if slide_now != history_slide.val:
                history_slide = ListNode(val=slide_now, prev=history_slide)
                print(f"В историю добавлен {history_slide.val}, предыдущий: {history_slide.prev.val}")
            break

def start_audio():
    with sd.InputStream(
        samplerate=settings.sample_rate,
        blocksize=block_size,
        channels=channels,
        callback=audio_callback,
        dtype=dtype
    ):
        print("Запись начата")
        try:
            while True:
                try:
                    command = command_queue.get_nowait()
                    print("Команда обрабатывается.")
                    do_command(command)
                except queue.Empty:
                    time.sleep(0.001)
        except KeyboardInterrupt:
            print("Запись остановлена.")

try:
    run_powerPoint()
    start_audio()
finally:
    exit_pp()
