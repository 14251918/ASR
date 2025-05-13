import datetime
import io
import os
import queue
import threading
import time
from collections import deque
from typing import List

import numpy as np
import sounddevice as sd
import soundfile as sf
import webrtcvad
from dotenv import load_dotenv
from groq import Groq
from langchain_core.messages import BaseMessage
from langchain_core.prompts.chat import (
    ChatPromptTemplate,
    SystemMessagePromptTemplate,
    HumanMessagePromptTemplate,
    MessagesPlaceholder,
)
from langchain_mistralai import ChatMistralAI
from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import START, END, StateGraph
from pydantic import BaseModel

from LinkedList import LinkedList
from PresentationSettings import presentation
from config import settings
from tools import tools

load_dotenv()

os.environ["GROQ_API_KEY"] = os.getenv("GROQ_API_KEY")
os.environ["MISTRAL_API_KEY"] = os.getenv("MISTRAL_API_KEY")

frame_size = int(settings.sample_rate * settings.frame_duration_ms / 1000)
last_recognize_time = 0
vad = webrtcvad.Vad(0)

num_padding_frames = int(settings.padding_ms / settings.frame_duration_ms)
vad_ring = deque(maxlen=num_padding_frames)
vad_voices = []
vad_triggered = False
command_queue = queue.Queue()
# model = ChatGroq(model_name=settings.llm_model_name, temperature=0)
model = ChatMistralAI(model=settings.llm_model_name, temperature=0)

system_template = SystemMessagePromptTemplate.from_template(settings.system_instructions)
human_message = HumanMessagePromptTemplate.from_template(settings.human_template)

chat_prompt = ChatPromptTemplate.from_messages([
    system_template,
    MessagesPlaceholder("messages"),
    human_message
])

chat_prompt = chat_prompt.partial(slide_context=presentation.slide_contexts_str)

was_audio = 0
was_LLM_recognized = 0

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
    trimmed_messages = all_messages[-settings.history_size:]
    return {"messages": trimmed_messages}


workflow = (StateGraph(PresentationState)
            .add_node("model", model_node)
            .add_edge(START, "model")
            .add_edge("model", END)
            )

memory = MemorySaver()
model_with_history = workflow.compile(checkpointer=memory)
config = {"configurable": {"thread_id": "unique"}}

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

    recognized_text = Groq().audio.transcriptions.create(
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
        state = PresentationState(
            messages=[],
            recognized_speech=recognized_speech,
            slide_now=presentation.history_slide.val
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
    tool_name = tool_call["name"]
    tool_args = tool_call["args"]
    print(f"tool: {tool_name}, args: {tool_args}")
    for tool in tools:
        if tool.name == tool_name:
            tool.invoke(input=tool_args)
            slide_now = presentation.slideshow.View.Slide.SlideIndex
            if slide_now != presentation.history_slide.val:
                presentation.history_slide = LinkedList(val=slide_now, prev=presentation.history_slide)
                print(f"В историю добавлен {presentation.history_slide.val}, "
                      f"предыдущий: {presentation.history_slide.prev.val}")
            break

def start_audio():
    with sd.InputStream(
        samplerate=settings.sample_rate,
        blocksize=frame_size,
        channels=1,
        callback=audio_callback,
        dtype="float32"
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
    presentation.run_powerPoint()
    start_audio()
finally:
    presentation.exit_pp()
