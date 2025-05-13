from langchain_core.tools import tool
from PresentationSettings import presentation

@tool
def next_slide():
    """Move to the next slide"""
    if presentation.count_slides > presentation.slideshow.View.Slide.SlideIndex:
        print("Переход на следующий слайд.")
        presentation.slideshow.View.Next()
    else:
        print("Это последний слайд.")


@tool
def prev_slide():
    """Move to 1 slide back"""
    if presentation.slideshow.View.Slide.SlideIndex > 1:
        print("Переход на предыдущий слайд.")
        presentation.slideshow.View.Previous()
    else:
        print("Это первый слайд.")


@tool
def move_to_slide(slide_number: str):
    """Move to n-th slide"""
    try:
        n = int(slide_number)
        if 1 <= n <= presentation.count_slides:
            presentation.slide_number = n
            print(f"Переход к {slide_number} слайду.")
            presentation.slideshow.View.GotoSlide(n)
        else:
            print(f"Недопустимый номер слайда {n}.")
    except ValueError:
        print("Недопустимый номер слайда. Ожидалось число.")


@tool
def close_slideshow():
    """Close slideshow"""
    if presentation.slideshow:
        print("Презентация закрыта.")
        presentation.slideshow.View.Exit()

@tool
def back_slide():
    """Return to slide what was before this"""
    if presentation.history_slide.prev:
        presentation.history_slide = presentation.history_slide.prev
    print(f"Возвращение обратно на {presentation.history_slide.val} слайд.")
    move_to_slide.invoke(str(presentation.history_slide.val))


@tool
def no_move():
    """Doing nothing if not need to do something"""
    print("Ничего не делать. Выполнено!")

tools = [next_slide, prev_slide, move_to_slide, close_slideshow, back_slide, no_move]
