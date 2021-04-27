"""
PyXLL Examples: Tk Custom Task Pane

This example shows how a Tkinter window can be hosted
in an Excel Custom Task Pane.

Custom Task Panes are Excel controls that can be docked in the
Excel application or be floating windows. These can be used for
creating more advanced UI Python tools within Excel.
"""
import asyncio
from pyxll import create_ctp, CTPDockPositionRight, get_event_loop, xl_app
import tkinter as tk
import logging
import threading
from contextlib import suppress
from pyxll import xl_on_close
import pyxll

_log = logging.getLogger(__name__)
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

try:
    # pip install pillow
    from PIL import Image, ImageTk
except ImportError:
    Image = ImageTk = None
    _log.warn("PIL not installed. Use 'pip install pillow' to install.")

mutex = threading.Lock()


class Frame(tk.Frame):

    def __init__(self, master):

        tk.Frame.__init__(self, master)
        self.__state = tk.StringVar()
        self.init_window()

    def init_window(self):
        global counter

        # allow the widget to take the full space of the root window
        self.pack(fill=tk.BOTH, expand=True)

        row = 0

        # caption
        label = tk.Label(self, text="Demo with async loop")
        label.grid(column=0, row=row, padx=0, pady=0, sticky="e")

        row += 1
        self.button_check = tk.Button(self, text="Press button check")
        self.button_check.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")
        self.button_check.bind("<Button-1>", self.check)

        row += 1
        self.button_save = tk.Button(self, text="Press button save", bg='#40E0D0')
        self.button_save.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")
        self.button_save.bind("<Button-1>", self.save)

        # row += 1
        # self.button_close = tk.Button(self, text="Close this panel", bg='#40E0D0')
        # self.button_close.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")
        # self.button_close.bind("<Button-1>", self.close)

        row += 1
        self.__state.set("State")
        self.label_state = tk.Label(self, textvar=self.__state)
        self.label_state.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")

        row += 1
        counter = tk.StringVar()
        self.label_counter = tk.Label(self, textvar=counter)
        self.label_counter.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")

        self.columnconfigure(11, weight=1)

    def check(self, event):
        try:
            self.__state.set("We are checking")
        except:
            raise

    def save(self, event):
        try:
            self.__state.set("Saved sucessfully")
        except:
            raise


class Controller:
    def __init__(self):
        self.__loop = asyncio.get_event_loop()
        self.__window = None
        self.__frame = None
        self.__task1 = None
        self.__task2 = None

    async def task_counter(self):
        global counter
        val_counter = 0
        while True:
            val_counter += 1
            counter.set(str(val_counter))
            await asyncio.sleep(0.5)

    async def task_tk_loop(self):
        while True:
            try:
                self.__window.update_idletasks()
                self.__window.update()
            except:
                pass
            await asyncio.sleep(0.1)

    async def main(self):
        self.__task1 = self.__loop.create_task(self.task_tk_loop())
        self.__task2 = self.__loop.create_task(self.task_counter())
        await asyncio.wait([self.__task1])
        await asyncio.wait([self.__task2])

    def __on_closing(self, event):
        self.__task1.cancel()
        self.__task2.cancel()

    def __set_window(self, new_window):
        self.__window = new_window
        self.__window.bind("<Destroy>", self.__on_closing)

    def __set_frame(self, new_frame):
        self.__frame = new_frame
        self.__frame.bind("<Destroy>", self.__on_closing)

    def __get_window(self):
        return self.__window

    def __get_frame(self):
        return self.__frame

    def __get_loop(self):
        return self.__loop

    window = property(__get_window, __set_window)
    frame = property(__get_frame, __set_frame)
    loop = property(__get_loop)


controller = Controller()


@xl_on_close
def on_close():
    # Let's cancel all running tasks:

    pending = asyncio.Task.all_tasks()
    for task in pending:
        task.cancel()
        # Now we should await task to execute it's cancellation.
        # Cancelled task raises asyncio.CancelledError that we can suppress:
        #
        with suppress(asyncio.CancelledError):
            controller.loop.run_until_complete(task)


def show_tk_ctp():
    global window
    if not mutex.locked():
        mutex.acquire()
        """Create a Tk window and embed it in Excel as a Custom Task Pane."""
        # Create the top level Tk window
        window = tk.Toplevel()
        # window = tk.Tk()
        window.title("MQTT-Broker settings")
        controller.window = window
        controller.frame = Frame(master=window)
        # Add the widget to Excel as a Custom Task Pane

        create_ctp(window, width=300, position=CTPDockPositionRight)
        # pyxll.schedule_call(controller.main())
        controller.loop.run_until_complete(controller.main())
        mutex.release()


if __name__ == "__main__":
    window = tk.Tk()
    window.title("MQTT-Broker settings")
    controller.window = window
    controller.frame = Frame(master=window)
    controller.loop.run_until_complete(controller.main())
    controller.loop.close()
