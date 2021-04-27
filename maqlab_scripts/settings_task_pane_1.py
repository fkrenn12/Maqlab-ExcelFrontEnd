"""
PyXLL Examples: Tk Custom Task Pane

This example shows how a Tkinter window can be hosted
in an Excel Custom Task Pane.

Custom Task Panes are Excel controls that can be docked in the
Excel application or be floating windows. These can be used for
creating more advanced UI Python tools within Excel.
"""
import asyncio

from pyxll import create_ctp, CTPDockPositionRight, xl_app
import tkinter as tk
import tkinter.messagebox as messagebox
import logging
import os
import configparser
import paho.mqtt.client as paho
import datetime
import time

_log = logging.getLogger(__name__)
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

try:
    # pip install pillow
    from PIL import Image, ImageTk
except ImportError:
    Image = ImageTk = None
    _log.warn("PIL not installed. Use 'pip install pillow' to install.")

client = paho.Client()
# Declare event loop
loop = None
event_connect = asyncio.Event()
stop = False
state = None
config_path = os.path.dirname(os.path.abspath(__file__)) + "/maqlab.cfg"


class MQTTSettingsFrame(tk.Frame):

    def __init__(self, master):

        tk.Frame.__init__(self, master)
        self.__config = configparser.ConfigParser()
        # self.__config_path = os.path.dirname(os.path.abspath(__file__)) + "/maqlab.cfg"
        self.__config.read(config_path)
        self.__hostname = "mqtt.com"
        self.__port = 1883
        self.__user = "user"
        self.__pass = "pass"
        try:
            self.__hostname = self.__config.get("MQTT", "hostname")
            self.__port = self.__config.get("MQTT", "port")
            self.__user = self.__config.get("MQTT", "user")
            self.__pass = self.__config.get("MQTT", "pass")
        except:
            pass
        self.init_window()

    def init_window(self):
        global state
        state = tk.StringVar()
        # allow the widget to take the full space of the root window
        self.pack(fill=tk.BOTH, expand=True)

        row = 0

        # caption
        label = tk.Label(self, text="Hostname or IP-address")
        label.grid(column=0, row=row, padx=0, pady=0, sticky="e")

        self.entry_hostname = tk.Entry(self)
        self.entry_hostname.delete(0)
        self.entry_hostname.insert(0, self.__hostname)
        self.entry_hostname.grid(column=1, row=row, padx=(2, 10), pady=3, ipadx=0, sticky="nsew")

        row += 1
        label = tk.Label(self, text="Port")
        label.grid(column=0, row=row, padx=0, pady=0, sticky="e")

        self.entry_port = tk.Entry(self)
        self.entry_port.delete(0)
        self.entry_port.insert(0, self.__port)
        self.entry_port.grid(column=1, row=row, padx=(2, 10), pady=(0, 3), ipady=0, sticky="nsew")

        row += 1
        label = tk.Label(self, text="Username")
        label.grid(column=0, row=row, padx=0, pady=0, sticky="e")

        self.entry_user = tk.Entry(self)
        self.entry_user.delete(0)
        self.entry_user.insert(0, self.__user)
        self.entry_user.grid(column=1, row=row, padx=(2, 10), pady=(0, 3), ipady=0, sticky="nsew")

        row += 1
        label = tk.Label(self, text="Password")
        label.grid(column=0, row=row, padx=0, pady=0, sticky="e")

        self.entry_pass = tk.Entry(self)
        self.entry_pass.delete(0)
        self.entry_pass.insert(0, self.__pass)
        self.entry_pass.grid(column=1, row=row, padx=(2, 10), pady=(0, 3), ipady=0, sticky="nsew")
        # Allow the last grid column to stretch horizontally

        row += 1
        self.button_check = tk.Button(self, text="Check connection")
        self.button_check.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")
        self.button_check.bind("<Button-1>", self.check)

        row += 1
        self.button_save = tk.Button(self, text="Save settings", bg='#40E0D0')
        self.button_save.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")
        self.button_save.bind("<Button-1>", self.save)

        # row += 1
        # self.button_close = tk.Button(self, text="Close this panel", bg='#40E0D0')
        # self.button_close.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")
        # self.button_close.bind("<Button-1>", self.close)

        row += 1
        state.set("State")
        self.label_state = tk.Label(self, textvar=state)
        self.label_state.grid(column=0, row=row, columnspan=2, padx=(20, 20), pady=10, sticky="nsew")

        self.columnconfigure(11, weight=1)

    def check(self, event):
        global client
        try:
            event_connect.set()
            # check inputs
            state.set("Connecting to " + self.entry_hostname.get() + ":" + self.entry_port.get() + "...")
        except:
            # client.loop_stop()
            state.set(
                str(datetime.datetime.now()) + "  :" + "MAQlab - Connection Error! Are you connected to the internet?")

    def save(self, event):
        try:
            self.__config.set("MQTT", "hostname", self.entry_hostname.get())
            self.__config.set("MQTT", "port", self.entry_port.get())
            self.__config.set("MQTT", "user", self.entry_user.get())
            self.__config.set("MQTT", "pass", self.entry_pass.get())
            with open(config_path, 'w') as configfile:
                self.__config.write(configfile)
            state.set("Settings saved sucessfully")
        except:
            raise

    def __get_hostname(self):
        return self.entry_hostname.get()

    def __get_port(self):
        return int(self.entry_port.get())

    def __get_user(self):
        return self.entry_user.get()

    def __get_pass(self):
        return self.entry_pass.get()

    hostname = property(__get_hostname)
    port = property(__get_port)
    user = property(__get_user)
    passw = property(__get_pass)


async def connector():
    global client
    global stop
    while True:
        await event_connect.wait()
        event_connect.clear()

        if stop:
            break

        client.reconnect_delay_set(min_delay=1, max_delay=5)
        client.username_pw_set(username=tkf.user, password=tkf.passw)
        state.set(tkf.hostname + ":" + str(tkf.port))
        client.connect_async(host=tkf.hostname, port=tkf.port)
        client.loop_start()
        await asyncio.sleep(1)
        count = 1
        while not client.is_connected():
            await asyncio.sleep(0.5)
            state.set("Attempt#" + str(count))
            count += 1
            if stop:
                break
            if count > 10:
                break
        if client.is_connected():
            print("Connected")
            state.set("Connected")
        else:
            state.set("NOT connected ! Check settings !")
        try:
            client.disconnect()
        except:
            pass
        try:
            client.loop_stop()
        except:
            pass

        event_connect.clear()


async def mqtt_loop():
    global client
    while True:
        try:
            pass
            # client.loop(0.001)
        except:
            pass
        await asyncio.sleep(0.001)
        if stop:
            break


async def tk_loop(w):
    while True:
        try:
            w.update_idletasks()
            w.update()
        except:
            break
        await asyncio.sleep(0.001)
        if stop:
            break


async def main(w, loop):
    global client
    task1 = loop.create_task(tk_loop(w))
    task2 = loop.create_task(mqtt_loop())
    task3 = loop.create_task(connector())
    await asyncio.wait([task3])
    await asyncio.wait([task1])
    await asyncio.wait([task2])
    # loop.stop()

    w.destroy()
    # exit(0)


def on_closing():
    global stop
    stop = True
    event_connect.set()


def show_tk_ctp():
    global stop
    global loop
    global tkf
    try:
        if loop is None:
            loop = asyncio.get_event_loop()
    except:
        pass
    event_connect.clear()
    stop = False
    """Create a Tk window and embed it in Excel as a Custom Task Pane."""
    # reading the settings in the configuration file
    # Create the top level Tk window
    # window = tk.Toplevel()
    window = tk.Tk()
    window.title("MQTT-Broker settings")
    # window.overrideredirect = False
    window.protocol("WM_DELETE_WINDOW", on_closing)
    # Add our example frame to it
    tkf = MQTTSettingsFrame(master=window)

    # Add the widget to Excel as a Custom Task Pane

    # create_ctp(window, width=300, position=CTPDockPositionRight)
    # Run the code until completing all task
    # if not loop.is_running():
    loop.run_until_complete(main(window, loop))
    # loop.close()

    stop = False
    # Close the loop
    # loop.close()
    # while True:
    #    window.update_idletasks()
    #    window.update()
    #    print("*")


if __name__ == "__main__":
    loop = asyncio.get_event_loop()
    root = tk.Tk()
    root.title("MQTT-Broker settings")
    root.protocol("WM_DELETE_WINDOW", on_closing)
    tkf = MQTTSettingsFrame(master=root)
    loop.run_until_complete(main(root, loop))
    # Close the loop
    loop.close()

    # while True:
    #    root.update_idletasks()
    #    root.update()
    #    print("*")
