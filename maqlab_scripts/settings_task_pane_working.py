"""
PyXLL Examples: Tk Custom Task Pane

This example shows how a Tkinter window can be hosted
in an Excel Custom Task Pane.

Custom Task Panes are Excel controls that can be docked in the
Excel application or be floating windows. These can be used for
creating more advanced UI Python tools within Excel.
"""
import asyncio
from pyxll import create_ctp, CTPDockPositionRight, get_event_loop
import tkinter as tk
import tkinter.messagebox as messagebox
import logging
import os
import configparser
import paho.mqtt.client as paho
import datetime
import threading
import time
import pyxll
from contextlib import suppress
import wx
from wxasync import WxAsyncApp

stop = False
run = False
from pyxll import xl_on_close

_log = logging.getLogger(__name__)
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

mutex = threading.Lock()

event_connect = asyncio.Event()
config_path = os.path.dirname(os.path.abspath(__file__)) + "/maqlab.cfg"


class WxMQTTSettingsPanel(wx.Panel):

    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)
        self.Bind(wx.EVT_WINDOW_DESTROY, self.on_close)

        self.__config = configparser.ConfigParser()
        self.__config.read(config_path)
        self._hostname = wx.TextCtrl(self)
        self._port = wx.TextCtrl(self)
        self._user = wx.TextCtrl(self)
        self._password = wx.TextCtrl(self)

        self._hostname.SetValue("mqtt.com")
        self._port.SetValue(str(1883))
        self._user.SetValue("user")
        self._password.SetValue("pass")

        try:
            self._hostname.SetValue(self.__config.get("MQTT", "hostname"))
            self._port.SetValue(self.__config.get("MQTT", "port"))
            self._user.SetValue(self.__config.get("MQTT", "user"))
            self._password.SetValue(self.__config.get("MQTT", "pass"))
        except:
            raise

        label_port = wx.StaticText(self, label="Port")
        label_hostname = wx.StaticText(self, label="Hostname")
        label_user = wx.StaticText(self, label="Username")
        label_password = wx.StaticText(self, label="Password")
        self.label_status = wx.StaticText(self, label="")
        label_dummy = wx.StaticText(self, label="")

        button_check = wx.Button(self, label="Check connection")
        button_check.Bind(wx.EVT_BUTTON, self.check)
        button_save = wx.Button(self, label="Save settings")
        button_save.Bind(wx.EVT_BUTTON, self.save)

        sizer = wx.BoxSizer(orient=wx.VERTICAL)

        fgs = wx.FlexGridSizer(4, 2, 9, 20)
        fgs1 = wx.FlexGridSizer(2, 2, 9, 2)

        fgs.AddMany([label_hostname, (self._hostname, 1, wx.EXPAND), label_port,
                     (self._port, 1, wx.EXPAND), (label_user, 1, wx.EXPAND), (self._user, 1, wx.EXPAND),
                     label_password, (self._password, 1, wx.EXPAND)])

        fgs1.AddMany([(self.label_status, 2, wx.EXPAND), label_dummy , button_check, button_save ])
        sizer.Add(fgs, proportion=1, flag=wx.ALL | wx.EXPAND, border=15)
        sizer.Add(fgs1, proportion=1, flag=wx.ALL | wx.EXPAND, border=15)
        self.SetSizer(sizer)
        self.Layout()

    def check(self, event):
        mqtt.connect = True

    def save(self, event):
        try:
            self.__config.set("MQTT", "hostname", self._hostname.GetValue())
            self.__config.set("MQTT", "port", self._port.GetValue())
            self.__config.set("MQTT", "user", self._user.GetValue())
            self.__config.set("MQTT", "pass", self._password.GetValue())
            with open(config_path, 'w') as configfile:
                self.__config.write(configfile)
            # state.set("Settings saved sucessfully")
        except:
            raise

    def on_close(self, event):
        try:
            mutex.release()
            mqtt.connect = False
        except:
            pass
        # self.Destroy()  # you may also do:  event.Skip()
        # since the default event handler does call Destroy(), too

    def __get_hostname(self):
        return self._hostname.GetValue()

    def __get_port(self):
        return int(self._port.GetValue())

    def __get_user(self):
        return self._user.GetValue()

    def __get_pass(self):
        return self._password.GetValue()

    def __get_status(self):
        return self.label_status.GetLabel()

    def __set_status(self, value):
        self.label_status.SetLabel(value)
        # if "Connected!" in value:
        #     self.label_status.SetBackgroundColour(0x121212)

    hostname = property(__get_hostname)
    port = property(__get_port)
    user = property(__get_user)
    passw = property(__get_pass)
    status = property(__get_status, __set_status)


class Mqtt:
    def __init__(self):
        if __name__ == "__main__":
            self.__loop = asyncio.get_event_loop()
        else:
            self.__loop = get_event_loop()
        self.__event_connect = asyncio.Event()
        self.__window = None
        self.__frame = None
        self.__task1 = None
        self.__task2 = None

    async def connector(self):
        global run
        global stop
        run = False
        while True:
            stop = False
            self.__event_connect.clear()
            while not run:
                await asyncio.sleep(0.1)
            # await self.__event_connect.wait()
            run = False
            self.__frame.status = "Connecting..."
            client = paho.Client()
            client.reconnect_delay_set(min_delay=1, max_delay=5)
            try:
                client.username_pw_set(username=self.__frame.user, password=self.__frame.passw)
                client.connect_async(host=self.__frame.hostname, port=self.__frame.port)
                client.loop_start()
            except:
                pass
            await asyncio.sleep(0.5)
            count = 1
            while not client.is_connected():
                if count > 10:
                    break
                await asyncio.sleep(0.25)
                self.__frame.status = "Attempt#" + str(count)
                # state.set("Attempt#" + str(count))
                count += 1
                if count > 10:
                    break
                await asyncio.sleep(0.25)
            if client.is_connected():
                # wx.MessageBox("Connected", style=wx.ICON_INFORMATION)
                self.__frame.status = "Connected!"
            else:
                # wx.MessageBox("Not Connected", style=wx.ICON_INFORMATION)
                self.__frame.status = "Not connected! Check settings and/or network"
                # state.set("NOT connected ! Check settings !")
            try:
                del client
            except:
                pass

    async def main(self):
        self.__task2 = self.__loop.create_task(self.connector())
        await asyncio.wait([self.__task2])

    def __on_closing(self, event):
        self.__task2.cancel()
        try:
            mutex.release()
        except:
            pass

    def __set_window(self, new_window):
        self.__window = new_window
        # self.__window.bind("<Destroy>", self.__on_closing)

    def __set_frame(self, new_frame):
        self.__frame = new_frame
        # self.__frame.bind("<Destroy>", self.__on_closing)

    def __get_window(self):
        return self.__window

    def __get_frame(self):
        return self.__frame

    def __get_loop(self):
        return self.__loop

    def __get_start(self):
        return None

    def __set_start(self, value):
        global run
        if value:
            run = True
            self.__event_connect.set()
        else:
            self.__task2.cancel()

    window = property(__get_window, __set_window)
    frame = property(__get_frame, __set_frame)
    loop = property(__get_loop)
    connect = property(__get_start, __set_start)


mqtt = Mqtt()


# @xl_on_close
# def on_close():
#    wx.MessageBox("Host", "CLOSE", style=wx.ICON_INFORMATION)


class WxCustomTaskPane(wx.Frame):

    def __init__(self, name):
        wx.Frame.__init__(self, parent=None)
        self.SetTitle(name)
        self.control = WxMQTTSettingsPanel(parent=self)


def show_wx_ctp():
    global mqtt
    global frame
    global app

    if not mutex.locked():
        mutex.acquire()
        # Make sure the wx App has been created
        app = wx.App.Get()
        if app is None:
            app = wx.App()
        # Create the frame to use as the Custom Task Pane
        frame = WxCustomTaskPane("MQTT-Broker settings")
        mqtt.frame = frame.control
        # Add the frame to Excel
        create_ctp(frame, width=300, height=500, position=CTPDockPositionRight)
        asyncio.run_coroutine_threadsafe(mqtt.main(), mqtt.loop)


if __name__ == "__main__":
    app = WxAsyncApp()
    loop = asyncio.events.get_event_loop()
    frame = WxCustomTaskPane("MQTT-Broker settings")
    mqtt.frame = frame.control
    frame.Show()
    loop.create_task(mqtt.main())
    loop.run_until_complete(app.MainLoop())

