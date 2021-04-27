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
import logging
import os
import configparser
import paho.mqtt.client as paho
import threading
import wx
from wxasync import WxAsyncApp
from pubsub import pub
import subpub as loc_mqtt

_log = logging.getLogger(__name__)
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

mutex = threading.Lock()
config_path = os.path.dirname(os.path.abspath(__file__)) + "/maqlab.cfg"

MQTT_STATUS_RESET = -1
MQTT_STATUS_WAIT = 0
MQTT_STATUS_CONNECTED = 1
MQTT_STATUS_CONNECTING = 2
MQTT_STATUS_ERROR_CONNECTING = 3
MQTT_STATUS_BG_CONNECTED = wx.Colour(0, 0xff, 0x20)
MQTT_STATUS_BG_ERROR_CONNECTING = "red"
MQTT_STATUS_BG_CONNECTING = "yellow"


class Controller:

    def __init__(self):
        self.__loop = get_event_loop()
        self.__view = None
        self.__mqtt = None
        self.__mqtt_status_prev = MQTT_STATUS_RESET
        self.__mqtt_status_count = 1

    def __set_status(self, value):
        if type(value) is str:
            self.__view.status = value
        else:
            if value != self.__mqtt_status_prev:
                # status has changed
                if self.__mqtt_status_prev == MQTT_STATUS_RESET and value == MQTT_STATUS_WAIT:
                    self.__view.statues = ""

                if value == MQTT_STATUS_CONNECTING:
                    self.__view.status = "Connecting..."
                if value == MQTT_STATUS_CONNECTED:
                    self.__view.status = "Connection OK!"
                if value == MQTT_STATUS_ERROR_CONNECTING:
                    self.__view.status = "Not connected! Check settings and/or network"
                self.__mqtt_status_count = 1
            else:
                # status has not changed
                if value == MQTT_STATUS_CONNECTING:
                    self.__view.status = "Attempt#" + str(self.__mqtt_status_count)
                    self.__mqtt_status_count += 1
            self.__mqtt_status_prev = value

    def __set_hostname(self, value):
        self.__view.hostname = value

    def __set_port(self, value):
        self.__view.port = value

    def __set_username(self, value):
        self.__view.username = value

    def __set_password(self, value):
        self.__view.password = value

    def __get_hostname(self):
        return self.__view.hostname

    def __get_status(self):
        return self.__view.status

    def __get_port(self):
        return self.__view.port

    def __get_username(self):
        return self.__view.username

    def __get_password(self):
        return self.__view.password

    def __set_mqtt(self, value):
        self.__mqtt = value

    def __set_view(self, value):
        self.__view = value

    def __get_mqtt(self):
        return self.__mqtt

    def __get_view(self):
        return self.__view

    def __get_loop(self):
        return self.__loop

    def __get_connect_to_mqtt(self):
        return "TBD"

    def __set_connect_to_mqtt(self, value):
        self.__mqtt.connect = value

    loop = property(__get_loop)
    view = property(__get_view, __set_view)
    mqtt = property(__get_mqtt, __set_mqtt)
    hostname = property(__get_hostname, __set_hostname)
    port = property(__get_port, __set_port)
    username = property(__get_username, __set_username)
    password = property(__get_password, __set_password)
    mqtt_status = property(__get_status, __set_status)
    connect_to_mqtt = property(__get_connect_to_mqtt, __set_connect_to_mqtt)


# -------------------------------------------------------------------------------
#
# -------------------------------------------------------------------------------

class WxMQTTSettingsPanel(wx.Panel):

    def __init__(self, parent, controller_):
        wx.Panel.__init__(self, parent=parent)
        self.Bind(wx.EVT_WINDOW_DESTROY, self.on_close)
        self.__controller = controller_
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

        button_check = wx.Button(self, label="Try to connect")
        button_check.Bind(wx.EVT_BUTTON, self.check)
        button_save = wx.Button(self, label="Save settings")
        button_save.Bind(wx.EVT_BUTTON, self.save)

        sizer = wx.BoxSizer(orient=wx.VERTICAL)

        fgs = wx.FlexGridSizer(4, 2, 9, 20)
        fgs1 = wx.FlexGridSizer(2, 2, 9, 2)

        fgs.AddMany([label_hostname, (self._hostname, 1, wx.EXPAND), label_port,
                     (self._port, 1, wx.EXPAND), (label_user, 1, wx.EXPAND), (self._user, 1, wx.EXPAND),
                     label_password, (self._password, 1, wx.EXPAND)])

        fgs1.AddMany([(self.label_status, 2, wx.EXPAND), label_dummy, button_check, button_save])
        sizer.Add(fgs, proportion=1, flag=wx.ALL | wx.EXPAND, border=15)
        sizer.Add(fgs1, proportion=1, flag=wx.ALL | wx.EXPAND, border=15)
        self.SetSizer(sizer)
        self.Layout()

    def check(self, event):
        self.__controller.connect_to_mqtt = True

    def save(self, event):
        try:
            self.__config.set("MQTT", "hostname", self._hostname.GetValue())
            self.__config.set("MQTT", "port", self._port.GetValue())
            self.__config.set("MQTT", "user", self._user.GetValue())
            self.__config.set("MQTT", "pass", self._password.GetValue())
            with open(config_path, 'w') as configfile:
                self.__config.write(configfile)
            # state.set("Settings saved sucessfully")
            self.status = "Settings saved!"
        except:
            raise

    def on_close(self, event):
        try:
            mutex.release()
            self.__controller.connect_to_mqtt = False
        except:
            pass

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
        self.label_status.SetBackgroundColour(self.GetBackgroundColour())
        if "OK" in value:
            self.label_status.SetBackgroundColour(MQTT_STATUS_BG_CONNECTED)
        else:
            if "Attempt" in value:
                self.label_status.SetBackgroundColour(MQTT_STATUS_BG_CONNECTING)
            else:
                if "Not connected" in value:
                    self.label_status.SetBackgroundColour(MQTT_STATUS_BG_ERROR_CONNECTING)
        self.label_status.SetLabel(value)

    hostname = property(__get_hostname)
    port = property(__get_port)
    username = property(__get_user)
    password = property(__get_pass)
    status = property(__get_status, __set_status)


# -------------------------------------------------------------------------------
#
# -------------------------------------------------------------------------------
class Mqtt:
    def __init__(self, controller_):

        self.__event_connect = asyncio.Event()
        self.__task = None
        self.__controller = controller_
        self.__mqtt_loop = None
        self.__client = paho.Client()
        self.__client.reconnect_delay_set(min_delay=1, max_delay=5)

    async def connector(self):
        self.__controller.mqtt_status = MQTT_STATUS_RESET
        while True:
            self.__event_connect.clear()
            self.__controller.mqtt_status = MQTT_STATUS_WAIT
            while not self.__event_connect.is_set():
                await asyncio.sleep(0.01)

            self.__controller.mqtt_status = MQTT_STATUS_CONNECTING

            try:
                self.__client.username_pw_set(username=self.__controller.username, password=self.__controller.password)
                self.__client.connect_async(host=self.__controller.hostname, port=self.__controller.port)
                self.__client.loop_start()
            except:
                pass

            count = 1
            while not self.__client.is_connected():
                self.__controller.mqtt_status = MQTT_STATUS_CONNECTING  # "Attempt#" + str(count)
                if count > 10:
                    break
                await asyncio.sleep(0.25)
                count += 1
            if self.__client.is_connected():
                self.__controller.mqtt_status = MQTT_STATUS_CONNECTED

            else:
                # wx.MessageBox("Not Connected", style=wx.ICON_INFORMATION)
                self.__controller.mqtt_status = MQTT_STATUS_ERROR_CONNECTING  # "Not connected! Check settings and/or network"

            self.__client.loop_stop()
            self.__client.disconnect()

    async def main(self):
        self.__task = self.__controller.loop.create_task(self.connector())
        await asyncio.wait([self.__task])

    def __get_start(self):
        return None

    def __set_start(self, value):
        if value:
            self.__event_connect.set()
        else:
            self.__task.cancel()
            self.__mqtt_loop.cancel()

    connect = property(__get_start, __set_start)


# -------------------------------------------------------------------------------
#
# -------------------------------------------------------------------------------
class WxCustomTaskPane(wx.Frame):

    def __init__(self, name, controller_):
        wx.Frame.__init__(self, parent=None)
        self.SetTitle(name)
        self.control = WxMQTTSettingsPanel(self, controller_)


# -------------------------------------------------------------------------------
#
# -------------------------------------------------------------------------------
def show_wx_ctp():
    if not mutex.locked():
        mutex.acquire()
        # Make sure the wx App has been created
        app = wx.App.Get()
        if app is None:
            app = wx.App()
        # Create the frame to use as the Custom Task Pane
        controller = Controller()
        mqtt = Mqtt(controller)
        frame = WxCustomTaskPane("MQTT-Broker settings", controller)
        controller.mqtt = mqtt
        controller.view = frame.control
        # Add the frame to Excel
        create_ctp(frame, width=300, height=500, position=CTPDockPositionRight)
        asyncio.run_coroutine_threadsafe(mqtt.main(), controller.loop)


# -------------------------------------------------------------------------------
#
# -------------------------------------------------------------------------------
if __name__ == "__main__":
    # app = wx.App()
    app = WxAsyncApp()
    loop = asyncio.events.get_event_loop()
    controller = Controller()
    frame = WxCustomTaskPane("MQTT-Broker settings", controller)
    mqtt = Mqtt(controller)
    controller.mqtt = mqtt
    controller.view = frame.control
    frame.Show()
    loop.create_task(mqtt.main())
    loop.run_until_complete(app.MainLoop())
    # app.MainLoop()
