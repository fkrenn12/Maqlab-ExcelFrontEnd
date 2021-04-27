print("Module 'maqlab_scripts' __init__.py called.")
import time
import paho.mqtt.client as paho
import datetime
import queue
import threading
import secrets
import sys
import os
import json
import ast
import tkinter
import maqlab_scripts.maqlab as maqlab
# Adding the ../MAQLab/.. folder to the system path of python
# It is temporarily used by this script only
script_dir = str()
try:
    script_dir = os.path.dirname(__file__)
    maqlab_dir = "\\maqlab_scripts"
    script_dir = script_dir[0:script_dir.index(maqlab_dir)] + maqlab_dir
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)
except:
    pass

# checking the config file
# we have to adjust the path
# this is not the best solution but it works so far
# it would be better to search after the config file in the file path
# TODO: search for  the file in the path downwards
maqlab_config = script_dir + '\\maqlab.cfg'
if not os.path.exists(maqlab_config):
    # maqlab_config = script_dir + '\\maqlab\\maqlab.cfg'
    maqlab_config = script_dir + '\\maqlab.cfg'

# default broker credentials
mqtt_hostname = "mqtt.techfit.at"
mqtt_port = 1883
mqtt_user = "maqlab"
mqtt_pass = "maqlab"

# reading the configuration file and set the values
config = str()
try:
    with open(maqlab_config, 'r') as file:
        config = file.read().split("\n")
        for line in config:
            if line.startswith("{") and line.endswith("}"):
                js = json.loads(line)
                try:
                    mqtt_hostname = js["mqtt_hostname"]
                    continue
                except:
                    pass
                try:
                    mqtt_port = js["mqtt_port"]
                    continue
                except:
                    pass
                try:
                    mqtt_user = js["mqtt_user"]
                    continue
                except:
                    pass
                try:
                    mqtt_pass = js["mqtt_password"]
                    continue
                except:
                    pass

except:
    print(str(
        datetime.datetime.now()) + "  :MAQLAB - Could not open file " + maqlab_config + " or data in file is corrupted")
    raise

print(str(datetime.datetime.now()) + "  :MAQLAB - Configuration loaded successfully")


class MqttMsg:
    def __init__(self, topic, payload=""):
        self.topic = topic
        self.payload = payload


class MAQLabError(Exception):
    pass


class MAQLab(Exception):
    pass


# --------------------------------------------------------------------------------
# Class                             M A Q L A B
# --------------------------------------------------------------------------------
class MAQLab:

    def __init__(self, host, port, user, password, session_id, stamp=""):
        try:
            self.__q_out = queue.Queue()
            print(str(datetime.datetime.now()) + "  :" + "MQTT - started")
            self.__static_stamp = stamp
            self.__mqtt_hostname = str(host)
            self.__mqtt_port = int(port)
            self.__mqtt_user = str(user)
            self.__mqtt_pass = str(password)
            self.__session_id = session_id
            self.__device_commands = list()
            self.__device_types = list()
            self.__device_models = list()
            self.__device_manufactorers = list()
            self.__device_accessnumbers = list()
            self.__lock = threading.Lock()
            self.__client = paho.Client()
            self.__client.on_connect = self.__on_connect
            self.__client.on_disconnect = self.__on_disconnect
            self.__client.on_message = self.__on_message
            self.__client.reconnect_delay_set(min_delay=1, max_delay=5)
            self.__client.username_pw_set(self.__mqtt_user, self.__mqtt_pass)
            self.__client.connect(self.__mqtt_hostname, self.__mqtt_port)
            self.__client.loop_start()
            attemptions = 1
            while not self.__client.is_connected():
                print(str(datetime.datetime.now()) + "  :MQTT - connecting...attempt#" + str(attemptions))
                time.sleep(1)
                attemptions += 1

            print(str(datetime.datetime.now()) + "  :" + "MQTT - ready")
        except Exception as _e:
            # print(_e)
            print(
                str(datetime.datetime.now()) + "  :" + "MAQlab - Connection Error! Are you connected to the internet?")
            raise _e

    # --------------------------------------------------------
    # MQTT Broker callback on_connect
    # --------------------------------------------------------
    def __on_connect(self, _client, userdata, flags, rc):
        if rc == 0:
            print(str(datetime.datetime.now()) + "  :" + "MQTT - connected.")
            self.__client.subscribe("maqlab/" + str(self.__session_id) + "/rep/#", qos=0)
            self.__client.subscribe("maqlab/" + str(self.__session_id) + "/+/rep/#", qos=0)
            print(str(datetime.datetime.now()) + "  :" + "MQTT - Subscriptions done.")

    # ------------------------------------------------------------------------------
    # MQTT Broker callback on_disconnect
    # ------------------------------------------------------------------------------
    def __on_disconnect(self, _client, userdata, rc):
        if rc != 0:
            print(str(datetime.datetime.now()) + "  :" + "Unexpected MQTT-Broker disconnection.")

    # ------------------------------------------------------------------------------
    # MQTT Broker callback on_message
    # ------------------------------------------------------------------------------
    def __on_message(self, _client, _userdata, _msg):
        # check topic
        try:
            topic = _msg.topic.decode("utf-8")
        except:
            try:
                topic = _msg.topic.replace(" ", " ")
            except:
                return
        # check payload
        try:
            payload = _msg.payload.decode("utf-8")
        except:
            try:
                payload = _msg.payload.replace(" ", " ")
            except:
                return

        # print(_msg.topic, _msg.payload)
        # on_message is called from an other thread and therefore
        # the object _msg could be manipulated immediately
        # after putting it on the queue before it is handled
        # from the following stage.
        # The solution is to send topic and payload as string rather than as object
        self.__q_out.put(str([topic, payload]), block=False, timeout=0)

    # ------------------------------------------------------------------------------
    #  Flush the queue
    # ------------------------------------------------------------------------------
    def __flush(self, block=False, timeout=0):
        while True:
            try:
                if self.__q_out.empty():
                    return
                else:
                    self.__q_out.get(block=block, timeout=timeout)
            except:
                return

    # ------------------------------------------------------------------------------
    #  Send ( internal used )
    # ------------------------------------------------------------------------------
    def __send(self, msg, stamp="_"):
        try:
            self.__flush()
            self.__client.publish("maqlab/" + self.__session_id + "/" + stamp + "/cmd" + msg.topic, msg.payload)
        except:
            raise MAQLabError("Send error")

    # ------------------------------------------------------------------------------
    #  Receive ( internal used )
    # ------------------------------------------------------------------------------
    def __receive(self, block=True, timeout=1.0, stamp="_"):
        try:
            rec_msg = self.__q_out.get(block=block, timeout=timeout)
            try:
                rec_msg = rec_msg.decode("utf-8")
            except:
                pass
            # eval to list object
            rec_msg = ast.literal_eval(rec_msg)
            try:
                msg = MqttMsg(topic=rec_msg[0], payload=rec_msg[1])
                if stamp in msg.topic.split("/"):
                    return msg
                else:
                    raise MAQLabError("Wrong message stamp - message discarded ")
            except:
                raise
        except:
            if block:
                raise MAQLabError("Timeout error")
            else:
                raise MAQLab("Empty")

    # ------------------------------------------------------------------------------
    # Send a message and wait for the answer
    # ------------------------------------------------------------------------------
    def send_and_receive(self, receive=True, accessnumber=None, command="", value="", msg=None, block=True,
                         timeout=1.0):
        try:
            value = str(value)
        except:
            value = ""

        if msg is None:
            if not command.startswith("/"):
                command = "/" + command
            if accessnumber is not None:
                command = "/" + str(accessnumber) + command
            msg = MqttMsg(command, value)
        else:
            if not msg.topic.startswith("/"):
                msg.topic = "/" + msg.topic
            if accessnumber is not None:
                msg.topic = "/" + str(accessnumber) + msg.topic

        with self.__lock:
            try:
                self.__flush()
                if not receive:
                    self.__send(msg=msg)
                else:
                    stamp = self.__static_stamp
                    if stamp == "":
                        stamp = str(int((time.time() * 1000) % 1000000))
                    self.__send(msg=msg, stamp=stamp)
                if receive:
                    return self.__receive(block=block, timeout=timeout, stamp=stamp)
            except Exception as _e:
                raise _e

    # ------------------------------------------------------------------------------
    # Send a message and returns a list of all answers
    # ------------------------------------------------------------------------------
    def send_and_receive_burst(self, accessnumber=None, command="", value="", msg=None, block=True, timeout=1.0,
                               burst_timout=1.0):
        try:
            value = str(value)
        except:
            value = ""

        if msg is None:
            if not command.startswith("/"):
                command = "/" + command
            if accessnumber is not None:
                command = "/" + str(accessnumber) + command
            msg = MqttMsg(command, value)
        else:
            if not msg.topic.startswith("/"):
                msg.topic = "/" + msg.topic
            if accessnumber is not None:
                msg.topic = "/" + str(accessnumber) + msg.topic

        _timeout = timeout
        msg_list = list()
        with self.__lock:
            try:
                stamp = self.__static_stamp
                if stamp == "":
                    stamp = str(int((time.time() * 1000) % 1000000))
                self.__flush()
                self.__send(msg=msg, stamp=stamp)
                while True:
                    try:
                        msg_received = self.__q_out.get(block=block, timeout=_timeout)
                        try:
                            msg_received = msg_received.decode("utf-8")
                        except:
                            pass
                        msg_received = ast.literal_eval(msg_received)
                        msg = MqttMsg(topic=msg_received[0], payload=msg_received[1])
                        if stamp in msg.topic:
                            _timeout = burst_timout
                            msg_list.append(msg)
                    except:
                        if len(msg_list) == 0:
                            raise MAQLabError("Empty data")
                        return msg_list
            except Exception as _e:
                raise _e

    # ------------------------------------------------------------------------------
    #
    # ------------------------------------------------------------------------------
    def load_devices(self):
        try:
            detected_devices_raw = self.send_and_receive_burst(command="/?", burst_timout=0.5)
            try:
                detected_devices = []
                for item in detected_devices_raw:
                    try:
                        if "accessnumber" in item.topic:
                            topic_splitted = item.topic.split("/")
                            devicename = topic_splitted[topic_splitted.index("rep") + 1]
                            accessnumber = int(item.payload)
                            detected_devices.append(tuple((devicename, accessnumber)))
                    except:
                        raise
            except:
                raise
            # we have the list of device
            # next task is to request the details of the device
            # reading the available commands and manufactor from each device
            # lets clear the actual list
            self.__device_commands.clear()
            self.__device_types.clear()
            self.__device_models.clear()
            self.__device_manufactorers.clear()

            for device in detected_devices:
                print(str(datetime.datetime.now()) + "  :" + "MAQlab - Detected: " + str(
                    device[0]) + " Accessnumber is " + str(device[1]))
                try:
                    number = int(device[1])
                except:
                    raise
                self.__device_accessnumbers.append(number)
                reps = maqlab.send_and_receive_burst(command=str(device[1]) + "/?", burst_timout=0.5)
                for rep in reps:
                    if "commands" in rep.topic:
                        self.__device_commands.append(rep.payload)
                    elif "manufactorer" in rep.topic:
                        self.__device_manufactorers.append(rep.payload)
                    elif "model" in rep.topic:
                        self.__device_models.append(rep.payload)
                    elif "devicetype" in rep.topic:
                        self.__device_types.append(rep.payload)

            # print(self.__device_models)
            # print(self.__device_accessnumbers)
            # print(self.__device_manufactorers)
            # print(self.__device_commands)
            # print(self.__device_types)

        except:
            raise
        return detected_devices

    # ------------------------------------------------------------------------------
    #
    # ------------------------------------------------------------------------------
    def close(self):
        try:
            self.__client.on_disconnect = None
            self.__client.on_connect = None
            self.__client.disconnect()
        except:
            pass

    def __get_model(self):
        return self.__device_models

    def __get_accessnumbers(self):
        return self.__device_accessnumbers

    def __get_commands(self):
        return self.__device_commands

    def __get_manufactorers(self):
        return self.__device_manufactorers

    def __get_types(self):
        return self.__device_types

    def __isconnected(self):
        return self.__client.is_connected()

    # ------------------------------------------------------------------------------
    #
    # ------------------------------------------------------------------------------
    def __str__(self):
        return self.__session_id

    device_models = property(__get_model)
    device_accessnumbers = property(__get_accessnumbers)
    device_types = property(__get_types)
    device_manufactorers = property(__get_manufactorers)
    device_commands = property(__get_commands)
    is_connected = property(__isconnected)


try:
    maqlab = MAQLab(host=mqtt_hostname,
                    port=mqtt_port,
                    user=mqtt_user,
                    password=mqtt_pass,
                    session_id=secrets.token_urlsafe(3).lower())
    maqlab.load_devices()
    # window = tkinter.Tk()
    # window.title = "MAQLAB device Monitor"

except Exception as e:
    maqlab = None
