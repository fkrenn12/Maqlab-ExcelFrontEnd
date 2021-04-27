from datetime import date, timedelta, datetime
import threading
import logging
import time
import asyncio
from pyxll import xl_func, xl_app, RTD, Formatter
import maqlab_scripts


@xl_func
def devices_accesnumber():
    return maqlab_scripts.maqlab.device_accessnumbers

@xl_func
def devices_model():
    return maqlab_scripts.maqlab.device_models

@xl_func("int interval: string")
def set_throttle_interval(interval):
    xl = xl_app()
    xl.RTD.ThrottleInterval = interval
    return "OK"


_log = logging.getLogger(__name__)


class AsyncRTDExample(RTD):

    def __init__(self):
        super().__init__(value=0)
        self.__stopped = False

    async def connect(self):
        while not self.__stopped:
            # Yield to the event loop for 1s
            await asyncio.sleep(1)

            # Update value (which notifies Excel)
            self.value += 1

    async def disconnect(self):
        self.__stopped = True


@xl_func(": rtd<int>", recalc_on_open=True)
def async_rtd_example():
    return AsyncRTDExample()


class CurrentTimeRTD(RTD):
    """CurrentTimeRTD periodically updates its value with the current
    date and time. Whenever the value is updated Excel is notified and
    when Excel refreshes the new value will be displayed.
    """

    def __init__(self, format):
        initial_value = datetime.now().strftime(format)
        super(CurrentTimeRTD, self).__init__(value=initial_value)
        self.__format = format
        self.__running = True
        self.__thread = threading.Thread(target=self.__thread_func)
        self.__thread.start()

    def connect(self):
        # Called when Excel connects to this RTD instance, which occurs
        # shortly after an Excel function has returned an RTD object.
        _log.info("CurrentTimeRTD Connected")

    def disconnect(self):
        # Called when Excel no longer needs the RTD instance. This is
        # usually because there are no longer any cells that need it
        # or because Excel is shutting down.
        self.__running = False
        _log.info("CurrentTimeRTD Disconnected")

    def __thread_func(self):
        while self.__running:
            # Setting 'value' on an RTD instance triggers an update in Excel
            new_value = datetime.now().strftime(self.__format)
            if self.value != new_value:
                self.value = new_value
            time.sleep(0.2)


@xl_func("string format: rtd", recalc_on_open=True)
def rtd_current_time(format="%Y-%m-%d %H:%M:%S"):
    """Return the current time as 'real time data' that
    updates automatically.

    :param format: datetime format string
    """
    return CurrentTimeRTD(format)


@xl_func
def hello(name):
    return "Hello, %s" % name


@xl_func
def constrain(value: float):
    if value > 10:
        return 10
    else:
        return value


date_formatter = Formatter(text_color=0x553311, interior_color=0xa2c212)


@xl_func(formatter=date_formatter)
def add_days(d: date, i: int) -> date:
    return d + timedelta(days=i)


@xl_func("float[][] array: float[][] ")
def py_sum(array):
    """return the sum of a range of cells"""
    total = 0.0

    # 2d array is a list of lists of floats
    for row in array:
        for cell_value in row:
            total += cell_value

    # return total
    return array


class CustomObject:
    def __init__(self, name):
        self.name = name


@xl_func("string name: object")
def create_object(x):
    return CustomObject(x)


@xl_func("object x: string")
def get_object_name(x):
    assert isinstance(x, CustomObject)
    return x.name
