import serial
import threading
import time

def is_float(s):
    """
    Test if a string can be parsed to a float
    >>> is_float("5")
    True
    >>> is_float("5.0")
    True
    >>> is_float("hello")
    False
    """
    try:
        float(s)
        return True
    except ValueError:
        return False


def _find_scale():
    readings = []
    comms = []
    for i in range(10):
        try:
            ser = serial.Serial(f'COM{i}',
                                baudrate=9600,
                                parity=serial.PARITY_NONE,
                                stopbits=serial.STOPBITS_ONE,
                                bytesize=serial.EIGHTBITS,
                                timeout=0.1)
            comms.append(i)

            #read three times to ensure a correct value
            readings.extend([ser.readline().decode('utf-8').strip() for i in range(3)][-1])
            try:
                val = float(readings[-1])
                #print(f"Found scale on COM{i}")
                return ser, val
            except:
                pass

        except serial.serialutil.SerialException:
            pass

    if not readings:
        if not comms:
            raise ConnectionError(f"Cannot make any RS232 connection to Preciso Scale Tronic.")
        else:
            raise ConnectionError(f"Cannot read values from Preciso Scale Tronic (tried COMs {comms}).")
    else:
        raise ConnectionError(f"Cannot read floating point value from Preciso Scale Tronic. Tried COMs {comms}, got values: {' '.join(readings)}")


class ScaleDevice:
    """
    This module reads values from a cheap Preciso scale through RS232. It opens a background thread to continuously
    reads values from the serial device, with the last value being accessible through the mass property. If the device
    reads 10 faulty values in a row the process crashes.
    """
    def __init__(self):
        self.ser, self._mass = _find_scale()

        self._health = True
        self._backlog = ['0']*10
        def background_reader():
            try:
                while(True):
                    s = self.ser.readline().decode('utf-8').strip()
                    self.ser.flushInput()
                    self._backlog = self._backlog[1:]+[s]

                    if len([i for i in self._backlog if is_float(i)]) == 0:
                        raise ValueError(f"Scale device is not recieving any float values: {self._backlog}")

                    if is_float(s):
                        self._mass = s
            finally:
                self._health = False

        self.thread = threading.Thread(target=background_reader)
        self.thread.start()


    @property
    def mass(self):
        if not self._health:
            raise ValueError("Device reading has stopped unexpectedly.")
        return self._mass


if __name__ == "__main__":

    # Example usage of scale device
    sd = ScaleDevice()
    for i in range(1000):
        time.sleep(0.1)
        print(sd.mass)