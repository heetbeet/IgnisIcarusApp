import minimalmodbus
from serial import SerialException

dev = minimalmodbus.Instrument("COM4",1)
dev.serial.baudrate = 9600

try:
  dev.write_register(0x0BB8,12,functioncode=6)
  # read = dev.read_register(1)
  # print(read)
  
except SerialException as e:
  print(e)