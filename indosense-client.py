import serial, openpyxl, time

# Commands
SET_REG = b'02'
READ_REG = b'03'

# Registers
LHR_DATA_LSB = b'38'
LHR_DATA_MID = b'39'
LHR_DATA_MSB = b'3A'
ALT_CONFIG = b'05'
D_CONFIG = b'0C'
DIG_CONF = b'03'
START_CONFIG = b'0B'
LHR_CONFIG = b'34'
DEVICE_ID = b'3F'

default_com_no = 5
baudrate = 115200

port = serial.Serial('COM' + str(default_com_no), baudrate)
port.flush()

port.write(SET_REG + ALT_CONFIG + b'01' + b'\r') # Swith off Rp module
print(port.read(32)[8])

port.write(SET_REG + D_CONFIG + b'01' + b'\r') # Don't track amplitude
print(port.read(32)[8])

port.write(SET_REG + DIG_CONF + b'F0' + b'\r') # F_min = 8MHz
print(port.read(32)[8])

port.write(SET_REG + LHR_CONFIG + b'03' + b'\r') # F_div = 8
print(port.read(32)[8])

port.write(SET_REG + START_CONFIG + b'00' + b'\r') # Continuous conversion mode
print(port.read(32)[8])


while True:
    ldata = 0x0000
    port.write(READ_REG + LHR_DATA_LSB + b'\r')
    ldata = ldata + port.read(32)[8]
    port.write(READ_REG + LHR_DATA_MID + b'\r')
    ldata = ldata + port.read(32)[8] * 256
    port.write(READ_REG + LHR_DATA_MSB + b'\r')
    ldata = ldata + port.read(32)[8] * 256 * 256 
    print(ldata)
    time.sleep(0.05)

