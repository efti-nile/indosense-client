import serial, openpyxl, time, msvcrt

track_recording = True

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

def read_ldata():
    global port
    ldata = 0x0000
    port.write(READ_REG + LHR_DATA_LSB + b'\r')
    ldata = ldata + port.read(32)[8]
    port.write(READ_REG + LHR_DATA_MID + b'\r')
    ldata = ldata + port.read(32)[8] * 256
    port.write(READ_REG + LHR_DATA_MSB + b'\r')
    ldata = ldata + port.read(32)[8] * 256 * 256
    return ldata

# asks whether a key has been acquired
def kbfunc():
    #this is boolean for whether the keyboard has bene hit
    x = msvcrt.kbhit()
    if x:
        #getch acquires the character encoded in binary ASCII
        ret = msvcrt.getch()
    else:
        ret = False
    return ret

while track_recording:
	raw_data = []
	filename = input('Enter file name: ')
	input('Press enter to star...')
	print('Press \'s\' to stop')
	while True:
		raw_data.append(read_ldata())
		time.sleep(0.05)
		key = kbfunc()
		if key != False:
			if key == b's':
				break;
	wb = openpyxl.Workbook()
	ws = wb.active
	for i in range(len(raw_data)):
		ws.append([raw_data[i]])
	wb.save(filename + '.xlsx')
