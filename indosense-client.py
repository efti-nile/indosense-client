import serial, openpyxl, time, msvcrt, os, collections

track_recording = True
monitoring = True

MATLAB_FOLDER = 'E:/work/indosens/'
MATLAB_SCRIPT_NAME = 'static_meas.m'

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
#print(port.read(32)[8])

port.write(SET_REG + D_CONFIG + b'01' + b'\r') # Don't track amplitude
#print(port.read(32)[8])

#port.write(SET_REG + DIG_CONF + b'F0' + b'\r') # F_min = 8MHz
#print(port.read(32)[8])

port.write(SET_REG + LHR_CONFIG + b'03' + b'\r') # F_div = 8
#print(port.read(32)[8])

port.write(SET_REG + START_CONFIG + b'00' + b'\r') # Continuous conversion mode
#print(port.read(32)[8])

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

while True:
	mode = input('Input \'m\' to enter monitor mod.\nInput \'r\' to enter monitor mod. \
	\nInput \'v\' to monitor standard deviation.\n')
	
	if mode == 'm':
		monitoring = True
		standev = False
		recording = False
		print('Press \'s\' to stop')
	elif mode == 'v':
		cb = collections.deque(maxlen = 20)
		monitoring = False
		standev = True
		recording = False
		print('Press \'s\' to stop')
	elif mode == 'r':
		monitoring = False
		standev = False
		recording = True
	else:
		continue
	
	while monitoring:	
		time.sleep(0.2)
		print('LHR_DATA = ' + str(read_ldata()))
		key = kbfunc()
		if key != False:
			if key == b's':
				break;

	while standev:
		time.sleep(0.2)
		cb.append(read_ldata())
		if len(cb) == 20:
			m = sum(cb) / 20
			sd = sum([pow(s - m, 2) for s in cb]) / 20
			print('STAND. DEV. = ' + str(sd))
		key = kbfunc()
		if key != False:
			if key == b's':
				break;
		
	
	if recording:
		raw_data = []
		throw_away_the_record = False
		filename = input('Enter output excel file name (without \'.xlsx\'): ')
		base_folder = os.path.join(MATLAB_FOLDER, filename)
		excel_file = os.path.join(base_folder, filename + '.xlsx')
		if os.path.exists(excel_file):
			ans = input('File already exists. Replace? (y/n) :')
			if ans == 'y':
				os.remove(excel_file)	
			else:
				continue
		input('Press enter to star...')
		print('Press \'s\' to stop\nPress \'d\' to throw away file')
		count = 0	
		port.flush()
		while True:
			raw_data.append(read_ldata())
			time.sleep(0.15)
			key = kbfunc()
			count = count + 1
			if count % 5 == 0:
				print('LHR_DATA = ' + str(raw_data[-1]))
			if key != False:
				if key == b's':
					break
				elif key == b'd':
					throw_away_the_record = True
					break
		if not throw_away_the_record:
			if not os.path.isdir(base_folder):
				os.mkdir(base_folder)
			wb = openpyxl.Workbook()
			ws = wb.active
			for i in range(len(raw_data)):
				ws.append([raw_data[i]])
			wb.save(excel_file)
