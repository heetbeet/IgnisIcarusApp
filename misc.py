import time
from datetime import datetime


def is_interactive():
    import __main__ as main
    return not hasattr(main, '__file__')

def get_ignis_spreadsheet():
    import pythoncom
    import win32api
    import win32com.client


    for moniker in pythoncom.GetRunningObjectTable():
        try:
            # Workbook implements IOleWindow so only consider objects implementing that
            window = moniker.BindToObject(pythoncom.CreateBindCtx(0), None, pythoncom.IID_IOleWindow)
            disp = window.QueryInterface(pythoncom.IID_IDispatch)


            # Get a win32com Dispatch object from the PyIDispatch object as it's
            # easier to work with.
            book = win32com.client.Dispatch(disp)

        except pythoncom.com_error:
            # Skip any objects we're not interested in
            continue

        try:
            book.Sheets(1) #Object is a book with sheets
        except:
            continue
            
        bookname = moniker.GetDisplayName(pythoncom.CreateBindCtx(0), None)
        print('Test workbook: ', bookname)


        inputs  = [i for i in book.Sheets if i.Name.lower() == 'inputs']
        outputs = [i for i in book.Sheets if i.Name.lower() == 'outputs']

        if len(inputs) and len(outputs):
            print('Yes -->', bookname)
            return book, inputs[0], outputs[0]
        

def str2bits(s):
    result = []
    for c in s:
        bits = bin(ord(c))[2:]
        bits = '00000000'[len(bits):] + bits
        result.extend([int(b) for b in bits])
    return result

def bits2str(bits):
    bitgroups = [bits[i:i+8] for i in range(0,len(bits),8)]
    int_list = []
    for bit in bitgroups:
        int_list.append(0)
        for i, val in enumerate(bit[::-1]):
            int_list[-1] += (2**(i))*bool(val)
    str_out = ''.join([chr(i) for i in int_list])
    return str_out

def bits2int(bits):
    out_int = 0
    for i, val in enumerate(bits[::-1]):
        out_int += (2**(i))*bool(val)
    return out_int

class timeStrober:
    def __init__(self, inpstr):
        self.set_timings(inpstr)
    
    def set_timings(self, inpstr):
        try:
            inpstr*1.0
        except: pass
        else:
            inpstr = 'on' if inpstr else 'off'
            
        inpstr = inpstr.lower().strip()

        if inpstr.startswith('t'):
            self.pperiod, self.pwidth = time.time(), float(inpstr[1:].strip())

        elif inpstr.startswith('s'):
            self.pperiod, self.pwidth = (float(inpstr[1:].split(',')[0].strip()),
                                         float(inpstr[1:].split(',')[1].strip()) )
            if self.pperiod == 0:
                self.pperiod += 0.05
        elif inpstr == 'on':
            self.pperiod, self.pwidth = time.time(), time.time()

        elif inpstr == 'off':
            self.pperiod, self.pwidth = time.time(), 0

        else:
            raise('Error timestrobe inputs.')
        
    def is_on(self):
        now = time.time()
        if now - int(now/self.pperiod)*self.pperiod < self.pwidth:
            return True
        else:
            return False

def get_instruments(comname):
    import minimalmodbus

    minimalmodbus.BAUDRATE = 9600
    minimalmodbus.CLOSE_PORT_AFTER_EACH_CALL = False
    ins1 = minimalmodbus.Instrument(comname, 1)
    ins2 = minimalmodbus.Instrument(comname, 2)
    ins3 = minimalmodbus.Instrument(comname, 3)
    ins4 = minimalmodbus.Instrument(comname, 4)
    ins5 = minimalmodbus.Instrument(comname, 5)
    ins6 = minimalmodbus.Instrument(comname, 6)

    return ins1, ins2, ins3, ins4, ins5, ins6


class test_writer:
    def __init__(self):
        self.curr_line = 6
    
    def do_readings(self, inputs_sheet, testing_sheet):
        for i in range(self.curr_line,60000):
            if not inputs_sheet.Range('A'+str(i)).Value:
                self.curr_line = i
                break
                
        cells = 'A%d:CC%d'%(self.curr_line, self.curr_line)
        inputs_sheet.Range(cells).Value = testing_sheet.Range(cells).Value


class inputs_writer:
    def __init__(self):
        self.curr_line = 6
    
    def do_readings(self, inputs_sheet, ins1, ins2, ins3, ins4, ins5, ins6):
        for i in range(self.curr_line,60000):
            if not inputs_sheet.Range('A'+str(i)).Value:
                self.curr_line = i
                break
                
        data =(
            [str(datetime.now())]+
            ins2.read_registers(512, 8)+
            ins2.read_registers(520, 8)+
            ins3.read_registers(512, 8)+
            ins3.read_registers(520, 8)+
            str2bits(ins1.read_string(320,1))[::-1][:8]+
            ins4.read_registers(512, 8)+
            ins4.read_registers(520, 8)+
            ins5.read_registers(512, 8)+
            ins5.read_registers(520, 8)+
            str2bits(ins6.read_string(320,1))[::-1][:8]
        )

        inputs_sheet.Range('A%d:CC%d'%(self.curr_line, self.curr_line)
                          ).Value = data
