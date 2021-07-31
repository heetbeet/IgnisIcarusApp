import time
from datetime import datetime
import os
import pythoncom
import win32api
import win32com.client
import minimalmodbus

minimalmodbus.BAUDRATE = 9600
minimalmodbus.CLOSE_PORT_AFTER_EACH_CALL = False

alph = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
num2col = [i for i in alph] + [i+j for i in alph for j in alph]

def is_interactive():
    import __main__ as main
    return not hasattr(main, '__file__')


def spread_iterator():
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

        yield bookname, book


def get_ignis_spreadsheet():
    for bookname, book in spread_iterator():
        print('Test workbook: ', bookname)

        inputs  = [i for i in book.Sheets if i.Name.lower() == 'inputs']
        outputs = [i for i in book.Sheets if i.Name.lower() == 'outputs']

        if len(inputs) and len(outputs):
            print('Yes -->', bookname)
            return book, inputs[0], outputs[0]


def get_spreadsheet_by_name(spreadname):
    for bookname, book in spread_iterator():
        print('Test workbook: ', bookname)
        fname = os.path.split(bookname)[-1].lower()

        fexts = ['.xls', '.csv', '.txt']
        for fext in fexts:
            if fext in fname:
                fname = fext.join(fname.split(fext)[:-1])
        if fname == spreadname.lower():
            return book


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


def get_instruments(comname, nr_of_devices):
    instances = []
    for i in range(1, nr_of_devices+1):
        instances.append(minimalmodbus.Instrument(comname, i))
    
    return instances


def get_instrument(comname, device):
    return minimalmodbus.Instrument(comname, device)


def write_to_inst(ins, bits):
    try:
        ins.write_register(320, bits2int(bits[::-1]))
        return True
    except OSError:
        return False
    
def get_mode_limit(wb):
    results_sheet = [i for i in wb.Sheets if i.Name.lower() == 'results'][0]
    compiled_sheet = [i for i in wb.Sheets if i.Name.lower() == 'compiled data'][0]
    return [results_sheet.Range("AW3").Value,
            compiled_sheet.Range("CX4").Value]


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
        self.sensitivity_col = None
        self.results_sheet = None
        self.inputs_sheet = None
    
    def do_readings(self, wb, inputs_sheet, ins1, ins2, ins3, ins4, ins5, ins6):
        if self.sensitivity_col is None:
            self.inputs_sheet   = [i for i in wb.Sheets if i.Name.lower() == 'inputs' ][0]
            self.results_sheet  = [i for i in wb.Sheets if i.Name.lower() == 'results'][0]
            for i, val in enumerate(self.results_sheet.Range("5:5").Value[0]):
                if str(val).lower().strip() == 'sensitivity':
                    self.sensitivity_col = num2col[i]
        
        for i in range(self.curr_line,60000):
            if not inputs_sheet.Range('A'+str(i)).Value:
                self.curr_line = i
                break
        try:        
            data =(
                [str(datetime.now())] +
                ([0]*16 if ins2 is None else ins2.read_registers(512, 8) + ins2.read_registers(520, 8)) +
                ([0]*16 if ins3 is None else ins3.read_registers(512, 8) + ins3.read_registers(520, 8)) +
                # noinspection PyTypeChecker
                ([0]*8 if ins1 is None else str2bits(ins1.read_string(320, 1))[::-1][:8]) +
                ([0]*16 if ins4 is None else ins4.read_registers(512, 8) + ins4.read_registers(520, 8)) +
                ([0]*16 if ins5 is None else ins5.read_registers(512, 8) + ins5.read_registers(520, 8)) +
                ([0]*8 if ins6 is None else str2bits(ins6.read_string(320, 1))[::-1][:8]) +
                [None]*16
            )
        except:
            return False

        #Some excel conversions ans lookups
        alph = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
        xl = alph+[i+j for i in alph for j in alph]
        
        row = self.curr_line
        
        col0 = xl[0]           #datalines from the instuments
        col1 = xl[len(data)-1]
        
        col2 = xl[len(data)]  #extra feedback from excel
        col3 = xl[len(data)+1]
        col4 = xl[len(data)+2]
        
        self.inputs_sheet.Range(f'{col3}{row}:{col4}{row}').Value = get_mode_limit(wb)
        self.inputs_sheet.Range(f'{col2}{row}').Value = 0 #self.results_sheet.Range('%s%d'%(self.sensitivity_col, self.curr_line-1)).Value
        self.inputs_sheet.Range(f'{col0}{row}:{col1}{row}').Value = data
        
        return True