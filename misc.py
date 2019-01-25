import time
from datetime import datetime


def is_interactive():
    import __main__ as main
    return not hasattr(main, '__file__')

def register_excel_application():
    from win32com.client import Dispatch
    excel = Dispatch('Excel.Application')
    try:
        excel.Visible = True
    except:
        pass
    return excel


def register_excel_testingbook(excel):
    books = []
    
    for i in range(1, 9999):
        try:
            books.append(excel.Workbooks.Item(i))
        except:
            break
        
    print('Workbook to select from: \n\t'+ '\n\t'.join([book.Name for book in books]))
    testing_sheet = None
    thebook = None
    
    for book in books:
        for i in range(50):
            try:
                sheet = book.Sheets(i)
            except: pass
            else:
                if sheet.Name == 'testingdataset':
                    testing_sheet = sheet
                    thebook = book
                    
                    
    if not testing_sheet:
        raise ValueError(
              'Workbook does not contain sheets named '  
              '"testingdataset". Create them save '
              'and try again.')    

    print('Workbook name:', thebook.Name)
    return thebook, testing_sheet            
    

def register_excel_workbook(excel):
    books = []
    
    for i in range(1, 9999):
        try:
            books.append(excel.Workbooks.Item(i))
        except:
            break
        
    print('Workbook to select from: \n\t'+ '\n\t'.join([book.Name for book in books]))
    inputs_sheet  = None
    outputs_sheet = None
    
    thebook = None
    
    for book in books:
        for i in range(50):
            try:
                sheet = book.Sheets(i)
            except: pass
            else:
                if sheet.Name == 'inputs':
                    inputs_sheet = sheet
                if sheet.Name == 'outputs':
                    outputs_sheet = sheet
                
                if inputs_sheet and outputs_sheet:
                    thebook = book
                    break
                    
    if not (inputs_sheet and outputs_sheet):
        raise ValueError(
              'Workbook does not contain sheets named '  
              '"inputs" or "outputs". Create them save '
              'and try again.')
        
    print('Workbook name:', thebook.Name)
    return thebook, inputs_sheet, outputs_sheet

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
