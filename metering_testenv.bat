
# coding: utf-8

# In[131]:


''' >NUL  2>NUL
@echo off
cd /d %~dp0
:loop
python %0 %*
goto loop
'''


import misc; 
import importlib; importlib.reload(misc)
from misc import *
from pywintypes import com_error
import numpy as np

excel =  register_excel_application()


# In[132]:


wbtest, testing_sheet = register_excel_testingbook(excel)


# In[133]:


wb, inputs_sheet, outputs_sheet = register_excel_workbook(excel)


# In[134]:


inst = test_writer()
prev = -9999999
prev_save = -9999999

print('Start')
try:
    
    i=-1
    while(1):
        i+=1
        #time.sleep(0.00001)
        try:
            inst.do_readings(inputs_sheet, testing_sheet)

        except com_error:
            if i%10 ==0:
                print('.', end='\n' if i%600==0 else '')
        #except ValueError:
        #    print('\nValueError occured')

except KeyboardInterrupt:
    print('\nStop')
    
    

