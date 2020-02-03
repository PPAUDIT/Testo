import os , sys
os.chdir('C:\\Users\\griggj\\Desktop\\')
import xlwings as xw
# open Excel app in the background
app_excel = xw.App(visible = False)

wbk = xw.Book( 'C:\Users\griggj\Desktop\test.xlsx')
wbk.api.RefreshAll()

# two options to save

wbk.save( 'C:\Users\griggj\Desktop\All Actions2.xlsx' ) # this will save the file with a name

# kill Excel process
app_excel.kill()
del app_excel