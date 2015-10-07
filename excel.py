from win32com import client as wc
import os
excel = wc.Dispatch('Excel.Application')
# format code https://msdn.microsoft.com/ZH-CN/library/office/ff198017.aspx
format_dict={'xls':56,'xlsx':51}
def convert(fromdir,todir):
    xls = excel.Workbooks.Open(fromdir)
    excel.DisplayAlerts=False
    file_type=''
    if todir.rfind('.')==-1:
        file_type='xlsx'
        todir=todir+'.xlsx'
    else:
        file_type=todir[todir.rfind('.')+1:]
    if format_dict.has_key(file_type):
        xls.SaveAs(todir,format_dict[file_type])
    xls.Close()

homedir = os.getcwd()+'\\'
convert(homedir+'test.xlsx',homedir+'tst.xls')
excel.Quit()
