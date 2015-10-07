from win32com import client as wc
import os
powerpoint = wc.Dispatch('PowerPoint.Application')
# format code https://msdn.microsoft.com/ZH-CN/library/office/ff746500.aspx
format_dict={'ppt':1,'pptx':11}
def convert(fromdir,todir):
    ppt = powerpoint.Presentations.Open(fromdir,WithWindow=0)
    file_type=''
    if todir.rfind('.')==-1:
        file_type='pptx'
        todir=todir+'.pptx'
    else:
        file_type=todir[todir.rfind('.')+1:]
    if format_dict.has_key(file_type):
        ppt.SaveAs(todir,format_dict[file_type])
    ppt.Close()

homedir = os.getcwd()+'\\'
convert(homedir+'test2.pptx',homedir+'tst.ppt')
powerpoint.Quit()
