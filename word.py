from win32com import client as wc
import os
word = wc.Dispatch('Word.Application')
# format code https://msdn.microsoft.com/zh-cn/library/office/ff839952.aspx
format_dict={'doc':0,'docx':16}
def convert(fromdir,todir):
    doc = word.Documents.Open(fromdir)
    file_type=''
    if todir.rfind('.')==-1:
        file_type='docx'
        todir=todir+'.docx'
    else:
        file_type=todir[todir.rfind('.')+1:]
    if format_dict.has_key(file_type):
        doc.SaveAs(todir,format_dict[file_type])
    doc.Close()

homedir = os.getcwd()+'\\'
convert(homedir+'tst.doc',homedir+'test2.docx')
word.Quit()
