from win32com import client as wc

ppt_support=['PPT','PPTX','ppt','pptx']
word_support=['DOC','DOCX','doc','docx']
excel_support=['XLS','XLSX','xls','xlsx']

class converter:
    'converter for file'
    def __init__(self,ft):
        self.app=None
        if ft in ppt_support:
            self.app=powerpoint()
        elif ft in excel_support:
            self.app=excel()
        elif ft in word_support:
            self.app=word()
        assert self.app!=None,'initializing failed'
    def convert(self,fromdir,todir):
        'input should be the full path of file'
        self.app.convert(fromdir,todir)
    def quit(self):
        self.app.quit()

class MSapp:
    app=''
    format_dict={}
    format=''
    f=None
    def openfile(self,fromdir):
        pass
    def save(self,todir):
        pass
    def convert(self,fromdir,todir):
        self.openfile(fromdir)
        file_type=''
        if todir.rfind('.')==-1:
            file_type=self.format
            todir=todir+'.'+self.format
        else:
            file_type=todir[todir.rfind('.')+1:]
        if self.format_dict.has_key(file_type):
            self.f.SaveAs(todir,self.format_dict[file_type])
        self.f.Close()
    def quit(self):
        self.app.Quit()

class excel(MSapp):
    def __init__(self):
        self.app=wc.Dispatch('Excel.Application')
        self.format_dict={'xls':56,'xlsx':51}
        self.format='xlsx'
    def openfile(self,fromdir):
        self.f = self.app.Workbooks.Open(fromdir)
        self.app.DisplayAlerts=False

class powerpoint(MSapp):
    def __init__(self):
        self.app=wc.Dispatch('PowerPoint.Application')
        self.format_dict={'ppt':1,'pptx':11}
        self.format='pptx'
    def openfile(self,fromdir):
        self.f = self.app.Presentations.Open(fromdir,WithWindow=0)

class word(MSapp):
    def __init__(self):
        self.app=wc.Dispatch('Word.Application')
        self.format_dict={'doc':0,'docx':16}
        self.format='docx'
    def openfile(self,fromdir):
        self.f = self.app.Documents.Open(fromdir)

def easy_convert(fromdir,todir):
    assert fromdir.rfind('.')!=-1,'input file path error'
    assert todir.rfind('.')!=-1,'target file path error'
    from_type=fromdir[fromdir.rfind('.')+1:]
    to_type=todir[todir.rfind('.')+1:]
    c=None
    if from_type in ppt_support and to_type in ppt_support:
        c=converter('PPT')
    elif from_type in word_support and to_type in word_support:
        c=converter('DOC')
    elif from_type in excel_support and to_type in excel_support:
        c=converter('XLS')
    assert c!=None,'file type error'
    c.convert(fromdir,todir)
    c.quit()
