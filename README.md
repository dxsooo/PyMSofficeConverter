# PyMSofficeConverter

## Overview  
Microsoft office files(.doc/.xls/.ppt etc) converter written in Python.  
Recently supports  

*  Microsoft word document convert between **.doc** and **.docx**  
*  Microsoft excel workbook convert between **.xls** and **.xlsx**  
*  Microsoft powerpoint presentation convert between **.ppt** and **.pptx**

## Requirements  
* Python 2.7
* Windows 7 or above
* Microsoft Office 2013 or above
* [pywin32](http://sourceforge.net/projects/pywin32/)

## Documentation
Copy the `MSofficeConverter.py` to current directory or under `site-packages` of your python.  
And then  
```python
    from MSofficeConverter import easy_convert
    
    easy_convert('x:/test.doc','x:/test.docx') 
    # full path of file, here convert test.doc to test.docx
```
Also, you can use a class to do this, but please remember to call `quit()` when your work over.  
```python
    from MSofficeConverter import converter

    xx=converter('PPT') # the file type you are going to convert, accept 'PPT','DOC','XLS'
    
    xx.convert('x:/test.pptx','x:/tst.ppt')
    # xx.convert('x:/tst.ppt','x:/test2.pptx') # you can do more before you quit
    
    xx.quit() # must do this
```
It is strongly recommanded that if you are dealing with a quantity of files, you should use the `converter` class and set a loop to call `convert()` instead of `easy_convert()` 