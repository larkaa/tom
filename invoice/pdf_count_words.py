#!/usr/bin/env python
# coding: utf-8

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
from easygui import fileopenbox, msgbox


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    return text


def main():
    openfiles = fileopenbox("Choose one or more pdf files","PDF word count",default=r"C:\Users\user\Desktop\*.pdf",filetypes="*.pdf",multiple=True)


    for filename in openfiles:
    
        res = convert_pdf_to_txt(filename)
        msg = '\n# words: '+ str(len(res.split())) + '\nOriginal file: '+filename
        msgbox(msg)



if __name__ == "__main__":
    main()
