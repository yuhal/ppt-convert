# -*- coding: utf-8 -*-
# !python3
"""
PPT to PDF and split
"""

import os

import filetype

import win32com

from win32com.client import Dispatch

from PyPDF2 import PdfFileReader, PdfFileWriter

def ppt2pdf(filename,dst_filename):
    """
    :param filename: PPT file path
    :param filename: Split into single-page PNG file storage path
    """
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # ppt.DisplayAlerts = False
    pptSel = ppt.Presentations.Open(filename, WithWindow = False)
    pptSel.SaveAs(dst_filename,32); # 32 for ppt to pdf
    ppt.Quit()

def split_pdf(infile, out_path):
    """
    :param infile: PDF file to be split
    :param out_path: Split into single-page PDF file storage path
    :return: 无
    """
    if not os.path.exists(out_path):
        os.makedirs(out_path)
    with open(infile, 'rb') as infile:
    
        reader = PdfFileReader(infile)
        number_of_pages = reader.getNumPages()  #计算此PDF文件中的页数
        
        for i in range(number_of_pages):
            writer = PdfFileWriter()
            writer.addPage(reader.getPage(i))
            out_file_name = out_path + str(i+1)+'.pdf'
            with open(out_file_name, 'wb') as outfile:
                writer.write(outfile)

if __name__ == '__main__':
    work_dir = os.getcwd() # Get the current working directory
    for fn in (fns for fns in os.listdir(work_dir) 
              if fns.endswith(('.ppt','.pptx'))):
        try:
            kind = filetype.guess(fn)
            if kind is None:
                print('Cannot guess file type ' + fn)
            elif kind.mime == 'application/zip':  # File type must be PPT
                file_name = os.path.splitext(fn)[0]
                print('Converting ' + fn)
                ppt_file = os.path.join(work_dir, fn)
                img_file = os.path.join(work_dir, file_name + '.pdf')
                ppt2pdf(ppt_file,img_file)
        except:
            print('Getting file type error ' + fn)
    print('pdf conversion completed')

    for pdf in (pdfs for pdfs in os.listdir(work_dir) 
              if pdfs.endswith(('.pdf'))):
        try:
            pdftype = filetype.guess(pdf)
            if pdftype is None:
                print('Cannot guess file type ' + pdf)
            elif pdftype.mime == 'application/pdf':  # File type must be PDF
                file_name = os.path.splitext(pdf)[0]
                in_File = pdf
                print('spliting ' + in_File)
                out_Path = file_name + '-pdf/'
                split_pdf(in_File, out_Path)
                os.remove(in_File)
        except:
            print('Getting file type error ' + fn)

    print('pdf split completed')


