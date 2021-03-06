# -*- coding: utf-8 -*-
# !python3
"""
PPT convert PDF
"""

import os

import filetype

import win32com

from win32com.client import Dispatch

def ppt2pdf(filename,dst_filename):
    """A folder with the same name as the PPT file will be created in the 
    same directory.This folder contains all PDF images generated by PPT 
    files.Where * filename * is the path to the PPT file.* dst_filename * 
    is the destination file format.
    """
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # ppt.DisplayAlerts = False
    pptSel = ppt.Presentations.Open(filename, WithWindow = False)
    pptSel.SaveAs(dst_filename,32); # 32 for ppt to pdf
    ppt.Quit()

ppt_dir = os.getcwd() # Get the current working directory

for fn in (fns for fns in os.listdir(ppt_dir) 
          if fns.endswith(('.ppt','.pptx'))):
    try:
        kind = filetype.guess(fn)
        if kind is None:
            print('Cannot guess file type ' + fn)
        elif kind.mime == 'application/zip':  # File type must be PPT
            file_name = os.path.splitext(fn)[0]
            print('Converting ' + fn)
            ppt_file = os.path.join(ppt_dir, fn)
            img_file = os.path.join(ppt_dir, file_name + '.pdf')
            ppt2pdf(ppt_file,img_file)
    except:
        print('Getting file type error ' + fn)

print('pdf conversion completed')