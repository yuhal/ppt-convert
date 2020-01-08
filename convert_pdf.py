# -*- coding: utf-8 -*-
# !python3
'''
PPT convert PDF
'''

import os

import time

import math

import codecs

import shutil

# import win32com

# from win32com.client import Dispatch, constants

# from PIL import Image, ImageDraw, ImageFont

# from reportlab import rl_settings
# from reportlab.lib.units import inch
# from reportlab.lib.pagesizes import letter, A4, landscape
# from reportlab.pdfgen import canvas
# from reportlab.platypus import SimpleDocTemplate

# ppt转成png文件
def ppt2png(filename,dst_filename):
    """ Return the ASCII characters in the file specified by *path* and *paths*.
    The file path is determined by concatenating *path* and any members of
    *paths* with a directory separator in between.
    """
    print(filename)
    print(dst_filename)
    exit(0)
    
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # ppt.DisplayAlerts = False
    pptSel = ppt.Presentations.Open(filename, WithWindow = False)
    pptSel.SaveAs(dst_filename,18); # with 17, jpeg
    ppt.Quit()

# 增加水印
def add_mark(imgFile, txtMark):
    img = Image.open(imgFile)
    imgWidth, imgHeight = img.size

    # http://blog.csdn.net/Dou_CO/article/details/17715919
    textImgW = int(imgWidth * 1.5)    # 确定写文字图片的尺寸，要比照片大
    textImgH = int(imgHeight * 1.5)
    blank = Image.new("RGB",(textImgW,textImgH),"white")  # 创建用于添加文字的空白图像
    d = ImageDraw.Draw(blank)
    d.ink = 0 + 0 * 256 + 0 * 256 * 256 
    markFont = ImageFont.truetype('simhei.ttf', size=18)
    fontWidth, fontHeight = markFont.getsize(txtMark)
    d.text(((textImgW - fontWidth)/2, (textImgH - fontHeight)/2), txtMark, font=markFont)
    textRotate = blank.rotate(30)

    rLen = math.sqrt((fontWidth/2)**2+(fontHeight/2)**2)   
    oriAngle = math.atan(fontHeight/fontWidth)
    cropW = rLen*math.cos(oriAngle + math.pi/6) *4   # 被截取区域的宽高
    cropH = rLen*math.sin(oriAngle + math.pi/6) *4
    box = [int((textImgW-cropW)/2-1),int((textImgH-cropH)/2-1)-50,int((textImgW+cropW)/2+1),int((textImgH+cropH)/2+1)]
    textImg = textRotate.crop(box)  # 截取文字图片
    pasteW,pasteH = textImg.size
    # 旋转后的文字图片粘贴在一个新的blank图像上
    textBlank = Image.new("RGB",(imgWidth,imgHeight),"white")
    pasteBox = (int((imgWidth-pasteW)/2-1),int((imgHeight-pasteH)/2-1))
    textBlank.paste(textImg,pasteBox)
    waterImage = Image.blend(img.convert('RGB'),textBlank,0.1)

    fileDir = os.path.dirname(imgFile) + '-png'
    fileName = os.path.join(fileDir, os.path.basename(imgFile))
    waterImage.save(fileName,'png')

# 合并输出PDF
def topdf(path,recursion=None,pictureType=None,sizeMode=None,width=None,height=None,fit=None,save=None):
    """
    Parameters
    ----------
    path : string
           path of the pictures
    pictureType : list
                  type of pictures,for example :jpg,png...
    sizeMode : int 
           None or 0 for pdf's pagesize is the biggest of all the pictures
           1 for pdf's pagesize is the min of all the pictures
           2 for pdf's pagesize is the given value of width and height
           to choose how to determine the size of pdf
    width : int
            width of the pdf page
    height : int
            height of the pdf page
    fit : boolean
           None or False for fit the picture size to pagesize
           True for keep the size of the pictures
           wether to keep the picture size or not
    save : string 
           path to save the pdf 
    """

    filelist = os.listdir(path)
    filelist = [os.path.join(path, f) for f in filelist]
    filelist.sort(key=lambda x: os.path.getmtime(x))

    maxw = 0
    maxh = 0
    if sizeMode == None or sizeMode == 0:
        for i in filelist:
            im = Image.open(i)
            if maxw < im.size[0]:
                maxw = im.size[0]
            if maxh < im.size[1]:
                maxh = im.size[1]
    elif sizeMode == 1:
        maxw = 999999
        maxh = 999999
        for i in filelist:
            im = Image2.open(i)
            if maxw > im.size[0]:
                maxw = im.size[0]
            if maxh > im.size[1]:
                maxh = im.size[1]
    else:
        if width == None or height == None:
            raise Exception("no width or height provid")
        maxw = width
        maxh = height

    maxsize = (maxw,maxh)
    if save == None:
        filename_pdf = os.path.join(path, path.split('\\')[-1])
    else:
        filename_pdf = os.path.join(save, path.split('\\')[-1])
    
    l = len(filelist)
    for i in range(l): 
        pdf_dir = filename_pdf.replace("png", "pdf") + '\\' + format(i+1) + '.pdf'
        c = canvas.Canvas(pdf_dir, pagesize=maxsize )
        print('准备生成' + pdf_dir)
        (w, h) =maxsize
        width, height = letter 
        if fit == True:
            c.drawImage(filelist[i] , 0,0) 
        else:
            c.drawImage(filelist[i] , 0,0,maxw,maxh) 
        c.showPage()  
        c.save()


ppt_dir = os.getcwd()
markText = 'yuhal.com'

for fn in (fns for fns in os.listdir(ppt_dir) if fns.endswith(('.ppt','.pptx')) if fns.startswith(('0','1'))):
    file_name = os.path.splitext(fn)[0]
    print(file_name)
    ppt_file = os.path.join(ppt_dir, fn)
    img_file = os.path.join(ppt_dir, file_name+'.png')
    ppt2png(ppt_file,img_file)
    img_dir = os.path.join(ppt_dir, file_name)
    
    imgFileList = os.listdir(img_dir)
    imgFileList = [os.path.join(img_dir, f) for f in imgFileList]
    imgFileList.sort(key=lambda x: os.path.getmtime(x))

    pngDir = img_dir + '-png'
    pdfDir = img_dir + '-pdf'
    os.makedirs(pngDir, exist_ok=True)
    os.makedirs(pdfDir, exist_ok=True)
    for imgFile in imgFileList:
        print('准备生成' + imgFile.replace(img_dir, pngDir))
        add_mark(imgFile, markText.strip('\n'))
    print("生成png完成")
    topdf(path=pngDir, save=ppt_dir)
    print("生成PDF完成")
    shutil.rmtree(img_dir)
