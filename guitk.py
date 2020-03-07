#   Author : github.com/slingthy
#   Time : March 4, 2020
#   ---------------------------------------------
#   The purpose is to extract someone's record 
#   from wechat group and store it in a docx file.
#   !!!JUST 'TEXT' message.
#   ---------------------------------------------
#   Before RUN:
#   1. COPY the WeChat Group message
#   2. PASTE into Word document as 'xxxxxx.docx'
#   3. SAVE it, do NOT make any change
#   ---------------------------------------------
#   Dependencies : python-docx, lxml

import os
import groupexporter
from tkinter.filedialog import askopenfilename
from tkinter import *
import tkinter.messagebox

class my_gui():
    def __init__(self,win):
        self.win=win
        
    def set_win(self):
        global pathroad
        global groupmember
        global savename
        pathroad=StringVar()
        groupmember=StringVar()
        savename=StringVar()
        self.win.title("WGExporterV1.0  by: slingthy")
        self.win.geometry('380x250')
        self.win.iconbitmap(".\\image\\slin.ico")
        Label(self.win,text='Path').grid(row=1,column=1,ipadx=20,ipady=20)
        E1=Entry(self.win,textvariable=pathroad)
        E1.grid(row=1,column=2,padx=20,pady=20)
        Button(self.win,text='choose',command=self.selectPath).grid(row=1,column=3)
        #row2
        Label(self.win,text='WechatID').grid(row=2,column=1,ipadx=20,ipady=20)
        E2=Entry(self.win,textvariable=groupmember)
        E2.grid(row=2,column=2)
        #row3
        Label(self.win,text='Save').grid(row=3,column=1,ipadx=20,ipady=20)
        E3=Entry(self.win,textvariable=savename)
        E3.grid(row=3,column=2)
        Label(self.win,text='.docx').grid(row=3,column=3)
        #trigger groupexporter.py
        Button(self.win,text='OK',activebackground='grey',command=self.pystart).grid(row=5,columnspan=4)

    def selectPath(self):
        path_=askopenfilename(filetypes=[('Word 文档', '*.docx')])
        pathroad.set(path_)

    def pystart(self):
        try:
            groupexporter.pymain(pathroad.get(),groupmember.get(),savename.get())
            tkinter.messagebox.showinfo("WeChatGroup Exporter_1.0  by: slingthy","successful!")
            self.win.destroy()
        except:
            tkinter.messagebox.showerror('Error','Back and Correct!')
    
    
def gui_start():
    win=Tk()
    portal = my_gui(win)
    # 设置根窗口默认属性
    portal.set_win()
    win.mainloop()     

gui_start()

