#!/usr/bin/env python
from Tkinter import Tk
from time import sleep
from tkMessageBox import showwarning
import win32com.client as win32

warn = lambda app:showwarning(app,'Exit?')
RANGE = range(3,8)

def excel():
    app = 'Excel'
    xl = win32.gencache.EnsureDispatch('%s.Application' % app)
    ss = xl.Workbooks.Add()
    sh = ss.ActiveSheet
    xl.Visible = True
    sleep(1)

    sh.Cells(1,1).Value = 'Python-to-%s Demo' % app
    sleep(1)
    for i in RANGE:
        sh.Cells(i,1).Value = 'Line %d' % i
        sleep(1)
    sh.Cells(i+2,1).Value = 'Th-th-th-that\'s all folks!'

    warn(app)
    ss.Close(False)
    xl.Application.Quit()

def word():
    app = 'Word'
    word = win32.gencache.EnsureDispatch('%s.Application'% app)
    doc = word.Documents.Add()
    word.Visible = True
    sleep(1)

    rng = doc.Range(0,0)
    rng.InsertAfter("Python-to-%s Test\r\n\r\n" % app)
    sleep(1)
    for i in RANGE:
        rng.InsertAfter('Line %d\r\n' % i)
        sleep(1)
    rng.InsertAfter("\r\nTh-th-th-that's all folks'")

    warn(app)
    doc.Close(False)
    word.Application.Quit()

def ppoint():
    app = 'PowerPoint'
    ppoint = win32.gencache.EnsureDispatch('%s.Application'% app)
    pres = ppoint.Presentations.Add()
    ppoint.Visible = True

    sl = pres.Slides.Add(1,win32.constants.ppLayoutText)
    sleep(1)
    sla = sl.Shapes[0].TextFrame.TextRange
    sla.Text = 'Python-to-%s Demo' % app
    sleep(1)
    slb = sl.shapes[1].TextFrame.TextRange
    for i in RANGE:
        slb.InsertAfter("Line %d\r\n" % i)
        sleep(1)
    slb.InsertAfter("\r\nTh-th-th-that's all floks!")

    warn(app)
    pres.Close()
    ppoint.Quit()

def outlook():
    app = 'Outlook'
    olook = win32.gencache.EnsureDispatch('%s.Application' % app)

    mail = olook.CreateItem(win32.constants.olMailItem)
    recip = mail.Recipients.Add('you@127.0.0.1')
    subj = mail.Subject = 'Python-to-%s Demo' % app
    body = ["Line %d" % i for i in RANGE]
    body.insert(0,'%s\r\n' % subj)
    body.append("\r\nTh-th-th-that's all floks")
    mail.Body = '\r\n'.join(body)
    mail.Send()

    ns = olook.GetNamespace("MAPI")
    obox = ns.GetDefaultFolder(win32.constants.olFolderOutbox)
    obox.Display()
    obox.Items.Item(1).Display()

    warn(app)
    olook.Quit()
    

Tk().withdraw()
outlook()
