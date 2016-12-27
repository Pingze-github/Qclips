# !/usr/bin/env python
# -*-coding=utf-8-*-

'''
窗口活动性检查
解决程序未响应问题 对键盘和鼠标的监听都会导致未响应! ???似乎不是这个问题？
启动界面自动活动
失去焦点自动关闭界面
优化界面
设置快捷键粘贴
*考虑将win32clipboard换成QClipboard，来获取剪贴板内容
*考虑是否可以将获取鼠标位置的模块从PyHook换成PyQt4
'''

import sys
import time
import threading
from PyQt4.QtGui import *
from PyQt4.QtCore import *
import win32clipboard as cb
import win32con
import pyhk
import pyHook
import pythoncom

#全局变量
clips = [] #储存剪贴板内容的列表
clip = ''
maxClips = 5 #最多储存的剪贴项目数
mousePosition = (0,0)

def getClip():
    '''
    获取剪贴板中最近的一个内容
    '''
    cb.OpenClipboard()
    clip = cb.GetClipboardData(win32con.CF_TEXT).decode('gb2312')
    cb.CloseClipboard()
    return clip

def setClip(string):
    '''
    更改剪贴板中最近一个内容为输入字符串
    '''
    cb.OpenClipboard()
    cb.SetClipboardData(win32con.CF_TEXT, string)
    cb.CloseClipboard()



def onMouseEvent(event):
    '''
    pyHook设定函数，鼠标移动时触发
    '''
    global mousePosition
    mousePosition = event.Position
    return True

def listenMouse():
    '''
    监听鼠标线程
    '''
    hm = pyHook.HookManager()
    hm.MouseMove = onMouseEvent 
    hm.HookMouse()
    pythoncom.PumpMessages()

def emitShowSig():
    yc.showSig.emit()

def emitQuitSig():
    yc.quitSig.emit()

def listenBoard():
    '''
    监听键盘线程
    '''
    sysHotkey = pyhk.pyhk()
    sysHotkey.addHotkey(['Ctrl','Alt','Z'], emitShowSig)
    sysHotkey.addHotkey(['Ctrl','Alt','F12'], emitQuitSig)
    sysHotkey.start()

def timerCheck():
    '''
    循环检查线程
    '''
    clipNow = getClip()
    if clip != clipNow:
        global clip
        clip = clipNow
        yc.clipsAppend.emit()
    if not yc.isActiveWindow():
        yc.isNotActiveWindow.emit()
    global timer
    timer = threading.Timer(1, timerCheck)
    timer.start()

class YourClips(QWidget):
    clipsReversed = [] #储存剪贴板内容的列表（逆序）
    clipLabels = [] #储存剪贴板Label的列表
    clip = '' #储存剪贴板
    clipsAppend = pyqtSignal() #剪贴板检查信号检查信号
    clipsChanged = pyqtSignal() #剪贴板列表变化信号
    showSig = pyqtSignal() #显示界面快捷键信号
    quitSig = pyqtSignal() #退出程序快捷键信号
    isNotActiveWindow = pyqtSignal()

    def __init__(self, parent=None):
        super(YourClips, self).__init__()
        self.move(0,0)
        self.vBox = QVBoxLayout()
        self.setLayout(self.vBox)
        self.clipsAppend.connect(self.appendClips)
        self.isNotActiveWindow.connect(self.hide)
        self.clipsChanged.connect(self.replot)
        self.showSig.connect(self.showYC)
        self.quitSig.connect(self.quitAll)
    
    def quitAll():
        '''
        全局快捷键退出程序
        '''
        tb.stop()
        tm.stop()
        tc.stop()
        sys.exit(0)

    def showYC(self):
        '''
        显示YourClips剪贴板界面并设为活动窗口
        '''
        self.move(mousePosition[0], mousePosition[1])
        self.showNormal()
        self.activateWindow()
        
    def keyPressEvent(self, event):
        '''
        设定剪贴板界面获得焦点时的快捷键V。最大列表数目不超过9，默认5
        '''
        if event.key() == Qt.Key_Escape:
            self.hide()
        elif event.key() == Qt.Key_1:
            self.pasteClip(1)
        elif event.key() == Qt.Key_2:
            self.pasteClip(2)
        elif event.key() == Qt.Key_3:
            self.pasteClip(3)
        elif event.key() == Qt.Key_4:
            self.pasteClip(4)
        elif event.key() == Qt.Key_5:
            self.pasteClip(5)

    def pasteClip(self, index):
        '''
        依据序号粘贴选定的剪贴板内容
        '''
        if index <= len(self.clipLabels):
            print(self.clipsReversed[index-1].encode('utf-8'))

    def appendClips(self):
        '''
        添加当前剪贴板内容进剪贴序列
        '''
        global clips
        clips.append(clip)
        if len(clips) > maxClips:
            clips = clips[-maxClips:]
        print('Num of clips: %d /// They are : %s' % (len(clips),clips))
        self.clipsChanged.emit()

    def clearQLayout(self):
        '''
        动态清空布局容器并释放标签
        '''
        num = self.vBox.count()
        while num > 0:
            w = self.vBox.itemAt(num-1).widget()
            self.vBox.removeWidget(w)
            w.deleteLater()
            num -= 1

    def replot(self):
        '''
        根据捕捉到的剪贴板序列重新绘制面板
        '''
        self.clipsReversed = clips[:]
        self.clipsReversed.reverse()
        self.clipLabels = []
        self.clearQLayout()
        index = 1
        for clipr in self.clipsReversed:
            label = QLabel(text=str(index)+': '+clipr)
            label.resize(250,30)
            self.clipLabels.append(label)
            self.vBox.addWidget(self.clipLabels[len(self.clipLabels)-1])
            self.resize(250, 30*self.vBox.count())
            index += 1
        self.repaint()

def main():
    '''
    主函数：创建了Qt主程序，创建了监听剪贴板、鼠标移动、键盘输入线程，并启动
    '''
    global yc, tb, tm, tc
    app = QApplication(sys.argv)
    yc = YourClips()
    QApplication.setQuitOnLastWindowClosed(False)
    yc.setWindowFlags(Qt.FramelessWindowHint) #无边框
    yc.setWindowFlags(Qt.WindowStaysOnTopHint) #总在最前
    yc.setWindowFlags(Qt.Tool) #隐藏任务栏标识
    tm = threading.Thread(target = listenMouse)
    tb = threading.Thread(target = listenBoard)
    tc = threading.Thread(target = timerCheck)
    tm.start()
    tb.start()
    tc.start()
    #yc.show() # ?加了这行就不会报QObject::startTimer: timers cannot be started from another thread
    #yc.hide()
    print('--- 程序初始化成功，进入事件循环 ---\n')
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
