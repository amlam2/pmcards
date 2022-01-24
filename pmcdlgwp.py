#coding=utf-8

import wx
from pmcards import applicationIcon
from pmcards import windowTitle

###############################################################################
# Начало класса диалога выбора РМ (SelectWP)
#
class SelectWP(wx.Dialog):
    def __init__(self, opsWPNameList, wpSelectList, parent):
        self.opsWPNameList = opsWPNameList
        self.wpSelectList  = wpSelectList
        wx.Dialog.__init__(self, None, -1, windowTitle, size=(-1, 570))
        
        ### > иконка приложения
        self.SetIcon(wx.IconFromBitmap(wx.Bitmap(applicationIcon)))
        ###
        
        ### > надпись на форме
        txt = wx.StaticText(self, -1, u"Выберите одно или несколько РМ:")
        ###
        
        ### > окно выбора ОПС
        self.checkListBox = wx.CheckListBox(self, -1, (5, 5), (430, -1),\
                                self.opsWPNameList, wx.LB_SINGLE)
        self.checkListBox.SetChecked(self.wpSelectList)
        ###
        
        ### > кнопка "ОК"
        self.okButton = wx.Button(self, wx.ID_OK, u"ОК", size =(100, 25))
        self.okButton.SetFocus()
        ###
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(txt, 0, wx.ALIGN_LEFT | wx.ALL, 5)
        sizer.Add(self.checkListBox, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.okButton, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        
        self.SetSizer(sizer)
        self.Layout()
#
# Конец класса диалога SelectWP
###############################################################################
