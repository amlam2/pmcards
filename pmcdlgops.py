#coding=utf-8

import wx
from pmcards import applicationIcon
from pmcards import windowTitle
from libs.liblore import opsDict, allNodesDict

###############################################################################
# Начало класса диалога выбора ОПС (SelectOPS)
#
class SelectOPS(wx.Dialog):
    def __init__(self, allOPSNameList, opsSelectList, parent):
        self.allOPSNameList    = allOPSNameList
        self.byNodeOPSNameList = allOPSNameList
        
        self.selectedOPSNameList = []
        
        #self.opsSelectList  = opsSelectList
        wx.Dialog.__init__(self, None, -1, windowTitle, size=(-1, 570))
        
        ### > иконка приложения
        self.SetIcon(wx.IconFromBitmap(wx.Bitmap(applicationIcon)))
        ###
        
        ### > формирование списка узлов
        mainNode = '3'
        self.nodeList = [('0', u"Все узлы")]
        self.nodeList.append((mainNode, allNodesDict.get(mainNode).get('label')))
        for item in allNodesDict.get(mainNode).get('upsList'):
            self.nodeList.append((item, allNodesDict.get(item).get('label')))
        ###
        
        ### > выпадающий список выбора узла
        self.comboBox = wx.ComboBox(self, 500, self.nodeList[0][1], (90, 50),
                                   (200, -1), [i[1] for i in self.nodeList],
                                   wx.CB_DROPDOWN|wx.CB_READONLY|wx.TE_PROCESS_ENTER)
        self.comboBox.SetToolTipString(u"Выберите узел почтовой связи")
        self.comboBox.Bind(wx.EVT_COMBOBOX, self.OnComboBox)
        ###
        
        ### > окно выбора ОПС
        self.checkListBox = wx.CheckListBox(self, -1, (5, 5), (430, -1),\
                                self.byNodeOPSNameList, wx.LB_SINGLE)
        #self.checkListBox.SetChecked(self.opsSelectList)
        self.checkListBox.Bind(wx.EVT_CHECKLISTBOX, self.OnCheckListBox)
        ###
        
        ### > кнопка "ОК"
        self.okButton = wx.Button(self, wx.ID_OK, u"ОК", size =(100, 25))
        self.okButton.SetFocus()
        ###
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.comboBox, 0, wx.ALL|wx.EXPAND, 5)
        sizer.Add(self.checkListBox, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.okButton, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        
        self.SetSizer(sizer)
        self.Layout()
    
    def OnComboBox(self, event):
        cb = event.GetEventObject()
        
        ### > преобразование имени узла в его цифровой код
        for node in self.nodeList:
            if node[1] == cb.GetValue():
                selectedNode = node[0]
                break
        ###
        
        ### > формирование списка ОПС в узле
        if selectedNode == '0':
            self.byNodeOPSNameList = self.allOPSNameList
        else:
            self.byNodeOPSNameList = []
            for index in opsDict:
                if (opsDict[index].get('nameOPS') in self.allOPSNameList) and\
                                    (opsDict[index].get('node') == selectedNode):
                    self.byNodeOPSNameList.append(opsDict[index].get('nameOPS'))
        self.byNodeOPSNameList.sort()
        ###
        
        self.selectedOPSNameList = []
        self.checkListBox.Clear()
        self.checkListBox.AppendItems(self.byNodeOPSNameList)
    
    def OnCheckListBox(self, event):
        clb = event.GetEventObject()
        
        if clb.IsChecked(event.GetInt()):
            self.selectedOPSNameList.append(self.byNodeOPSNameList[event.GetInt()])
        else:
            self.selectedOPSNameList.remove(self.byNodeOPSNameList[event.GetInt()])
        

#
# Конец класса диалога SelectOPS
###############################################################################
