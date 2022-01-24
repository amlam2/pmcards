#!/usr/bin/env python
# coding=utf-8

import os
import wx.grid
import MySQLdb
from threading import Thread
from wx.lib.pubsub import Publisher
import string
import time, datetime
from decimal import Decimal
from re import compile
from libs.liblore import opsDict
from libs.libcard import bpkDict, operationDict
from libs.libcard import hostDict, rcDict
from libs.libwork import toUserView2 #, toUserView
from libs.libwork import dateTmplStr
from pmcards import applicationIcon, windowTitle
from pmcards import dbHost, dbUser, dbPasswd, dbName
from pmcards import picsDir, reportDir, startDir

typeOfDevice  = 'all'

periodList  = [ u"сегодня",
                u"вчера",
                u"позавчера",
                u"последние 3 дня",
                u"последние 10 дней",
                u"последние 30 дней"]

filterDict  = { u"все операции"                   : {'order'    : 0,
                                                     'RC'       : []},
                u"успешные операции"              : {'order'    : 1,
                                                     'RC'       : [
                                                            '00', '08', '10'\
                                                            '11', '16', '32'
                                                                  ]},
                u"неуспешные операции"            : {'order'    : 2,
                                                     'RC'       : [
                                                            '01', '02', '03',\
                                                            '04', '05', '06',\
                                                            '07', '09', '12',\
                                                            '13', '14', '15',\
                                                            '17', '18', '19',\
                                                            '20', '21', '22',\
                                                            '23', '24', '25',\
                                                            '26', '27', '28',\
                                                            '29', '30', '31',\
                                                            '33', '34', '35',\
                                                            '36', '37', '38',\
                                                            '39', '40', '41',\
                                                            '42', '43', '44',\
                                                            '45', '46', '47',\
                                                            '48', '49', '50',\
                                                            '51', '52', '53',\
                                                            '54', '55', '56',\
                                                            '57', '58', '59',\
                                                            '60', '61', '62',\
                                                            '63', '64', '65',\
                                                            '66', '67', '68',\
                                                            '69', '70', '71',\
                                                            '72', '73', '74',\
                                                            '75', '76', '77',\
                                                            '78', '79', '80',\
                                                            '81', '82', '83',\
                                                            '84', '85', '86',\
                                                            '87', '88', '89',\
                                                            '90', '91', '92',\
                                                            '93', '94', '95',\
                                                            '96', '97', '98',\
                                                            '99'
                                                                  ]},
                u"проблемные операции"            : {'order'    : 3,
                                                     'RC'       : [
                                                            '12', '22', '63',\
                                                            '88', ''
                                                                  ]},
                u"пополнение"                     : {'order'     : 4,
                                                     'exclude'   : 'pstOnly',
                                                     'Operation' : ['8', 'P']},
                u"оплата услуг"                   : {'order'     : 5,
                                                     'typeLuno'  : 'pos',
                                                     'Operation' : ['1']},
                u"снятие наличных"                : {'order'     : 6,
                                                     'exclude'   : 'pstOnly',
                                                     'typeLuno'  : 'pvn',
                                                     'Operation' : ['1']},
                u"ануляция"                       : {'order'     : 7,
                                                     'exclude'   : 'pstOnly',
                                                     'Operation' : ['2']},
                u"автосторнирование"              : {'order'     : 8,
                                                     'exclude'   : 'pstOnly',
                                                     'Operation' : ['O']},
                u"ошибка шифрования (88)"         : {'order'     : 9,
                                                     'RC'        : ['88']}}


#POS - Оплата услуг
#PVN - Выдача наличных, пополнение


# подключение диалогов
import pmcdlgops    # диалог выбора ОПС
import pmcdlgwp     # диалог выбора РМ

###############################################################################
# Начало класса элемента интерфейса (сетка) "Результат проверки" (GridPanel)
#
class GridPanel(wx.Panel):
    """
    Grid to display the query result
    """
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)

        ### > названия и ширина колонок
        colNamesSizes = [
                        (u"Дата",             70),
                        (u"Время",            60),
                        (u"ОПС",             130),
                        (u"РМ",              100),
                        (u"pos",              40),
                        (u"pvn",              40),
                        (u"Операция",        140),
                        (u"Сумма",            80),
                        (u"Хост",            120),
                        (u"Ответ хоста",     170),
                        (u"Код авторизации", 150)     # код авторизации или REF-номер
                        ]
        ###

        ### > инициализация, начальные установки сетки
        self.grid = wx.grid.Grid(self, -1, size=(-1, -1))
        self.grid.CreateGrid(0, len(colNamesSizes))
        self.grid.EnableEditing(0)          # колонки не редактировать
        self.grid.EnableDragGridSize(0)
        self.grid.SetRowLabelSize(50)       # ширина колонки с номером
        self.grid.SetColLabelSize(20)       # высота строки с названиями столбцов
        self.grid.SetFocus()
        ###

        ### > создание сетки
        for colNumber, colName in enumerate(colNamesSizes):
            self.grid.SetColLabelValue(colNumber, colName[0])
            self.grid.SetColSize(colNumber, colName[1])
        ###

        ### > поместить сетку на сайзерах
        box = wx.StaticBox(self, -1, u"Операции:")
        bsizer = wx.StaticBoxSizer(box, wx.VERTICAL)
        bsizer.Add(self.grid, 1, wx.ALL|wx.EXPAND, 5)

        border = wx.BoxSizer()
        border.Add(bsizer, 1, wx.EXPAND|wx.LEFT|wx.RIGHT, 10)

        self.SetSizer(border)
        self.Layout()
        ###
#
# Конец класса элемента интерфейса GridPanel
###############################################################################


###############################################################################
# Начало класса элемента интерфейса "По ОПС:" (OpsPanel)
#
class OpsPanel(wx.Panel):
    """
    Choice of post offices
    """
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)

        box = wx.StaticBox(self, -1, u"По ОПС:")
        bsizer = wx.StaticBoxSizer(box, wx.VERTICAL)
        hsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        hsizer2 = wx.BoxSizer(wx.HORIZONTAL)

        ### > первоначальная инициализация переменных
        self.allOPSNameList = []
        self.opsSelectList  = []
        self.opsWPNameList  = []
        self.wpSelectList   = []
        ###

        ### > картинки
        self.pngChoiceOK = wx.Bitmap(os.path.join(picsDir, 'choiceOK.png'), wx.BITMAP_TYPE_PNG)
        self.pngChoiceNO = wx.Bitmap(os.path.join(picsDir, 'choiceNO.png'), wx.BITMAP_TYPE_PNG)
        self.stBmp_1 = wx.StaticBitmap(self, -1, self.pngChoiceNO)
        self.stBmp_2 = wx.StaticBitmap(self, -1, self.pngChoiceNO)
        ###

        ### > виджеты
        self.checkBox = wx.CheckBox(self, -1, u"все")
        self.selectOPSButton = wx.Button(self, -1, label=u"Выбрать ОПС", size=(100, 24))
        self.selectWPButton = wx.Button(self, -1, label=u"Выбрать РМ", size=(100, 24))

        self.selectOPSButton.SetToolTipString(u"Выберите одно или несколько ОПС")
        self.selectWPButton.SetToolTipString(u"Выберите одно или несколько рабочих мест в ОПС")
        ###


        ### > на горизонтальные сайзеры кнопки и картинки
        hsizer1.Add(self.selectOPSButton, 0, wx.LEFT|wx.RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        hsizer1.Add(self.stBmp_1, 0, wx.LEFT|wx.ALIGN_CENTER_VERTICAL, 5)

        hsizer2.Add(self.selectWPButton, 0, wx.LEFT|wx.RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        hsizer2.Add(self.stBmp_2, 0, wx.LEFT|wx.ALIGN_CENTER_VERTICAL, 5)
        ###

        ### > чекбокс и горизонтальные сайзеры на вертикальный сайзер
        bsizer.Add(self.checkBox, 0, wx.ALL|wx.EXPAND, 5)
        bsizer.Add(hsizer1, 0, wx.EXPAND, 5)
        bsizer.Add(hsizer2, 1, wx.EXPAND, 5)
        ###

        ### > связыватие событий
        self.checkBox.Bind(wx.EVT_CHECKBOX, self.OnToggleCheckBox)
        self.selectOPSButton.Bind(wx.EVT_BUTTON, self.OnSelectOPS)
        self.selectWPButton.Bind(wx.EVT_BUTTON, self.OnSelectWP)
        ###

        ### > первоначальное состояние виджетов
        self.checkBox.SetValue(True)
        self.selectOPSButton.Disable()
        self.selectWPButton.Disable()
        self.stBmp_1.Disable()
        self.stBmp_2.Disable()
        ###

        border = wx.BoxSizer()
        border.Add(bsizer, 1, wx.EXPAND|wx.ALL, 10)
        self.SetSizer(border)
        #border.Fit(self)
        self.Layout()


    # --- метод генерирует подсказку по выбранным ОПС -------------------------
    def SetForOPSToolTip(self):
        opsSelNameList = [self.allOPSNameList[i] for i in self.opsSelectList]

        if opsSelNameList == []:
            toolTipStr = u"Ничего не выбрано"
        elif len(opsSelNameList) == 1:
            toolTipStr = u"Отделение почтовой связи:\n      " + opsSelNameList[0]
        else:
            toolTipStr = u"Отделения почтовой связи:"
            if len(opsSelNameList) == len(self.allOPSNameList):
                toolTipStr += "\n      все"
            else:
                for ops in opsSelNameList:
                    toolTipStr += "\n      " + ops
        self.stBmp_1.SetToolTipString(toolTipStr)


    # --- метод генерирует подсказку по выбранным РМ --------------------------
    def SetForWPToolTip(self):
        wpSelNameList  = [self.opsWPNameList[i] for i in self.wpSelectList]

        if wpSelNameList == []:
            toolTipStr = u"Все рабочие места"
        elif len(wpSelNameList) == len(self.opsWPNameList):
            toolTipStr = u"Все рабочие места"
        elif len(wpSelNameList) == 1:
            toolTipStr = u"Рабочее место:\n      " + wpSelNameList[0]
        else:
            toolTipStr = u"Рабочие места:"
            for wp in wpSelNameList:
                toolTipStr += "\n      " + wp
        self.stBmp_2.SetToolTipString(toolTipStr)


    # --- обработка события изменения состояния чекбокса ----------------------
    def OnToggleCheckBox(self, event):
        if not event.GetInt():
            ### > формирование списка ОПС с установленными БПК
            self.allOPSNameList = []
            for key in bpkDict:
                nameOPS = opsDict[bpkDict[key].get('instPlace')].get('nameOPS')
                if not (nameOPS in self.allOPSNameList):
                    self.allOPSNameList.append(nameOPS)
            self.allOPSNameList = sorted(self.allOPSNameList)
            ###

            ### > переинициализация переменных
            self.opsSelectList = []
            self.opsWPNameList = []
            self.wpSelectList  = []
            global typeOfDevice
            typeOfDevice = 'all'
            ###

            ### > засветка элементов интерфейса
            self.selectOPSButton.Enable()
            self.stBmp_1.Enable()
            ###

            ### > формирование подсказки по ОПС
            self.SetForOPSToolTip()
            ###
        else:
            ### > переинициализация переменных
            self.allOPSNameList = []
            self.opsSelectList  = []
            self.opsWPNameList  = []
            self.wpSelectList   = []
            ###

            self.selectOPSButton.Disable()
            self.selectWPButton.Disable()
            self.stBmp_1.Disable()
            self.stBmp_1.SetBitmap(self.pngChoiceNO)
            self.stBmp_2.SetBitmap(self.pngChoiceNO)


    # --- обработка события нажатия кнопки "Выбрать ОПС" ----------------------
    def OnSelectOPS(self, event):
        ### > первоначальное состояние переменных и виджетов
        self.opsWPNameList  = []
        self.wpSelectList   = []
        global typeOfDevice
        typeOfDevice = 'all'
        self.selectWPButton.Disable()
        self.stBmp_2.SetBitmap(self.pngChoiceNO)
        ###

        ### > отображение диалога со списком всех ОПС
        dlgOPS = pmcdlgops.SelectOPS(self.allOPSNameList, self.opsSelectList, OpsPanel)
        if (dlgOPS.ShowModal() == wx.ID_OK):
            #self.opsSelectList = list(dlgOPS.checkListBox.GetChecked())
                        
            ### > преобразование имени ОПС в порядковый номер ОПС из списка self.allOPSNameList
            self.opsSelectList = []
            for number, name in enumerate(self.allOPSNameList):
                if name in dlgOPS.selectedOPSNameList:
                    self.opsSelectList.append(number)
            ###
        dlgOPS.Destroy()
        ###

        ### > отображение галочки, если хоть одно ОПС выбрано
        if self.opsSelectList != []:
            self.stBmp_1.SetBitmap(self.pngChoiceOK)
        else:
            self.stBmp_1.SetBitmap(self.pngChoiceNO)
            self.selectWPButton.Disable()
        ###

        ### > формирование подсказки по ОПС
        self.SetForOPSToolTip()
        ###

        if len(self.opsSelectList) == 1:
            ### > получить индекс ОПС по имени ОПС
            nameOPS = self.allOPSNameList[self.opsSelectList[0]]
            for index in opsDict:
                if opsDict[index].get('nameOPS') == nameOPS:
                    indexOPS = index
                    break
            ###

            ### > получить РМ отделения по индексу
            self.opsWPNameList = []
            for key in bpkDict:
                if bpkDict[key].get('instPlace') == indexOPS:
                    self.opsWPNameList.append(bpkDict[key].get('instPoint'))
            self.opsWPNameList = sorted(self.opsWPNameList)
            ###

            '''
            for x in self.opsWPNameList:
                print x
            #
            wx.MessageBox(str(self.opsWPNameList),
                            u"Отладка", wx.OK | wx.ICON_INFORMATION, self)
            #'''


            ### > если РМ больше одного, то засветить кнопку
            if len(self.opsWPNameList) > 1:
                self.selectWPButton.Enable()
                self.stBmp_2.Enable()

                ### > формирование подсказки по РМ
                self.SetForWPToolTip()
                ###
            else:
                self.selectWPButton.Disable()
            ###
        else:
            self.opsWPNameList = []
            self.wpSelectList  = []

            self.selectWPButton.Disable()
            self.stBmp_2.SetBitmap(self.pngChoiceNO)


    # --- обработка события нажатия кнопки "Выбрать РМ" -----------------------
    def OnSelectWP(self, event):
        ### > отображение диалога со списком всех РМ в ОПС
        dlgWP = pmcdlgwp.SelectWP(self.opsWPNameList, self.wpSelectList, OpsPanel)
        if (dlgWP.ShowModal() == wx.ID_OK):
            self.wpSelectList = list(dlgWP.checkListBox.GetChecked())
        dlgWP.Destroy()
        ###

        ### > определение типа устройств
        opsSelNameList = [self.allOPSNameList[i] for i in self.opsSelectList]
        wpSelNameList  = [self.opsWPNameList[i] for i in self.wpSelectList]

        # > получить список индексов ОПС из списка имён ОПС
        opsSelIndexList = []
        for nameOPS in opsSelNameList:
            for indexOPS in opsDict:
                if opsDict[indexOPS].get('nameOPS') == nameOPS:
                    opsSelIndexList.append(indexOPS)
        #

        # > подсчитать количество выбранных ПСТ в данном ОПС
        countPST = 0
        for key in bpkDict:
            if bpkDict[key].get('instPlace') == opsSelIndexList[0] and\
                (bpkDict[key].get('instPoint') in wpSelNameList):
                if bpkDict[key].get('pvn') == 'None':
                    countPST +=1
        #

        # > если выбранные РМ только ПСТ, то изменить переменную typeOfDevice
        global typeOfDevice
        if len(wpSelNameList) == countPST:
            typeOfDevice = 'pstOnly'
            #AppMainFrame.filterPanel.SetFilterPST()
        else:
            typeOfDevice = 'all'
        #
        ###

        ### > отображение галочки, если хоть одно РМ выбрано
        if len(self.wpSelectList) != 0:
            self.stBmp_2.SetBitmap(self.pngChoiceOK)

            ### > формирование подсказки по РМ
            self.SetForWPToolTip()
            ###
        else:
            self.stBmp_2.SetBitmap(self.pngChoiceNO)
        ###
#
# Конец класса элемента интерфейса OpsPanel
###############################################################################


###############################################################################
# Начало класса элемента интерфейса "За период:" (PeriodPanel)
#
class PeriodPanel(wx.Panel):
    """
    Select a template of the payment period
    """
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)

        box = wx.StaticBox(self, -1, u"За период:")
        bsizer = wx.StaticBoxSizer(box, wx.VERTICAL)
        hsizer = wx.BoxSizer(wx.HORIZONTAL)

        ### > виджеты
        self.checkBox = wx.CheckBox(self, -1, u"задать диапазон дат")
        self.label1 = wx.StaticText(self, -1, u"с")
        self.label2 = wx.StaticText(self, -1, u"по")
        self.dateBefore = wx.DatePickerCtrl(self, -1,
                        style=wx.DP_DROPDOWN|wx.DP_SHOWCENTURY)
        self.dateBefore.SetToolTipString(u"Начальная дата, включительно")
        self.dateLater  = wx.DatePickerCtrl(self, -1,
                        style=wx.DP_DROPDOWN|wx.DP_SHOWCENTURY)
        self.dateLater.SetToolTipString(u"Конечная дата, включительно")
        self.comboBox = wx.ComboBox(self, 500, periodList[0], (90, 50),
                                 (180, -1), periodList,
                                 wx.CB_DROPDOWN|wx.CB_READONLY|wx.TE_PROCESS_ENTER)
        self.comboBox.SetToolTipString(u"Период по шаблону")
        ###

        ### > хранит предыдущую дату конца диапазона
        self.prevEndDateRange = self.dateLater.GetValue()
        ###

        ### > диапазон дат на горизонтальный сайзер
        hsizer.Add(self.label1, 0, wx.RIGHT|wx.ALIGN_CENTER_VERTICAL, 3)
        hsizer.Add(self.dateBefore, 0, wx.ALIGN_CENTER_VERTICAL, 0)
        hsizer.Add(self.label2, 0, wx.LEFT|wx.RIGHT|wx.ALIGN_CENTER_VERTICAL, 3)
        hsizer.Add(self.dateLater, 0, wx.ALIGN_CENTER_VERTICAL, 0)
        ###

        ### > горизонтальный сайзер и остальные виджеты дат на вертикальный сайзер
        bsizer.Add(self.checkBox, 0, wx.ALL|wx.EXPAND, 5)
        bsizer.Add(hsizer, 1, wx.ALL|wx.EXPAND, 5)
        bsizer.Add(self.comboBox, 0, wx.ALL|wx.EXPAND, 5)
        ###

        ### > события
        self.checkBox.Bind(wx.EVT_CHECKBOX,
                           self.OnToggleCheckBox)       # изменение состояния чекбокса
        self.dateBefore.Bind(wx.EVT_DATE_CHANGED,
                             self.OnDateBeforeChanged)  # изменение даты начала диапазона
        self.dateLater.Bind(wx.EVT_DATE_CHANGED,
                            self.OnDateLaterChanged)    # изменение даты конца диапазона
        ###        

        ### > первоначальное состояние виджетов
        self.label1.Disable()
        self.label2.Disable()
        self.dateBefore.Disable()
        self.dateLater.Disable()
        ###

        border = wx.BoxSizer()
        border.Add(bsizer, 1, wx.EXPAND|wx.ALL, 10)
        self.SetSizer(border)
        #border.Fit(self)
        self.Layout()


    # --- обработка события изменения состояния чекбокса -------------------------
    def OnToggleCheckBox(self, event):
        if event.GetInt():
            self.label1.Enable()
            self.label2.Enable()
            self.dateBefore.Enable()
            self.dateLater.Enable()
            self.comboBox.Disable()
        else:
            self.label1.Disable()
            self.label2.Disable()
            self.dateBefore.Disable()
            self.dateLater.Disable()
            self.comboBox.Enable()


    # --- обработка события изменения даты начала диапазона -------------------
    def OnDateBeforeChanged(self, event):
        dateBefore = event.GetDate()
        self.dateLater.SetValue(dateBefore)
        
        self.CheckModifiedDate(dateBefore)
        
        self.prevEndDateRange = self.dateLater.GetValue()


    # --- обработка события изменения даты конца диапазона --------------------
    def OnDateLaterChanged(self, event):
        dateBefore = self.dateBefore.GetValue()
        dateLater  = event.GetDate()
        
        self.CheckModifiedDate(dateLater)
        
        if dateBefore > dateLater:
            self.dateLater.SetValue(self.prevEndDateRange)
        else:
            self.prevEndDateRange = self.dateLater.GetValue()


    # --- проверка выставленной даты на предмет совпадения с текущей ----------
    def CheckModifiedDate(self, date):
        todayDate  = datetime.datetime.now().date()
        selectDate = datetime.date(date.GetYear(), date.GetMonth()+1, date.GetDay())
        if selectDate > todayDate:
            ### > заголовок окна
            title = u"Предупреждение"
            ###
            
            ### > показать информационное сообщение
            msg = u"Дата %s г. ещё не наступила: за эту дату не может быть операций!"\
                    % selectDate.strftime("%d.%m.%Y")
            msg += 10 * " "
            
            wx.MessageBox(msg,
                          title,
                          wx.OK | wx.ICON_EXCLAMATION,
                          self)
            ###
#
# Конец класса элемента интерфейса PeriodPanel
###############################################################################


###############################################################################
# Начало класса элемента интерфейса (комбобокс) "Фильтр операций:" (OprFilterPanel)
#
class OprFilterPanel(wx.Panel):
    """
    Sampling of the pattern
    """
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)

        box = wx.StaticBox(self, -1, u"Фильтр по операциям:")
        bsizer = wx.StaticBoxSizer(box, wx.VERTICAL)

        ### > получить список фильтров для всех устройств
        self.allFilterList = sorted([filterDict[i].get('order') for i in filterDict.keys()])
        for filter in filterDict.keys():
            self.allFilterList[filterDict[filter].get('order')] = filter
        ###

        ### > получить список фильтров для ПСТ
        self.pstFilterList = [i for i in self.allFilterList if filterDict[i].get('exclude') != 'pstOnly']
        ###

        self.comboBox = wx.ComboBox(self, 500, self.allFilterList[0], (90, 50),
                                 (200, -1), self.allFilterList,
                                 wx.CB_DROPDOWN|wx.CB_READONLY|wx.TE_PROCESS_ENTER)
        self.comboBox.SetToolTipString(u"Выберите шаблон проверки")
        #self.comboBox.Bind(wx.EVT_COMBOBOX, self.OnSetFocus)

        bsizer.Add(self.comboBox, 0, wx.ALL|wx.EXPAND, 4)

        border = wx.BoxSizer()
        border.Add(bsizer, 1, wx.EXPAND|wx.RIGHT|wx.LEFT|wx.TOP, 10)
        self.SetSizer(border)
        #border.Fit(self)
        self.Layout()


    def OnSetFocus(self, event):
        def setList(list):
            self.comboBox.Clear()
            for item in list:
                self.comboBox.Append(item)

        if typeOfDevice == 'all':
            self.dinamicList = self.allFilterList
        elif typeOfDevice == 'pstOnly':
            self.dinamicList = self.pstFilterList

        setList(self.dinamicList)
        self.comboBox.SetSelection(0)


    #@classmethod
    #@staticmethod
    def SetFilterAll(self):
        self.comboBox.Clear()
        for item in self.allFilterList:
            self.comboBox.Append(item)
        self.comboBox.SetSelection(0)


    def SetFilterPST(self):
        self.comboBox.Clear()
        for item in self.pstFilterList:
            self.comboBox.Append(item)
        self.comboBox.SetSelection(0)
#
# Конец класса элемента интерфейса OprFilterPanel
###############################################################################


###############################################################################
# Начало класса элемента интерфейса (комбобокс) "Фильтр хостов:" (HstFilterPanel)
#
class HstFilterPanel(wx.Panel):
    """
    Sampling of the pattern
    """
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)

        box = wx.StaticBox(self, -1, u"Фильтр по хостам:")
        bsizer = wx.StaticBoxSizer(box, wx.VERTICAL)

        ### > получить список всех хостов
        self.hostFilterList = sorted(hostDict.values())
        self.hostFilterList.insert(0, u"все хосты")
        ###

        self.comboBox = wx.ComboBox(self, 500, self.hostFilterList[0], (90, 50),
                                   (200, -1), self.hostFilterList,
                                   wx.CB_DROPDOWN|wx.CB_READONLY|wx.TE_PROCESS_ENTER)
        self.comboBox.SetToolTipString(u"Выберите хост")

        bsizer.Add(self.comboBox, 0, wx.ALL|wx.EXPAND, 4)

        border = wx.BoxSizer()
        border.Add(bsizer, 1, wx.EXPAND|wx.RIGHT|wx.LEFT|wx.BOTTOM, 10)
        self.SetSizer(border)
        #border.Fit(self)
        self.Layout()
#
# Конец класса элемента интерфейса HstFilterPanel
###############################################################################


###############################################################################
# Начало класса элемента интерфейса (прогресс) (GaugePanel)
#
class GaugePanel(wx.Panel):
    """
    Mapping process of the query
    """
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)

        box = wx.StaticBox(self, -1, u"")
        bsizer = wx.StaticBoxSizer(box, wx.VERTICAL)

        self.text = u" " * 70

        self.stText = wx.StaticText(self, -1, self.text, style=wx.ALIGN_CENTER)

        self.gauge = wx.Gauge(self, -1, 50, (-1, -1), (300, -1))
        self.gauge.SetBezelFace(5)
        self.gauge.SetShadowWidth(5)
        #self.gauge.Hide()

        ### > инициализация таймера (wxPython in action, стр. 569)
        self.timer = wx.Timer(self)
        ###

        ### > связыватие событий таймера
        self.Bind(wx.EVT_TIMER, self.OnTimer)
        ###

        bsizer.Add(self.stText, 0, wx.ALL|wx.ALIGN_CENTRE, 20)
        bsizer.Add(self.gauge, 0, wx.ALL|wx.EXPAND, 5)

        border = wx.BoxSizer()
        border.Add(bsizer, 1, wx.EXPAND|wx.ALL, 10)
        self.SetSizer(border)
        #border.Fit(self)
        self.Layout()


    # --- метод запуска "прогрессбара" ----------------------------------------
    def StartTimer(self):
        self.timer.Start(50)
        self.gauge.SetValue(0)
        self.stText.SetLabel(u"Выполняется запрос к БД. Ждите...")


    # --- метод остановки "прогрессбара" --------------------------------------
    def StopTimer(self):
        self.timer.Stop()
        self.gauge.ClearBackground()
        self.stText.SetLabel(u"Запрос к БД выполнен")
        #self.gauge.Refresh(True)
        #wx.Bell()


    # --- обработчик событий таймера ------------------------------------------
    def OnTimer(self, event):
        self.gauge.Pulse()
        #wx.CallLater(100, self.OnTimer)
        #wx.Yield()
#
# Конец класса элемента интерфейса GaugePanel
###############################################################################


###############################################################################
# Начало класса элемента интерфейса (кнопка) "Проверить" (ButtonPanel)
#
class ButtonPanel(wx.Panel):
    """
    Button to execute the query
    """
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)

        sizer = wx.BoxSizer(wx.VERTICAL)

        self.buttonQuery = wx.Button(self, -1, label=u"Выбрать", size=(100, 30))
        self.buttonQuery.SetToolTipString(u"Найти операции по указанным Вами критериям")

        self.buttonQuit = wx.Button(self, -1, label=u"Выход", size=(100, 30))
        self.buttonQuit.SetToolTipString(u"Выйти из программы")

        sizer.Add(self.buttonQuery, 0, wx.ALL|wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 10)
        sizer.Add(self.buttonQuit, 0, wx.ALL|wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 10)

        border = wx.BoxSizer()
        border.Add(sizer, 1, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 10)
        self.SetSizer(border)
        #border.Fit(self)
        self.Layout()
#
# Конец класса элемента интерфейса ButtonPanel
###############################################################################


###############################################################################
# Начало класса потока (формирование и выполнение запроса к БД) (ExecSQLThread)
#
class ExecSQLThread(Thread):
    """
    Running a query in a separate thread
    """
    def __init__(self, queryDB):
        self.queryDB = queryDB
        Thread.__init__(self)
        self.daemon = True
        self.start()            # запустить поток


    # --- выполняется запрос к БД в отдельном потоке --------------------------
    def run(self):
        try:
            cardDB = MySQLdb.connect(dbHost, dbUser, dbPasswd, dbName)
            cursor = cardDB.cursor()
            cursor.execute(self.queryDB)
            data = cursor.fetchall()
            cursor.close()
            cardDB.close()
            ### > отправить полученные из БД данные
            wx.CallAfter(Publisher().sendMessage, "dat_msg", data)
            ###
        except:
            ### > отправить сообщение об ошибке
            wx.CallAfter(Publisher().sendMessage, "err_msg",\
                        u"Не удалось соединиться с базой данных!!!")
            ###
#
# Конец класса потока ExecSQLThread
###############################################################################


###############################################################################
# Начало класса создания "прогрессбара" для строки состояния (ProgressStatusBar)
#
class ProgressStatusBar:
    """
    Mapping process of the query
    """
    # --- конструктор класса --------------------------------------------------
    def __init__ (self, parent, statusbar, sbarfields=1, sbarfield=0, maxcount=50):
        rect = statusbar.GetFieldRect(sbarfield)
        barposn = (rect[0], rect[1])

        if sbarfield + 1 == sbarfields and 'wxMSW' in wx.PlatformInfo:
        #if (sbarfield + 1 == sbarfields) and (wx.Platform == '__WXMSW__'):
            barsize = (rect[2] + 35, rect[3])   # полностью заполнить поле
        else:
            barsize = (rect[2], rect[3])

        self.progressBar = wx.Gauge(statusbar,
                                    -1,
                                    maxcount,
                                    barposn,
                                    barsize
                                    #style=wx.GA_SMOOTH
                                    )

        self.progressBar.SetBezelFace(3)
        self.progressBar.SetShadowWidth(3)
        self.progressBar.Hide()

        self.progressTimer = wx.PyTimer(self.OnUpdProgress)
        self.OnUpdProgress()


    # --- деструктор класса ---------------------------------------------------
    def __del__(self):
        if self.progressTimer.IsRunning():
            self.progressTimer.Stop()


    # --- метод обновления "прогрессбара" -------------------------------------
    def OnUpdProgress(self):
        self.progressBar.Pulse()


    # --- метод запуска "прогрессбара" ----------------------------------------
    def Start(self):
        self.progressTimer.Start(25) # запустить счётчик интервала приращения, мс
        self.progressBar.Show(True)


    # --- метод остановки "прогрессбара" --------------------------------------
    def Stop(self):
        self.progressBar.Show(False)
        self.progressTimer.Stop()
        #wx.Bell()
#
# Конец класса ProgressStatusBar
###############################################################################


###############################################################################
# Начало класса основного окна программы (AppMainFrame)
#
class AppMainFrame(wx.Frame):
    """
    The main application form
    """
    # --- конструктор формы ---------------------------------------------------
    def __init__(self, *args, **kwds):
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        super(AppMainFrame, self).__init__(*args, **kwds)
        #wx.Frame.__init__(self, *args, **kwds) # либо так
        self.SetTitle(windowTitle)

        self.queryDB = ''       # в самом начале строка запроса пуста

        ### > иконка приложения
        self.SetIcon(wx.IconFromBitmap(wx.Bitmap(applicationIcon)))
        ###

        ### > при закрытии окна программы 'крестиком' будет вызывться метод 'OnClose'
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        ###

        ### > создать экземпляры классов элементов интерфейса
        self.gridPanel      = GridPanel(self)
        self.opsPanel       = OpsPanel(self)
        self.periodPanel    = PeriodPanel(self)
        #self.searchPanel    = SearchPanel(self)
        self.oprFilterPanel = OprFilterPanel(self)
        self.hstFilterPanel = HstFilterPanel(self)
        self.gaugePanel     = GaugePanel(self)
        self.buttonPanel    = ButtonPanel(self)
        ###

        ### > создать сайзеры
        szr_1 = wx.BoxSizer(wx.VERTICAL)
        szr_2 = wx.BoxSizer(wx.HORIZONTAL)
        szr_3 = wx.BoxSizer(wx.VERTICAL)
        ###

        ### > связать событя нажатия кнопок
        self.buttonPanel.buttonQuery.Bind(wx.EVT_BUTTON, self.OnQuery)
        self.buttonPanel.buttonQuit.Bind(wx.EVT_BUTTON, self.OnQuit)
        ###

        ### > разместить экземпляры классов на сайзерах
        szr_1.Add(self.gridPanel, 6, wx.ALL|wx.EXPAND, 0)
        szr_1.Add(szr_2, 1, wx.ALL|wx.EXPAND, 0)
        szr_2.Add(self.opsPanel, 0, wx.ALL|wx.EXPAND, 0)
        szr_2.Add(self.periodPanel, 0, wx.ALL|wx.EXPAND, 0)
        szr_2.Add(szr_3, 0, wx.ALL|wx.EXPAND, 0)
        szr_2.Add(self.gaugePanel, 0, wx.ALL|wx.EXPAND, 0)
        szr_2.Add(wx.Size(), 0, wx.EXPAND, 0)
        szr_2.Add(self.buttonPanel, proportion=1, flag=wx.EXPAND, border=0)

        szr_3.Add(self.oprFilterPanel, 1, wx.ALL|wx.EXPAND, 0)
        szr_3.Add(self.hstFilterPanel, 1, wx.ALL|wx.EXPAND, 0)

        self.SetSizer(szr_1)
        szr_1.Fit(self)
        self.SetMinSize((1240, 916))
        self.Layout()
        ###

        ### > разместить строку меню в интерфейсе
        self.MakeMenuBar()
        ###

        ### > разместить строку состояния в интерфейсе
        self.MakeStatusBar()
        ###

        #self.Centre()                               # центрировать окно
        self.CenterOnScreen()
        self.Show(True)
        self.buttonPanel.buttonQuery.SetFocus()     # фокус на кнопку

        ### > создать приёмники сообщений из потока
        Publisher().subscribe(self.DataReceiver,  "dat_msg")  # зпрашиваемые данные
        Publisher().subscribe(self.ErrorReceiver, "err_msg")  # об ошибке
        ###


    # --- создание строки меню ------------------------------------------------
    def MakeMenuBar(self):
        menuBar  = wx.MenuBar()

        menuFile   = wx.Menu()
        menuEdit   = wx.Menu()
        menuView   = wx.Menu()
        menuSearch = wx.Menu()
        menuHelp   = wx.Menu()

        menuBar.Append(menuFile,   u"&Файл")
        menuBar.Append(menuEdit,   u"&Правка")
        menuBar.Append(menuView,   u"&Вид")
        menuBar.Append(menuSearch, u"П&оиск")
        menuBar.Append(menuHelp,   u"&Справка")

        ### > пункт меню "Файл" -- "Сохранить"
        menuFileSave = wx.MenuItem(menuFile, wx.NewId(), u"&Сохранить операции")
        menuFileSave.SetHelp(u"Сохратить операции по БПК на текущий момент")
        menuFileSave.SetBitmap(wx.Bitmap(os.path.join(picsDir, 'save.png')))
        menuFile.AppendItem(menuFileSave)
        self.Bind(wx.EVT_MENU, self.OnSave, menuFileSave)
        ###

        menuFile.AppendSeparator()

        ### > пункт меню "Файл" -- "Выход"
        menuFileQuit = wx.MenuItem(menuFile, wx.NewId(), u"&Выход")
        menuFileQuit.SetHelp(u"Выйти из программы")
        menuFileQuit.SetBitmap(wx.Bitmap(os.path.join(picsDir, 'quit.png')))
        menuFile.AppendItem(menuFileQuit)
        self.Bind(wx.EVT_MENU, self.OnQuit, menuFileQuit)
        ###

        ### > пункт меню "Правка" -- "Последний SQL-запрос к БД"
        menuEditLastQueryDB = wx.MenuItem(menuEdit, wx.NewId(), u"&Последний SQL-запрос к БД")
        menuEditLastQueryDB.SetHelp(u"Показать последний SQL-запрос к базе данных")
        menuEditLastQueryDB.SetBitmap(wx.Bitmap(os.path.join(picsDir, 'lastsql.png')))
        menuEdit.AppendItem(menuEditLastQueryDB)
        self.Bind(wx.EVT_MENU, self.OnLastQueryDB, menuEditLastQueryDB)
        ###

        ### > пункт меню "Поиск" -- "Найти операцию"
        menuSearchFind = wx.MenuItem(menuSearch, wx.NewId(), u"&Найти операцию...")
        menuSearchFind.SetHelp(u"Найти операцию в БД по коду авторизации на чеке")
        menuSearchFind.SetBitmap(wx.Bitmap(os.path.join(picsDir, 'search.png')))
        menuSearch.AppendItem(menuSearchFind)
        self.Bind(wx.EVT_MENU, self.OnFind, menuSearchFind)
        ###


        ### > пункт меню "Справка" -- "О программе..."
        menuHelpAbout = wx.MenuItem(menuHelp, wx.NewId(), u"&О программе...")
        menuHelpAbout.SetHelp(u"Краткое описание программы")
        menuHelpAbout.SetBitmap(wx.Bitmap(os.path.join(picsDir, 'about.png')))
        menuHelp.AppendItem(menuHelpAbout)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuHelpAbout)
        ###

        ### > установить строку меню
        self.SetMenuBar(menuBar)
        ###


    # --- создание строки строки состояния ------------------------------------
    def MakeStatusBar(self):
        #   == назначение полей ==
        #    0 - поле сообщений
        #    1 - прогрессбар
        #    2 - текущая дата
        #    3 - текущее время
        #    4 - количество операций
        #    5 - сумма по всем операциям

        ### > создать строку состояния
        self.sb = wx.StatusBar(self, wx.ID_ANY)
        ###

        ### > установить количество полей
        self.sb.SetFieldsCount(6)
        ###

        ### > ширина каждого поля
        self.sb.SetStatusWidths([355, 300, 70, 70, 130, -1])
        ###

        ### > установить строку состояния
        self.SetStatusBar(self.sb)
        ###

        ### > поместить текст в поля строки состояния
        self.sb.SetStatusText(u"Готов!", 0)
        self.sb.SetStatusText(time.strftime("%d.%m.%Y", time.localtime(time.time())), 2)
        self.sb.SetStatusText(u"Операций: - -", 4)
        self.sb.SetStatusText(u"Сумма: - - руб.", 5)
        ###

        ### > создать "прогрессбар" в строке состояния
        self.progressStatusBar = ProgressStatusBar(self, self.sb, 6, 1)
        ###

        ### > создать таймер для часов в строке состояния
        self.clockTimer = wx.PyTimer(self.OnUpdClock)
        ###

        ### > обновление каждые 1000 миллисекунд
        self.clockTimer.Start(1000)
        ###

        self.OnUpdClock()


    # --- обработчик обновления времени в строке состояния --------------------
    def OnUpdClock(self):
        t = time.localtime(time.time())
        st = time.strftime("%H:%M:%S", t)
        self.sb.SetStatusText(st, 3)


    # -->>>---- О Б Р А Б О Т Ч И К И   В Ы Б О Р А   П У Н К Т О В   М Е Н Ю
    #
    # --- метод вызова окна "Найти операцию..." -------------------------------
    def OnFind(self, event):
        ### > заголовок окна
        title = u"Найти операцию БД по коду авторизации"
        ###
        
        ### > сообщение на форме
        msg = u"Введите ref-код с фискального чека"
        ###
        
        while True:
            dlgFind = wx.TextEntryDialog(self,
                                         message=msg,
                                         caption=title)
            
            if dlgFind.ShowModal() == wx.ID_OK:
                refCode = dlgFind.GetValue()
                
                if refCode == '':
                    ### > показать информационное сообщение
                    infMsg = u'Введите код авторизации, состоящий из 12-и цифр!'
                    infMsg += 10 * ' '
                    
                    wx.MessageBox(infMsg,
                                  title,
                                  wx.OK | wx.ICON_EXCLAMATION,
                                  self)
                    ###
                    continue
                elif not len(refCode) == 12 or not refCode.isdigit():
                    ### > показать информационное сообщение
                    infMsg = u'Код авторизации должен состоять из 12-и цифр!'
                    infMsg += 10 * ' '
                    infMsg += '\n'
                    infMsg += u'Повторите ввод.'
                    
                    wx.MessageBox(infMsg,
                                  title,
                                  wx.OK | wx.ICON_EXCLAMATION,
                                  self)
                    ###
                    dlgFind.SetValue(wx.EmptyString)
                    continue
                else:
                    ### > сформировать и выполнить запрос
                    self.ExecSQL(refCode, event)
                    self.sb.SetStatusText(u"Операций: - -", 4)
                    self.sb.SetStatusText(u"Сумма: - - руб.", 5)
                    ###
                    break
            else:
                break
        dlgFind.Destroy()


    # --- метод вызова окна "Последний SQL-запрос к БД" -----------------------
    def OnLastQueryDB(self, event):
        import wx.lib.dialogs
        ### > заголовок окна
        title = u"Последний SQL-запрос к БД"
        ###
        
        if self.queryDB == '':
            ### > показать информационное сообщение
            infMsg = u'Сначала сформируйте запрос, нажав на кнопку "Выбрать"!'
            infMsg += 10 * ' '
            infMsg += '\n'
            infMsg += u'При необходимости задайте критерии выборки.'
            
            wx.MessageBox(infMsg,
                          title,
                          wx.OK | wx.ICON_EXCLAMATION,
                          self)
            ###
        else:
            ### > показать последний SQL-запрос к БД в окне
            dlgLastQueryDB = wx.lib.dialogs.ScrolledMessageDialog(self,
                                                                  self.queryDB,
                                                                  title)
            dlgLastQueryDB.ShowModal()
            ###


    # --- метод вызова окна "Сохранить операции" ------------------------------
    def OnSave(self, event):
        ### > заголовок окна
        title = u"Сохратить операции по БПК на текущий момент"
        ###
        
        if self.gridPanel.grid.GetNumberRows() == 0:
            ### > показать информационное сообщение
            infMsg = u'Сначала выберите желаемые операции, нажав на кнопку "Выбрать"!'
            infMsg += 10 * ' '
            infMsg += '\n'
            infMsg += u'При необходимости задайте критерии выборки.'
            
            wx.MessageBox(infMsg,
                          title,
                          wx.OK | wx.ICON_EXCLAMATION,
                          self)
            ###
        else:
            import csv
            try:
                import xlwt

                ### > список фильтров
                wildcard = u"Книга MS Office Excel (*.xls)|*.xls|"     \
                           u"Файл CSV, разделители точка с запятой (*.csv)|*.csv"
                ###
            except ImportError:
                ### > список фильтров
                wildcard = u"Файл CSV, разделители точка с запятой (*.csv)|*.csv"
                ###

                ### > показать информационное сообщение
                infMsg = u'Не установлен модуль записи в Excel-файл xlwt!'
                infMsg += 10 * ' '

                wx.MessageBox(infMsg,
                              title,
                              wx.OK | wx.ICON_EXCLAMATION,
                              self)
                ### > показать информационное сообщение

            ### > генерируем имя файла
            prefixName = "PMCreport_"
            dtStamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            svFileName = prefixName + dtStamp
            ###

            ### > регулярные выражения даты, времени, суммы платежа
            datePattern   = compile(r'^\d{1,2}?.?\d{2}?.?\d{4}$')
            timePattern   = compile(r'^\d{1,2}?:?\d{2}?:?\d{2}$')
            amountPattern = compile(r'\d{3}') # r'^\d{1,3}? ?\d{3}'
            ###

            ### > описание диалога сохранения
            saveDlg = wx.FileDialog(self,
                                    message     = title,
                                    defaultDir  = reportDir,
                                    defaultFile = svFileName,
                                    wildcard    = wildcard,
                                    style       = wx.SAVE)
            saveDlg.SetFilterIndex(0)
            ###

            if saveDlg.ShowModal() == wx.ID_OK:
                path = saveDlg.GetPath()
                if os.path.splitext(path)[1] == '.xls':
                    """ сохранение в Excel-файл """
                    book = xlwt.Workbook(encoding='utf-8')
                    sheet = book.add_sheet(u'Операции по БПК')

                    ### > обрамление тонкой линией в ячейке
                    borders        = xlwt.Borders()
                    borders.left   = xlwt.Borders.THIN
                    borders.right  = xlwt.Borders.THIN
                    borders.top    = xlwt.Borders.THIN
                    borders.bottom = xlwt.Borders.THIN
                    ###

                    #pattern = xlwt.Pattern()
                    #pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                    #pattern.pattern_fore_colour = 0x0A

                    ### > стиль отбражения даты в Excel-файле
                    dateStyle = xlwt.XFStyle()
                    dateStyle.num_format_str = 'D.MM.YYYY'
                    dateStyle.borders = borders
                    ###

                    ### > стиль отбражения времени в Excel-файле
                    timeStyle = xlwt.XFStyle()
                    timeStyle.num_format_str = 'h:mm:ss'
                    timeStyle.borders = borders
                    ###

                    ### > обычный стиль отбражения в Excel-файле
                    regularStyle = xlwt.XFStyle()
                    regularStyle.borders = borders
                    ###

                    #print self.gridPanel.grid.GetColLabelSize()
                    ### > установить ширину колонок в Excel-файле
                    for col in range(self.gridPanel.grid.GetNumberCols()):
                        colSize = self.gridPanel.grid.GetColSize(col)
                        sheet.col(col).width = colSize * 42
                    ###

                    ### > выгрузить данные в Excel-файл
                    for row in range(self.gridPanel.grid.GetNumberRows()):
                        for col in range(self.gridPanel.grid.GetNumberCols()):
                            cellData = self.gridPanel.grid.GetCellValue(row, col)

                            ### > если в ячейке дата|время|число - обработать соответственно
                            if datePattern.search(cellData) is not None:
                                oDate = datetime.datetime.strptime(cellData, "%d.%m.%Y")
                                sheet.row(row).write(col, oDate, dateStyle)
                            elif timePattern.search(cellData) is not None:
                                tm = [int(i) for i in cellData.split(':')]
                                oTime = datetime.time(tm[0], tm[1], tm[2])
                                sheet.row(row).write(col, oTime, timeStyle)
                            elif amountPattern.search(cellData) is not None:
                                oAmount = Decimal(float(cellData.replace(' ', '')))
                                #sheet.row(row).write(col, oAmount)
                                sheet.row(row).set_cell_number(col, oAmount, regularStyle)
                            else:
                                sheet.row(row).write(col, cellData, regularStyle)
                            ###
                    ###
                    book.save(path)

                elif os.path.splitext(path)[1] == '.csv':
                    """ сохранение в CSV-файл """
                    ### > открытие csv-файла на запись
                    csvFile = open(path, 'wb')
                    writer = csv.writer(csvFile, delimiter=';',
                                quoting=csv.QUOTE_MINIMAL, lineterminator='\r\n')
                    ###

                    ### > выгрузить данные в csv-файл
                    for row in range(self.gridPanel.grid.GetNumberRows()):
                        rowData = []
                        for col in range(self.gridPanel.grid.GetNumberCols()):
                            cellData = self.gridPanel.grid.GetCellValue(row, col)

                            ### > удрать пробелы из чисел
                            if amountPattern.search(cellData) is not None:
                                cellData = cellData.replace(' ', '')
                            ###

                            rowData.append(cellData.encode('cp1251'))
                        writer.writerow(rowData)
                    ###
                    csvFile.close()

            saveDlg.Destroy()


    # --- метод вызова окна "О программе..." ----------------------------------
    def OnAbout(self, event):
        import wx.lib.dialogs

        ### > загрузить текст описания из файла
        aboutFile = open(os.path.join(startDir, 'about.txt'), 'r')
        aboutText = aboutFile.read()
        aboutFile.close()
        ###

        ### > показать текст описания в окне
        dlgAbout = wx.lib.dialogs.ScrolledMessageDialog(self,
                                                        aboutText.decode('utf-8'),
                                                        u"О программе...")
        dlgAbout.ShowModal()
        ###



    # -->>>---- О Б Р А Б О Т Ч И К И   Н А Ж А Т И Я   В И Д Ж Е Т О В
    #
    # --- выход из программы --------------------------------------------------
    def OnQuit(self, event):
        self.Show(False)
        self.Close()


    # --- закрытие формы "крестиком" ------------------------------------------
    def OnCloseWindow(self, event):
        self.Show(False)
        self.Destroy()


    # --- нажатие кнопки "Выбрать" --------------------------------------------
    def OnQuery(self, event):
        self.ExecSQL(None, event)
        self.sb.SetStatusText(u"Операций: - -", 4)
        self.sb.SetStatusText(u"Сумма: - - руб.", 5)



    # -->>>---- М Е Т О Д Ы   К Л А С С А
    #
    # --- основной метод (формирование и осуществление запроса к БД) ----------
    def ExecSQL(self, authorizationCode, event):
        self.sb.SetStatusText(u"Идёт поиск...", 0)

        ### > очистить сетку, если она не пуста
        if self.gridPanel.grid.GetNumberRows() > 0:
            self.gridPanel.grid.DeleteRows(0, self.gridPanel.grid.GetNumberRows())
        ###

        ### > динамическое формирование запроса к БД < ###
        selectCriteria  = []

        columnsRequest  = []
        lunoList        = []
        rcList          = []

        ### > запрашиваемые колонки
        columnsRequest.append('DateTime')
        columnsRequest.append('Luno')
        columnsRequest.append('Operation')
        columnsRequest.append('Amount')
        columnsRequest.append('Host')
        columnsRequest.append('RC')
        columnsRequest.append('RRN')        # RRN (retrieval reference number) – уникальный номер транзакции
        ###

        ### > какой тип луно нужен?
        typeLuno = filterDict[self.oprFilterPanel.comboBox.GetValue()].get('typeLuno')
        ###

        ### > формирование списка Luno
        if self.opsPanel.checkBox.IsChecked():      # для всех ОПС
            for key in bpkDict:
                if typeLuno == None:
                    lunoList.append(bpkDict[key].get('pvn'))
                    lunoList.append(bpkDict[key].get('pos'))
                else:
                    lunoList.append(bpkDict[key].get(typeLuno))
        else:                                       # для выбранных ОПС
            allOPSNameList  = self.opsPanel.allOPSNameList
            opsSelectList   = self.opsPanel.opsSelectList
            opsWPNameList   = self.opsPanel.opsWPNameList
            wpSelectList    = self.opsPanel.wpSelectList

            opsSelNameList = [allOPSNameList[i] for i in opsSelectList]
            wpSelNameList  = [opsWPNameList[i] for i in wpSelectList]

            ### > получить список индексов ОПС из списка имён ОПС
            opsSelIndexList = []
            for nameOPS in opsSelNameList:
                for indexOPS in opsDict:
                    if opsDict[indexOPS].get('nameOPS') == nameOPS:
                        opsSelIndexList.append(indexOPS)
            ###

            ### > по спискам ОПС и РМ сформировать список Luno
            if len(wpSelNameList) == 0:
                for index in opsSelIndexList:
                    for key in bpkDict:
                        if bpkDict[key].get('instPlace') == index:
                            if typeLuno == None:
                                lunoList.append(bpkDict[key].get('pvn'))
                                lunoList.append(bpkDict[key].get('pos'))
                            else:
                                lunoList.append(bpkDict[key].get(typeLuno))
            else:
                index = opsSelIndexList[0]
                for workPlace in wpSelNameList:
                    for key in bpkDict:
                        if bpkDict[key].get('instPlace') == index and\
                            bpkDict[key].get('instPoint') == workPlace:
                            if typeLuno == None:
                                lunoList.append(bpkDict[key].get('pvn'))
                                lunoList.append(bpkDict[key].get('pos'))
                            else:
                                lunoList.append(bpkDict[key].get(typeLuno))
            ###
        ###

        ### > удаление элементов None списка lunoList
        if lunoList.count("None") > 0:
            while lunoList.count("None") != 0:
                lunoList.remove("None")
        ###

        ### > обрамление элеметов списка lunoList в кавычки
        lunoList = ['"' + i + '"' for i in lunoList]
        ###

        ### > добавление в список строки критерия выборки по луно
        selectCriteria.append('Luno IN(%s)' % string.join(lunoList, ', '))
        ###

        ### > формирование шаблонов запрашиваемых дат
        if self.periodPanel.checkBox.GetValue():
            bf = self.periodPanel.dateBefore.GetValue()
            lt = self.periodPanel.dateLater.GetValue()
            bfDate = datetime.date(bf.GetYear(), bf.GetMonth()+1, bf.GetDay())
            ltDate = datetime.date(lt.GetYear(), lt.GetMonth()+1, lt.GetDay())
            dateTmplString = dateTmplStr('DateTime', 'LIKE', 'OR', bfDate, ltDate)
        else:
            prdUserChoice = self.periodPanel.comboBox.GetValue()
            now = datetime.datetime.now()

            if prdUserChoice == periodList[0]:      # сегодня
                dateTmplString = dateTmplStr('DateTime', 'LIKE', 'OR', 1)
            elif prdUserChoice == periodList[1]:    # вчера
                dif = datetime.timedelta(days = 1)
                dateTmplString = 'DateTime LIKE "%s______"' % \
                                    (now - dif).strftime("%Y%m%d")
            elif prdUserChoice == periodList[2]:    # позавчера
                dif = datetime.timedelta(days = 2)
                dateTmplString = 'DateTime LIKE "%s______"' % \
                                    (now - dif).strftime("%Y%m%d")
            elif prdUserChoice == periodList[3]:    # последние 3 дня
                dateTmplString = dateTmplStr('DateTime', 'LIKE', 'OR', 3)
            elif prdUserChoice == periodList[4]:    # последние 10 дней
                dateTmplString = dateTmplStr('DateTime', 'LIKE', 'OR', 10)
            elif prdUserChoice == periodList[5]:    # последние 30 дней
                dateTmplString = dateTmplStr('DateTime', 'LIKE', 'OR', 30)
        ###

        ### > добавление в список строки критерия выборки по дате (диапазону дат)
        selectCriteria.append('(%s)' % dateTmplString)
        ###

        ### > формирование критериев поиска из словаря
        filterChoice = self.oprFilterPanel.comboBox.GetValue()
        for template in filterDict[filterChoice].keys():
            tmplList = []
            if template in ['order', 'typeLuno', 'exclude']:
                pass
            else:
                tmplList.extend(filterDict[filterChoice].get(template))

                ### > обрамление элеметов списка tmplList в кавычки
                tmplList = ['"' + i + '"' for i in tmplList]
                ###

                if len(tmplList) == 0:
                    pass
                elif len(tmplList) == 1:
                    ### > добавление в список строки критерия выборки из словаря
                    selectCriteria.append('%s=%s' % (template, tmplList[0]))
                else:
                    selectCriteria.append('%s IN(%s)' % (template, string.join(tmplList, ', ')))
                    ###
        ###

        ### > добавление в список строки критерия выборки по хосту
        hostChoice = self.hstFilterPanel.comboBox.GetValue()
        if hostChoice != self.hstFilterPanel.hostFilterList[0]:
            for key, value in hostDict.items():
                if value == hostChoice:
                    selectCriteria.append('Host="%s"' % key)
                    break
        ###

        if authorizationCode == None:
            if len(lunoList) == 0: # не выбрано ни одно ОПС
                self.gaugePanel.stText.SetLabel(u"")
                self.sb.SetStatusText(u"Операций: - -", 4)
                self.sb.SetStatusText(u"Сумма: - - руб.", 5)

                ### > показать информационное сообщение
                wx.MessageBox(u'Выберите хотя-бы одно ОПС!' + 10 * ' ',
                              u'Выбор ОПС',
                              wx.OK | wx.ICON_EXCLAMATION,
                              self)
                ###
                opsSelection = False
            else:
                opsSelection = True
        else:
            if authorizationCode != None:
                opsSelection = True

                ### > формирование критерия поиска уникальной операции
                selectCriteria = []
                #selectCriteria.append('RRN RLIKE "%s$"' % authorizationCode)
                selectCriteria.append('RRN="%s"' % authorizationCode)
                ###
            else:
                opsSelection = False
                self.sb.SetStatusText(u"Готов!", 0)

        if opsSelection:
            ### > результирующий sql-запрос к БД
            self.queryDB = 'SELECT %s FROM pnc_trans WHERE %s ORDER BY DateTime DESC;' % \
            (string.join(columnsRequest, ', '), string.join(selectCriteria, ' AND '))
            ###

            ### > выполнение SQL-запроса в отдельном потоке
            ExecSQLThread(self.queryDB)
            self.buttonPanel.buttonQuery = event.GetEventObject()
            ###

            ### > состояние элементов интерфейса
            ###self.gaugePanel.StartTimer()
            #self.StartProgress()###
            self.progressStatusBar.Start()###
            self.buttonPanel.buttonQuery.Disable()
            self.gridPanel.grid.Disable()
            ###



    # -->>>---- П Р И Ё М Н И К И
    #
    # --- метод-приёмник данных из потока -------------------------------------
    def DataReceiver(self, msg):
        # > поля таблицы базы данных
        # rec[0] - дата и время ('DateTime')
        # rec[1] - луно         ('Luno')
        # rec[2] - операция     ('Operation')
        # rec[3] - сумма        ('Amount')
        # rec[4] - хост         ('Host')
        # rec[5] - ответ хоста  ('RC')

        tmCorr = 0 # коррекция времени
        data = msg.data

        if data == ():
            ###self.gaugePanel.StopTimer()
            #self.StopProgress()###
            self.progressStatusBar.Stop()###
            self.sb.SetStatusText(u"", 0)
            self.sb.SetStatusText(u"Операций: 0", 4)
            self.sb.SetStatusText(u"Сумма: 0 руб.", 5)
            wx.MessageBox(u"Ничего не найдено" + 10 * " ",
                          u"Выборка и поиск",
                          wx.OK | wx.ICON_INFORMATION,
                          self)
            self.sb.SetStatusText(u"Готов к следующему поиску!", 0)
            #self.sb.SetStatusText(u"Сумма: - - руб.", 5)
            self.buttonPanel.buttonQuery.SetFocus()
        else:
            rowNumber = 0    # счётчик строк
            summ      = 0    # счётчик общей суммы
            for rec in data:
                dt = datetime.datetime(*time.strptime(rec[0], '%Y%m%d%H%M%S')[:6])

                ### > дата проведения операции
                dateToGrid = '%s.%s.%s' % (dt.day, str(dt.month).zfill(2), dt.year)
                ###

                ### > временя проведения операции
                timeToGrid = '%s:%s:%s' % (dt.hour + tmCorr, str(dt.minute).zfill(2), str(dt.second).zfill(2))
                ###

                ### > ОПС, рабочее место и ЛУНО устройства
                for key in bpkDict:
                    if rec[1] == bpkDict[key].get('pvn') or rec[1] == bpkDict[key].get('pos'):
                        opsToGrid   = opsDict[bpkDict[key].get('instPlace')].get('nameOPS')
                        pointToGrid = bpkDict[key].get('instPoint')
                        posToGrid   = bpkDict[key].get('pos')
                        pvnToGrid   = bpkDict[key].get('pvn')
                        break
                ###

                ### > тип операции
                for key in bpkDict:
                    if rec[1] == bpkDict[key].get('pos'):
                        typeLuno = 'pos'
                        break
                else:
                    typeLuno = 'pvn'

                if rec[2] == '1' and typeLuno == 'pos':
                    operationToGrid = u"Оплата услуг"
                else:
                    operationToGrid = operationDict.get(rec[2], u"Не определена")
                ###

                ### > сумма операции
                amount = Decimal(rec[3])/100
                amountToGrid = toUserView2(str(amount.quantize(Decimal('.00'))), 1)
                summ += amount
                
                '''
                if rec[4] == 'GASPROM':
                    rub, kop = divmod(int(rec[3]), 100)
                    amountToGrid = toUserView(str(rub))
                    summ += int(rub)
                    if kop > 0:
                        amountToGrid += '.'
                        amountToGrid += str(kop).zfill(2)
                else:
                    amountToGrid = toUserView(rec[3])
                    summ += int(rec[3])
                '''
                ###

                ### > хост
                hostToGrid = hostDict.get(rec[4], u"Неизвестен")
                ###

                ### > ответ хоста
                if rec[5] == '':
                    anwToGrid = rcDict.get(rec[5])
                else:
                    anwToGrid = u"%s - %s" % (rec[5], rcDict.get(rec[5], u"Не определён"))
                ###

                ### > ref-код операции (уникальный номер транзакции)
                refToGrid = rec[6]
                ###

                ### > заполнение сетки полученными данными
                notList = ("None", None)
                dataToGrid = []
                dataToGrid.append(dateToGrid)
                dataToGrid.append(timeToGrid)
                dataToGrid.append(opsToGrid)
                dataToGrid.append(pointToGrid)
                dataToGrid.append('' if posToGrid in notList else posToGrid)
                dataToGrid.append('' if pvnToGrid in notList else pvnToGrid)
                dataToGrid.append(operationToGrid)
                dataToGrid.append(amountToGrid + '  ')
                dataToGrid.append(hostToGrid)
                dataToGrid.append(anwToGrid)
                dataToGrid.append(refToGrid)
                # добавить строку
                self.gridPanel.grid.AppendRows()
                # заполнить строку данными
                for cellNumber, cellData in enumerate(dataToGrid):
                    self.gridPanel.grid.SetCellValue(rowNumber, cellNumber, cellData)
                ###

                ### > установить выравнивание и шрифт в столбце "Сумма"
                self.gridPanel.grid.SetCellAlignment(rowNumber, 7, wx.ALIGN_RIGHT,
                                                                  wx.ALIGN_CENTRE)
                self.gridPanel.grid.SetCellFont(rowNumber, 7, wx.Font(8, wx.SWISS,
                                                                        wx.ITALIC,
                                                                        wx.BOLD))
                ###
                rowNumber += 1

            ###self.gaugePanel.StopTimer()
            #self.StopProgress()###
            self.progressStatusBar.Stop()###
            self.sb.SetStatusText(u"Готов к следующему поиску!", 0)
            self.sb.SetStatusText(u"Операций: %s" % str(rowNumber), 4)
            self.sb.SetStatusText(u"Сумма: %s руб." % str(summ), 5)
            self.gridPanel.grid.SetFocus()
        self.gridPanel.grid.Enable()
        self.buttonPanel.buttonQuery.Enable()


    # --- метод-приёмник сообщений об ошибках из потока -----------------------
    def ErrorReceiver(self, msg):
        self.timer.Stop()
        self.gaugePanel.gauge.ClearBackground()
        self.sb.SetStatusText(u"", 0)
        self.sb.SetStatusText(u"Операций: - -", 4)
        self.sb.SetStatusText(u"Сумма: - - руб.", 5)
        wx.MessageBox(msg.data + 10 * " ",
                      u"Выборка и поиск",
                      wx.OK | wx.ICON_ERROR,
                      self)
        self.gridPanel.grid.Enable()
        self.buttonPanel.buttonQuery.Enable()
        self.buttonPanel.buttonQuery.SetFocus()
        self.sb.SetStatusText(u"Неудача. Попробуйте ещё раз!", 0)
#
# Конец класса основного окна программы AppMainFrame
###############################################################################
