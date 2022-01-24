#!/usr/bin/env python
# coding=utf-8

# подключение необходимых модулей
import os
import wx

# Глобальные переменные
###############################################################################
# каталоги программы
###############################################################################
startDir  = os.getcwd()
tempDir   = os.path.join(startDir, 'temp')
picsDir   = os.path.join(startDir, 'pics')
reportDir = os.path.join(startDir, 'report')
iconDir   = os.path.join(picsDir,  'icons')

if not os.path.exists(reportDir):
    os.makedirs(reportDir)

windowTitle     = u'Операции по БПК - Берёзовский РУПС' # заголовок окна
applicationIcon = os.path.join(iconDir, 'pmcards.png')  # иконка приложения

### ->-- параметры соединения с БД операций с использованием БПК
dbHost, dbUser, dbPasswd, dbName = "192.168.1.14", "viewer", "xw7jswrjgQ", "bpk"
###

# подключение диалогов
import pmcframe


###############################################################################
# Начало класса приложения (CardApp)
#
class CardApp(wx.App):
    def OnInit(self):
        wx.InitAllImageHandlers()
        '''
        ### > показать заставку
        image = wx.Image(os.path.join(picsDir, 'payment_card.png'), wx.BITMAP_TYPE_PNG)
        wx.SplashScreen(image.ConvertToBitmap(),
                        wx.SPLASH_CENTRE_ON_SCREEN | wx.SPLASH_TIMEOUT,
                        500,
                        None,
                        -1)
        ###
        
        wx.Yield()
        '''
        frame = pmcframe.AppMainFrame(None, -1, "")
        self.SetTopWindow(frame)
        frame.Show()
        return True
#
# Конец класса CardApp
###############################################################################

def main():
    app = CardApp(0)
    app.MainLoop()

if __name__ == "__main__":
    main()
