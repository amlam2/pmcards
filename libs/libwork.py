# coding=utf-8

import os, sys
import zipfile
import datetime, calendar
from liblore import codeSDODict, monthDict

### -- > словарь типов архивов < ----------------------------------------------
sfTypeDict = {
                'er_imp': u'',
                'db_ktg': u'Подписной каталог пакета "ПОЧТА"',
                'db_kta': u'Изменение сроков подписки',
                'db_lim': u'Нормы вадачи в доставку',
                'db_blt': u'База "Балтпост"',
                'db_tar': u'Тарифы',
                'fl_ver': u'Обновление версии',
                'db_pns': u'База пенсий',
                'db_els': u'База платежей ЕЛС',
                'db_sep': u'База платежей',
                'fl_inf': u'Информационный файл',
                'db_pdp': u'Подписной каталог программы "Подписка"',
                'ac_upd': u'Архив обновления файлов на диске D',
                'fl_ukn': u'Не определено',
                'db_lot': u'Лотерея'
             }
# -----------------------------------------------------------------------------


### -- > словарь видов лотерейных билетов < -----------------------------------
sfLotDict = {
                'm': u'Мiласэрнасць',
                'p': u'Придвинье',
                'z': u'Скарбнiца',
                'g': u'Гомельчанка',
                'h': u'Столица',
                'i': u'Наша спадчына',
                'j': u'Молодёжная',
                'k': u'За безопасность',
                'l': u'Большие деньги',
                'o': u'Шчодрая крама',
                'q': u'Бобруйская крепость',
                'r': u'Приднепровье-1',
                't': u'Праздничная',
                'u': u'Победа',
                'w': u'Принёманская',
                'x': u'Я люблю Могилёв',
                'y': u'Брестская крепость-герой',
                's': u'Суперлото',
                'v': u'Ваше лото',
                '5': u'Пятёрочка',
                'd': u'Суперлото (блок)',
                'e': u'Ваше лото (блок)',
                'f': u'Пятёрочка (блок)'
            }
# -----------------------------------------------------------------------------


### -- > словарь дополнительных параметров архивов < --------------------------
infoTypeDict = {
                'half-year'     : u'полугодие',
                'month'         : u'месяц',
                'year'          : u'год',
                'load'          : u'загрузка',
                'ver'           : u'версия',
                'release'       : u'релиз',
                'correction'    : u'коррекция',
                'period'        : u'период',
                'ops'           : u'ОПС',
                'tirage'        : u'тираж',
                'sdo'           : u'код СДО',
                'damaged'       : u'повреждено'
                }
# -----------------------------------------------------------------------------


# -----------------------------------------------------------------------------
# -- функция преобразует числовую строку вида "145563800" в строку
# -- "145 563 800", т. е. разбивает число по разрядам
def toUserView(inString): # >> str
    if int(inString) >= 1000:
        resultList = []
        x = int(inString)
        while True:
            x, y = divmod(x, 1000)
            resultList.append(str(y).zfill(3))
            if x < 1000:
                resultList.append(str(x))
                break
            elif x >= 1000:
                continue
        return ' '.join(resultList[::-1])
    else:
        return str(int(inString))
# -----------------------------------------------------------------------------


# -----------------------------------------------------------------------------
# -- 
# -- 
def toUserView2(inString, flag=0): # >> str
    rub, kop = inString.split('.')
    part_1, part_2 = int(rub), int(kop)
    if int(rub) >= 1000:
        resultList = []
        x = int(rub)
        while True:
            x, y = divmod(x, 1000)
            resultList.append(str(y).zfill(3))
            if x < 1000:
                resultList.append(str(x))
                break
            elif x >= 1000:
                continue
        rub = ' '.join(resultList[::-1])
    if flag:
        if not part_1:
            return u"%s к." % (kop)
        #elif not part_2:
        #    return u"%s р." % (rub)
        else:
            return u"%s р. %s к." % (rub, kop)
    else:
        return '.'.join([rub, kop])
# -----------------------------------------------------------------------------


# -----------------------------------------------------------------------------
# -- функция формирует строку шаблонов дат для динамического формирования
# -- SQL-запроса к базе данных.
# -- Пример использования:
# -| d1 = datetime.datetime(2011, 10, 26)
# -| d2 = datetime.datetime(2011, 10, 28)
# -| print dateTmplStr('DateTime', 'LIKE', 'OR', d1, d2) # от даты d1 до даты d2
# -| >>> DateTime LIKE "20111028______" OR
# -|     DateTime LIKE "20111027______" OR
# -|     DateTime LIKE "20111026______"
# -| print dateTmplStr('DateTime', 'LIKE', 'OR', 3) # от текущей даты назад 3 дня
# -| >>> DateTime LIKE "20111028______" OR
# -|     DateTime LIKE "20111027______" OR
# -|     DateTime LIKE "20111026______"
# -| print dateTmplStr('DateTime', 'LIKE', 'OR', 3, d2) # от даты d2 назад 3 дня
# -| >>> DateTime LIKE "20111028______" OR
# -|     DateTime LIKE "20111027______" OR
# -|     DateTime LIKE "20111026______"

tempDict    = {}
temp        = {}

def dateTmplStr(field, operator, combiner, before, later=datetime.datetime.now()):
    ### > инициализация временных словарей
    tempDict    = {}
    temp        = {}
    ###

    ### > формирование списка шаблонов
    if isinstance(before, int):
        delta = before - 1
    else:
        delta = (later - before).days
    result = []
    result.append(later.strftime("%Y%m%d"))
    for item in iter(range(delta)):
        dif = datetime.timedelta(days = item + 1)
        tmpl = later - dif
        result.append(tmpl.strftime("%Y%m%d"))
    ###

    ### > заполнение временного словаря
    for tmpl in iter(result):
        if tempDict.get(tmpl[:6]) == None:
            tempDict[tmpl[:6]] = {int(tmpl[6]): []}
        elif tempDict.get(tmpl[:6]).get(int(tmpl[6])) == None:
            temp = tempDict[tmpl[:6]]
            temp[int(tmpl[6])] = []
            tempDict[tmpl[:6]] = temp
        tempDict.get(tmpl[:6]).get(int(tmpl[6])).append(int(tmpl[-2:]))
    ###

    ### > сокращение списка шаблонов
    for tmpl in iter(tempDict):
        monthDays = calendar.monthrange(int(tmpl[:4]), int(tmpl[-2:]))[1]
        dayList = []
        for item in iter(tempDict[tmpl].keys()):
            dayList.extend(tempDict[tmpl].get(item))
        if len(dayList) == monthDays:
            for item2 in iter(dayList):
                result.remove(tmpl + str(item2).zfill(2))
            result.append(tmpl)
        else:
            bgn = [1, 10, 20, 30]
            end = [10, 20, 30, monthDays + 1]
            if monthDays/10 == 2:
                bgn.remove(30)
                end.remove(30)
            for dec in iter(tempDict[tmpl].keys()):
                if sorted(tempDict[tmpl].get(dec, [])) == range(bgn[dec], end[dec]):
                    for item2 in iter(tempDict[tmpl].get(dec)):
                        result.remove(tmpl + str(item2).zfill(2))
                    result.append(tmpl + str(dec))
    ###

    result = sorted([x + '_'*(14-len(x)) for x in result])
    result = [field + ' ' + operator + ' "' + x + '"' for x in result]

    return (' ' + combiner + ' ').join(result)
# -----------------------------------------------------------------------------


# -----------------------------------------------------------------------------
# -- функция проверяет последовательность list на наличие элемента tmpl без
# -- учёта регистра символов и возвращает этот элемент. Если в
# -- последовательности несколько одинаковых элементов - возвращает первое
# -- вхождение.
# -- В случае отсутствия элемента в последовательности возвращает None.
def cmp_IGNORECASE(list, tmpl):
    for item in list:
        if item.lower() == tmpl.lower():
            return item
            break
    else:
        return None
# -----------------------------------------------------------------------------


# -----------------------------------------------------------------------------
#
def findFile(dir, template):
    from glob import glob1
    from random import randint

    fileList = glob1(dir, template)
    if len(fileList) > 0:
        return fileList[randint(0, len(fileList) - 1)]
    else:
        return None
# -----------------------------------------------------------------------------


# -----------------------------------------------------------------------------
# -- функция рассчитывает маску сети из её длины и результат
# -- возвращает в виде строки
# -- length - длина маски (int 1...32)
def calcLen2Mask(length): # >> str
    if 0 < length <= 32:
        string = ('1'*length).ljust(32, '0')
        result = []
        for i in xrange(0, 32, 8):
            result.append(str(int(string[i:i+8], 2)))
        return '.'.join(result)
    else:
        return ''
# -----------------------------------------------------------------------------


### -- > этот класс содержит методы для работы с отправляемыми файлами < ------
class SFInfo(object):
    def __nt(file):
        pass
    def __np(file):
        pass
    def __nl(file):
        pass
    def __ns(file):
        pass
    def __vn(file):
        pass
    def __nb(file):
        pass
    def __limit(file):
        pass
    def __katalog_zip(file):
        pass
    def __pdpskops_rar(file):
        pass
    def __info_txt(file):
        pass
    def __fupdate_rar(file):
        pass
    # -------------------------------------------------------------------------
    # -- анализ отправляемого на ОПС файла
    # -- возвращат его содержание в виде списка кодов
    @staticmethod
    def sfInfo(sfPath): # >> list
        sfCode = []
        try:
            from dbfpy import dbf
        except ImportError:
            sfCode.append(u"Не установлен модуль dbfpy")
            sfCode.insert(0, 'er_imp')

        sfName = os.path.basename(sfPath).lower().strip()
        ### > временный каталог
        if sys.platform == 'win32':
            tmpDir = os.environ['TEMP']
        else:
            tmpDir = os.getcwd()

        try:
            # здесь обрабатываются zip-архивы
            zf = zipfile.ZipFile(sfPath, 'r')
            if sfName.lower() == 'katalog.zip':
                ### > для подписного каталога пакета ПОЧТА (db_ktg)
                agentFile   = None
                katalogFile = None
                ouslFile    = None

                for item in zf.namelist():
                    if item.find('k') == 0 or item.find('K') == 0:
                        katalogFile = item
                    elif item.find('a') == 0 or item.find('A') == 0:
                        agentFile = item
                    elif item.find('o') == 0 or item.find('O') == 0:
                        ouslFile = item

                halfYear = os.path.splitext(zf.namelist()[0])[0][-1:]
                sfCode.append('half-year=' + halfYear)
                appendList = []
                for item in zf.namelist():
                    appendList.append(item.lower())
                sfCode.append(sorted(appendList))

                if agentFile != None and katalogFile == None and ouslFile == None:
                    sfCode.insert(0, 'db_kta')
                else:
                    sfCode.insert(0, 'db_ktg')
            elif os.path.splitext(sfName)[0] == 'limit':
                ### > для норм выдачи в доставку (db_lim)
                for item in zf.namelist():
                    if item.find('l') == 0 or item.find('L') == 0:
                        limitFile = item
                        break

                zf.extract(limitFile, tmpDir)
                dbfFile = dbf.Dbf(os.path.join(tmpDir, limitFile))
                sfCode.append('month=' + str(dbfFile[0]['MES']).zfill(2))
                sfCode.append('year=' + str(dbfFile[0]['GOD']))
                dbfFile.close()
                os.remove(os.path.join(tmpDir, limitFile))
                appendList = []
                for item in zf.namelist():
                    appendList.append(item.lower())
                sfCode.append(sorted(appendList))
                sfCode.insert(0, 'db_lim')
            elif sfName[:2] == 'nb':
                ### > для базы BaltPost (db_blt)
                sfCode.insert(0, 'db_blt')
            elif sfName[:2] == 'nt':
                ### > для базы тарифов (db_tar)
                appendList = []
                for item in zf.namelist():
                    appendList.append(item.lower())
                sfCode.append(sorted(appendList))
                sfCode.insert(0, 'db_tar')
            elif sfName[:2] == 'vn':
                ### > для файла обновления версии (fl_ver)
                ver = str(int(os.path.splitext(sfName)[0][3]) - 1) + '.' +\
                        os.path.splitext(sfName)[0][4:6]
                release = str(int(os.path.splitext(sfName)[0][-2:])) + '(' +\
                            os.path.splitext(sfName)[1][-2] + ')'
                correction = os.path.splitext(sfName)[1][-1]
                sfCode.append('ver=' + ver)
                sfCode.append('release=' + release)
                #sfCode.append('correction=' + correction)
                sfCode.insert(0, 'fl_ver')
            elif sfName[:2] == 'ns':
                ### > для базы пенсий (db_pns)
                sfCode.append('period=' + os.path.splitext(sfName)[1][-1:])
                sfCode.append('month=' + os.path.splitext(sfName)[0][-4:-2].zfill(2))
                sfCode.append('year=' + '20' + os.path.splitext(sfName)[0][-2:])
                sfCode.append('ops=' + os.path.splitext(zf.namelist()[0])[0][-6:])
                sfCode.insert(0, 'db_pns')
            elif sfName[:2] == 'nl':
                ### > для базы тиражей лотерей (db_lot)
                for fileLot in zf.namelist():
                    if fileLot[0].lower() == 'l':
                        sfCode.append('tirage=' + os.path.splitext(fileLot)[0][1:].lower())
                if sfCode:
                    sfCode.insert(0, 'db_lot')
                else:
                    sfCode.insert(0, 'fl_ukn')
            elif sfName[:2] == 'np':
                for item in zf.namelist():
                    dFile = None
                    if item.find('D') == 0 or item.find('d') == 0:
                        dFile = item
                        break
                if dFile != None:
                    if os.path.splitext(dFile)[0][-3:] == '100':
                        ### > для базы ЕЛС (db_els)
                        zf.extract(dFile, tmpDir)
                        dbfFile = dbf.Dbf(os.path.join(tmpDir, dFile))
                        for rec in dbfFile:
                            sfCode.append('sdo=' + str(rec['KOD_DN']))
                        dbfFile.close()
                        os.remove(os.path.join(tmpDir, dFile))
                        sfCode.insert(0, 'db_els')
                    else:
                        ### > для отдельной базы (db_sep)
                        sfCode.append('sdo=' + os.path.splitext(dFile)[0][-3:])
                        sfCode.insert(0, 'db_sep')
                elif dFile == None:
                    if os.path.splitext(zf.namelist()[0])[0][-3:] == '100':
                        ### > для базы ЕЛС (db_els)
                        sfCode.append('damaged')
                        sfCode.insert(0, 'db_els')
                    else:
                        ### > для отдельной базы (db_sep)
                        sfCode.append('sdo=' + os.path.splitext(zf.namelist()[0])[0][-3:])
                        sfCode.insert(0, 'db_sep')
            else:
                #if os.path.splitext(sfName)[1] in ['.zip', '.rar', '.7z']:
                ### > для неизвестного файла (fl_ukn)
                sfCode.insert(0, 'fl_ukn')
        except:
            # здесь обрабатываются остальные файлы
            if sfName == 'info.txt':
                ### > для информационного файла (fl_inf)
                sfCode.insert(0, 'fl_inf')
            elif sfName == 'pdpskops.rar':
                ### > для подписного каталога программы "Подписка" Парадокс (db_pdp)
                sfCode.insert(0, 'db_pdp')
            elif sfName == 'fupdate.rar':
                ### > для архива обновления файлов на диске D (ac_upd)
                sfCode.insert(0, 'ac_upd')
            else:
                ### > для неизвестного файла (fl_ukn)
                sfCode.insert(0, 'fl_ukn')
        else:
            ### > для неизвестного файла (fl_ukn)
            #sfCode.insert(0, 'fl_ukn')
            pass
        return sfCode
    # -------------------------------------------------------------------------


    # -------------------------------------------------------------------------
    # -- преобразование полученного в предыдущем методе
    # -- списка кодов в короткую строку информации о файле
    @staticmethod
    def sfInfoShortString(codeList): # >> str
        resultStr = ''
        sfType = codeList.pop(0)
        resultStr += sfTypeDict.get(sfType)

        if sfType == 'db_els' and codeList[0] != 'damaged':
            resultStr += u': СДО'
            for item in iter(codeList):
                resultStr += u' '
                resultStr += item.split('=')[1]
                resultStr += u','
        elif sfType == 'db_sep':
            resultStr += u': '
            resultStr += codeList[0].split('=')[1]
            resultStr += u' - '
            resultStr += codeSDODict.get(codeList[0].split('=')[1]) or u"Не определено"
        elif sfType == 'db_lot':
            if len(codeList) == 1:
                lot_tir = codeList[0].split('=')[1]
                resultStr += u' "'
                resultStr += sfLotDict.get(lot_tir[0])
                resultStr += u'": тираж '
                resultStr += str(int(lot_tir[1:]))
            elif 1 < len(codeList) <= 3:
                resultStr = resultStr[:-1] + u'и'
                resultStr += u':'
                for item in iter(codeList):
                    resultStr += u' "'
                    resultStr += sfLotDict.get(item.split('=')[1][0])
                    resultStr += u'",'
            elif len(codeList) > 3:
                resultStr = u'Сборник %s' % (resultStr[:-1].lower() + u'й',)
        else:
            if len(codeList) > 0:
                # если элемент списка является списком, то его удалить
                for code in iter(codeList[:]):
                    if isinstance(code, list):
                        codeList.remove(code)
                resultStr += u':'
                for code in iter(codeList):
                    if code == 'damaged':
                        resultStr += u' '
                        resultStr += infoTypeDict.get(code)
                    else:
                        resultStr += u' '
                        resultStr += infoTypeDict.get(code.split('=')[0])
                        resultStr += u' - '
                        if code.split('=')[0] == 'month':
                            resultStr += monthDict.get(code.split('=')[1])
                        else:
                            resultStr += code.split('=')[1]
                        resultStr += u','
        return resultStr.rstrip(',').rstrip(':')
    # -------------------------------------------------------------------------


    # -------------------------------------------------------------------------
    # -- получение описания файла из самого этого файла если он представляет
    # -- собой zip-ахив. Ищет в архиве файл описания (по-умолчанию info.txt),
    # -- читает его и возвращает это описание в виде строки. Если описания
    # -- нет - возвращает по-умолчанию пустую строку (пожно переопределить).
    # -- Если файл не является zip-ахивом - возвращает пустую строку
    @staticmethod
    def sfGetDescription(fileName,\
                         default=u'',\
                         dsTemplate = 'info.txt',
                         coding     = 'cp1251'): # >> str
        if sys.platform == 'win32':
            tmpDir = os.environ['TEMP']
        else:
            tmpDir = os.getcwd()

        try:
            with zipfile.ZipFile(fileName, 'r') as zf:
                dsFile = cmp_IGNORECASE(zf.namelist(), dsTemplate)
                if dsFile:
                    zf.extract(dsFile, tmpDir)
                    with open(os.path.join(tmpDir, dsFile), 'rb') as f:
                        descript = u''
                        for line in f:
                            line = unicode(line, coding)
                            descript += line
                    return descript
                else:
                    return default
        except:
            return u''
    # -------------------------------------------------------------------------


    # -------------------------------------------------------------------------
    # -- сколько суток файлу
    @staticmethod
    def sfHowOldFile(file): # >> int
        dtToday = datetime.datetime.now()
        dtStamp = os.path.getmtime(file)
        dtMtime = datetime.datetime.fromtimestamp(dtStamp)
        return (dtToday - dtMtime).days
    # -------------------------------------------------------------------------


    # -------------------------------------------------------------------------
    # -- время последнего изменения файла
    @staticmethod
    def sfMTimeFile(file): # >> str
        dtStamp = os.path.getmtime(file)
        dtMtime = datetime.datetime.fromtimestamp(dtStamp)
        return dtMtime.strftime("%d.%m.%Y | %H:%M")
    # -------------------------------------------------------------------------


    # -------------------------------------------------------------------------
    # -- delta=0 - дата создания файла
    # -- delta=1 - следующая дата после даты создания etc.
    @staticmethod
    def sfReceivDate(file, delta=0): # >> str
        dtStamp = os.path.getmtime(file)
        dtMtime = datetime.datetime.fromtimestamp(dtStamp)
        if delta:
            return (dtMtime + datetime.timedelta(days=delta)).strftime("%d.%m.%Y")
        else:
            return dtMtime.strftime("%d.%m.%Y")
    # -------------------------------------------------------------------------
