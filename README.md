#Мониторинг операций по Банковским Пластиковым Карточкам

**Название приложения:** pmcards -- PayMent Cards.

**Назначение приложения:** мониторинг операций, проведённых по Банковским Пластиковым Карточкам (БПК) в отделениях почтовой связи Берёзовского РУПС.

В приложении имеются фильры по конкретному отделению почтовой связи, по дате проведения операции, по банку-эквайеру, по статусу операции. После определения пользователем критериев поиска приложение осуществляет доступ к БД сервера БПК филиала, находит нужный платёж (платежи) и отображает в GUI необходимую информацию по нему (им) в удобном для пользователя виде. Используется инженерами группы поддержки для контроля факта наличия операции, даты и времени проведения операции, суммы платежа а также его статуса (код ответа банка-эквайера). Позволяет ускорить обработку сбойных и спорных ситуаций.


Приложение написано на **Python 2.7** с использованием библиотеки кроссплатформенного графического интерфейса пользователя (GUI) **WxPython**.


Используются модули, не входящие в стандартную поставку интерпретатора:
        - wxPython;
        - MySQLdb;
        - xlwt
