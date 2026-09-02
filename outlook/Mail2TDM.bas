Attribute VB_Name = "Mail2TDM"
' ============================================================================
'  Mail2TDM - пересылка задач из Outlook в канал мессенджера TDM
'  СКЕЛЕТ (первая версия). Требует токен бота от техподдержки TDM.
' ============================================================================
'
'  КАК УСТАНОВИТЬ:
'    1. В Outlook нажми Alt+F11 (откроется редактор макросов).
'    2. Меню Insert -> Module. Появится пустой модуль.
'    3. СКОПИРУЙ текст этого файла и ВСТАВЬ в модуль.
'       ВАЖНО: именно копировать-вставить, а НЕ импортировать файл -
'       при импорте могут испортиться русские буквы.
'    4. Переименуй модуль в Mail2TDM (панель Properties, поле (Name)).
'    5. Отдельно вставь код из файла ThisOutlookSession.vba.txt
'       в объект ThisOutlookSession (он в списке слева).
'    6. Создай файл настроек %APPDATA%\mail2tdm\config.ini
'       по образцу config.example.ini и заполни его.
'
'  ПЕРВЫЙ ЗАПУСК:
'    В настройках оставь TestMode=1 - тогда НИЧЕГО не отправляется,
'    а всё, что было бы отправлено, пишется в журнал. Убедись, что
'    письма отбираются и оформляются правильно, и только потом ставь 0.
'
'  ПРОВЕРКА СВЯЗИ:
'    Запусти макрос TestSend - он отправит в канал одно тестовое сообщение
'    (работает только при TestMode=0).
' ============================================================================

Option Explicit

' --- Настройки (заполняются из config.ini) ---
Private cfgBaseUrl As String
Private cfgToken As String
Private cfgWorkspaceId As String
Private cfgGroupId As String
Private cfgMailboxName As String
Private cfgFolderPath As String
Private cfgSubjectContains As String
Private cfgSenderContains As String
Private cfgTestMode As Boolean
Private cfgMaxLength As Long
Private cfgLogPath As String

Private cfgLoaded As Boolean
Private lastSendTime As Date       ' для соблюдения лимита 60 запросов в минуту


' ===========================================================================
'  ГЛАВНОЕ: обработка одного входящего письма
'  Вызывается автоматически из ThisOutlookSession при получении письма.
' ===========================================================================
Public Sub ProcessMail(ByVal mail As Object)
    On Error GoTo Fail

    If Not LoadConfig() Then Exit Sub
    If mail Is Nothing Then Exit Sub

    ' --- Отбор: то ли это письмо, которое нам нужно ---
    If Not MailMatches(mail) Then
        WriteLog "ПРОПУЩЕНО (не подходит по фильтру): " & SafeSubject(mail)
        Exit Sub
    End If

    ' --- Оформляем текст задачи ---
    Dim text As String
    text = BuildMessage(mail)

    If Len(Trim$(text)) = 0 Then
        WriteLog "ПРОПУЩЕНО (пустой текст): " & SafeSubject(mail)
        Exit Sub
    End If

    ' --- Отправляем (или только записываем в журнал) ---
    If cfgTestMode Then
        WriteLog "ТЕСТОВЫЙ РЕЖИМ - НЕ отправлено. Было бы отправлено:" & vbCrLf & text
    Else
        If SendToTdm(text) Then
            WriteLog "ОТПРАВЛЕНО: " & SafeSubject(mail)
        Else
            WriteLog "ОШИБКА ОТПРАВКИ: " & SafeSubject(mail)
        End If
    End If
    Exit Sub

Fail:
    WriteLog "СБОЙ в ProcessMail: " & Err.Number & " " & Err.Description
End Sub


' ===========================================================================
'  Проверка: подходит ли письмо под фильтры из настроек
' ===========================================================================
Private Function MailMatches(ByVal mail As Object) As Boolean
    Dim okSubject As Boolean, okSender As Boolean
    okSubject = True
    okSender = True

    If Len(cfgSubjectContains) > 0 Then
        okSubject = (InStr(1, SafeSubject(mail), cfgSubjectContains, vbTextCompare) > 0)
    End If

    If Len(cfgSenderContains) > 0 Then
        okSender = (InStr(1, SafeSender(mail), cfgSenderContains, vbTextCompare) > 0)
    End If

    MailMatches = (okSubject And okSender)
End Function


' ===========================================================================
'  Оформление текста задачи для мессенджера
'  ЗДЕСЬ настраивается внешний вид сообщения - правь под свои письма.
' ===========================================================================
Private Function BuildMessage(ByVal mail As Object) As String
    Dim s As String
    Dim body As String

    body = SafeBody(mail)
    body = CleanupBody(body)

    s = "Новая задача из почты" & vbCrLf & _
        "Тема: " & SafeSubject(mail) & vbCrLf & _
        "От: " & SafeSender(mail) & vbCrLf & _
        "Получено: " & Format$(SafeReceived(mail), "dd.mm.yyyy HH:nn") & vbCrLf & _
        vbCrLf & body

    ' Обрезаем слишком длинное сообщение
    If cfgMaxLength > 0 And Len(s) > cfgMaxLength Then
        s = Left$(s, cfgMaxLength - 20) & vbCrLf & "... (текст обрезан)"
    End If

    BuildMessage = s
End Function


' Убираем лишние пустые строки и подписи-разделители
Private Function CleanupBody(ByVal body As String) As String
    Dim t As String
    t = Replace(body, vbCr & vbLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    ' схлопываем три и более переводов строки в два
    Do While InStr(t, vbLf & vbLf & vbLf) > 0
        t = Replace(t, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop
    CleanupBody = Trim$(t)
End Function


' ===========================================================================
'  ОТПРАВКА в TDM Bot API
'  POST {BaseUrl}/botapi/v1/messages/sendTextMessage/{workspaceId}/{groupId}
'  Заголовок Authorization: <токен>,  тело: {"message":"..."}
' ===========================================================================
Private Function SendToTdm(ByVal text As String) As Boolean
    On Error GoTo Fail
    SendToTdm = False

    If Len(cfgBaseUrl) = 0 Or Len(cfgToken) = 0 _
       Or Len(cfgWorkspaceId) = 0 Or Len(cfgGroupId) = 0 Then
        WriteLog "НЕ НАСТРОЕНО: заполни BaseUrl, Token, WorkspaceId, GroupId в config.ini"
        Exit Function
    End If

    ' Лимит сервера - 60 запросов в минуту. Держим паузу не меньше секунды.
    Do While DateDiff("s", lastSendTime, Now) < 1
        DoEvents
    Loop
    lastSendTime = Now

    Dim url As String
    url = TrimSlash(cfgBaseUrl) & "/botapi/v1/messages/sendTextMessage/" & _
          cfgWorkspaceId & "/" & cfgGroupId

    Dim json As String
    json = "{""message"":""" & JsonEscape(text) & """}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", cfgToken
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    ' Отправляем байтами в UTF-8, иначе русский текст придёт испорченным
    http.Send Utf8Bytes(json)

    If http.Status >= 200 And http.Status < 300 Then
        WriteLog "Ответ сервера: " & http.Status & " " & Left$(http.ResponseText, 200)
        SendToTdm = True
    Else
        WriteLog "Сервер вернул ошибку " & http.Status & ": " & Left$(http.ResponseText, 500)
    End If
    Exit Function

Fail:
    WriteLog "СБОЙ при отправке: " & Err.Number & " " & Err.Description
End Function


' Ручная проверка связи: отправляет одно тестовое сообщение.
Public Sub TestSend()
    If Not LoadConfig() Then Exit Sub
    If cfgTestMode Then
        MsgBox "Сейчас включён ТЕСТОВЫЙ РЕЖИМ (TestMode=1) - реальная отправка отключена." & vbCrLf & _
               "Чтобы проверить связь по-настоящему, поставь TestMode=0 в config.ini.", vbInformation
        Exit Sub
    End If
    If SendToTdm("Проверка связи: mail2tdm подключён.") Then
        MsgBox "Сообщение отправлено. Проверь канал в TDM.", vbInformation
    Else
        MsgBox "Отправить не удалось. Открой журнал: " & cfgLogPath, vbExclamation
    End If
End Sub


' ===========================================================================
'  ЧТЕНИЕ НАСТРОЕК из %APPDATA%\mail2tdm\config.ini
' ===========================================================================
Private Function LoadConfig() As Boolean
    On Error GoTo Fail
    LoadConfig = False

    Dim path As String
    path = Environ$("APPDATA") & "\mail2tdm\config.ini"

    If Len(Dir$(path)) = 0 Then
        MsgBox "Не найден файл настроек:" & vbCrLf & path & vbCrLf & vbCrLf & _
               "Создай его по образцу config.example.ini.", vbExclamation
        Exit Function
    End If

    Dim content As String
    content = ReadUtf8File(path)

    cfgBaseUrl = IniValue(content, "BaseUrl")
    cfgToken = IniValue(content, "Token")
    cfgWorkspaceId = IniValue(content, "WorkspaceId")
    cfgGroupId = IniValue(content, "GroupId")
    cfgMailboxName = IniValue(content, "MailboxName")
    cfgFolderPath = IniValue(content, "FolderPath")
    cfgSubjectContains = IniValue(content, "SubjectContains")
    cfgSenderContains = IniValue(content, "SenderContains")
    cfgTestMode = (IniValue(content, "TestMode") <> "0")     ' по умолчанию БЕЗОПАСНО
    cfgMaxLength = Val(IniValue(content, "MaxLength"))
    If cfgMaxLength <= 0 Then cfgMaxLength = 3000

    cfgLogPath = IniValue(content, "LogPath")
    If Len(cfgLogPath) = 0 Then
        cfgLogPath = Environ$("APPDATA") & "\mail2tdm\mail2tdm.log"
    End If

    cfgLoaded = True
    LoadConfig = True
    Exit Function

Fail:
    MsgBox "Не удалось прочитать настройки: " & Err.Description, vbCritical
End Function


' Достаём значение по имени параметра (строки вида Имя=Значение, ; - комментарий)
Private Function IniValue(ByVal content As String, ByVal key As String) As String
    Dim lines() As String, i As Long, ln As String, p As Long
    content = Replace(content, vbCr & vbLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    lines = Split(content, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) <> ";" And Left$(ln, 1) <> "#" Then
                p = InStr(ln, "=")
                If p > 1 Then
                    If StrComp(Trim$(Left$(ln, p - 1)), key, vbTextCompare) = 0 Then
                        IniValue = Trim$(Mid$(ln, p + 1))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    IniValue = ""
End Function


' ===========================================================================
'  Вспомогательные функции
' ===========================================================================

' Экранирование текста для JSON
Private Function JsonEscape(ByVal s As String) As String
    Dim r As String, i As Long, ch As String, code As Long
    r = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case ch
            Case """": r = r & "\"""
            Case "\":  r = r & "\\"
            Case vbCr: r = r & "\r"
            Case vbLf: r = r & "\n"
            Case vbTab: r = r & "\t"
            Case Else
                If code < 32 Then
                    r = r & "\u" & Right$("000" & Hex$(code), 4)
                Else
                    r = r & ch
                End If
        End Select
    Next i
    JsonEscape = r
End Function

' Преобразуем текст в байты UTF-8 (иначе русский приедет испорченным)
Private Function Utf8Bytes(ByVal s As String) As Variant
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' текст
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1                 ' переключаем на байты
    st.Position = 3             ' пропускаем метку BOM
    Utf8Bytes = st.Read
    st.Close
End Function

' Чтение файла в кодировке UTF-8
Private Function ReadUtf8File(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText
    st.Close
End Function

Private Function TrimSlash(ByVal s As String) As String
    Do While Len(s) > 0 And Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
    Loop
    TrimSlash = s
End Function

' Безопасное чтение полей письма (не падаем, если поля нет)
Private Function SafeSubject(ByVal mail As Object) As String
    On Error Resume Next
    SafeSubject = CStr(mail.Subject)
End Function

Private Function SafeSender(ByVal mail As Object) As String
    On Error Resume Next
    SafeSender = CStr(mail.SenderEmailAddress)
    If Len(SafeSender) = 0 Then SafeSender = CStr(mail.SenderName)
End Function

Private Function SafeBody(ByVal mail As Object) As String
    On Error Resume Next
    SafeBody = CStr(mail.Body)
End Function

Private Function SafeReceived(ByVal mail As Object) As Date
    On Error Resume Next
    SafeReceived = mail.ReceivedTime
    If SafeReceived = 0 Then SafeReceived = Now
End Function


' Запись в журнал - чтобы всегда можно было понять, что произошло
Public Sub WriteLog(ByVal msg As String)
    On Error Resume Next
    Dim path As String
    path = cfgLogPath
    If Len(path) = 0 Then path = Environ$("APPDATA") & "\mail2tdm\mail2tdm.log"

    Dim f As Integer
    f = FreeFile
    Open path For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd HH:nn:ss") & "  " & msg
    Close #f
End Sub


' Настройки нужны и объекту ThisOutlookSession - отдаём их наружу
Public Function GetMailboxName() As String
    If Not cfgLoaded Then LoadConfig
    GetMailboxName = cfgMailboxName
End Function

Public Function GetFolderPath() As String
    If Not cfgLoaded Then LoadConfig
    GetFolderPath = cfgFolderPath
End Function
