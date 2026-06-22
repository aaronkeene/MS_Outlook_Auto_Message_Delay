Option Explicit

' ============================================================
' SETTINGS — edit these to change behavior
' ============================================================
Const START_OF_DAY              As Long = 7             ' Work day start hour (0-23)
Const END_OF_DAY                As Long = 19            ' Work day end hour (0-23)
Const DELAY_ENABLED             As Boolean = True       ' False = disable all delays
Const BYPASS_HIGH_IMPORTANCE    As Boolean = True       ' True = send urgent mail immediately
Const CHECK_HOLIDAYS            As Boolean = True       ' True = check Outlook calendar for holidays
Const HOLIDAY_CATEGORY          As String = "Holiday"   ' Outlook calendar category name for holidays
Const MAX_HOLIDAY_SEARCH        As Long = 14            ' Max days to search forward past a holiday chain



' ============================================================
' ENTRY POINT
' ============================================================

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    If Not DELAY_ENABLED Then
        Exit Sub
    End If

    If Item.Class <> olMail Then
        Exit Sub
    End If

    Dim msg As Outlook.MailItem
    Set msg = Item

    Call DelayMessageToNextWorkDay(msg, START_OF_DAY, END_OF_DAY, BYPASS_HIGH_IMPORTANCE, False, CHECK_HOLIDAYS)

End Sub



' ============================================================
' CORE DELAY LOGIC
' ============================================================

Private Sub DelayMessageToNextWorkDay(msg As Outlook.MailItem, _
                                      WorkHourStart As Long, _
                                      WorkHourEnd As Long, _
                                      Optional BypassForHighImportance As Boolean = True, _
                                      Optional UserConfirmDelay As Boolean = False, _
                                      Optional CheckHolidays As Boolean = True)

    If msg Is Nothing Then Exit Sub

    If BypassForHighImportance And msg.Importance = olImportanceHigh Then
        Exit Sub
    End If

    Dim msgSendDateTime     As Date
    Dim msgDeferredDateTime As Date

    msgSendDateTime = Now()
    msgDeferredDateTime = DeferredDeliveryTime(msgSendDateTime, WorkHourStart, WorkHourEnd, CheckHolidays)

    If msgDeferredDateTime <= msgSendDateTime Then Exit Sub

    If Not UserConfirmDelay Then
        msg.DeferredDeliveryTime = msgDeferredDateTime
        Exit Sub
    End If

    If MsgBox("Delay delivery until " & msgDeferredDateTime & "?", vbYesNo, "Delay Delivery?") = vbYes Then
        msg.DeferredDeliveryTime = msgDeferredDateTime
    End If

End Sub



Private Function DeferredDeliveryTime(MessageSendDate As Date, _
                                      WorkHourStart As Long, _
                                      WorkHourEnd As Long, _
                                      CheckHolidays As Boolean) As Date

    Dim msgDay      As Long
    Dim msgHour     As Long
    Dim msgMinute   As Long
    Dim sendDate    As Date

    msgDay = Weekday(MessageSendDate, vbSunday)
    msgHour = Hour(MessageSendDate)
    msgMinute = Minute(MessageSendDate)

    If msgDay = vbSunday Then
        sendDate = BuildStartOfDay(MessageSendDate, 1, WorkHourStart, msgHour, msgMinute)
        DeferredDeliveryTime = ResolveHolidays(sendDate, WorkHourStart, CheckHolidays)
        Exit Function
    End If

    If msgDay = vbSaturday Then
        sendDate = BuildStartOfDay(MessageSendDate, 2, WorkHourStart, msgHour, msgMinute)
        DeferredDeliveryTime = ResolveHolidays(sendDate, WorkHourStart, CheckHolidays)
        Exit Function
    End If

    If msgDay = vbFriday And msgHour >= WorkHourEnd Then
        sendDate = BuildStartOfDay(MessageSendDate, 3, WorkHourStart, msgHour, msgMinute)
        DeferredDeliveryTime = ResolveHolidays(sendDate, WorkHourStart, CheckHolidays)
        Exit Function
    End If

    If msgHour < WorkHourStart Then
        sendDate = DateAdd("h", WorkHourStart - msgHour, MessageSendDate)
        sendDate = DateAdd("n", -msgMinute, sendDate)
        sendDate = DateAdd("s", -Second(sendDate), sendDate)
        DeferredDeliveryTime = ResolveHolidays(sendDate, WorkHourStart, CheckHolidays)
        Exit Function
    End If

    If msgHour >= WorkHourEnd Then
        sendDate = BuildStartOfDay(MessageSendDate, 1, WorkHourStart, msgHour, msgMinute)
        DeferredDeliveryTime = ResolveHolidays(sendDate, WorkHourStart, CheckHolidays)
        Exit Function
    End If

    If CheckHolidays And IsHoliday(MessageSendDate) Then
        sendDate = BuildStartOfDay(MessageSendDate, 1, WorkHourStart, msgHour, msgMinute)
        DeferredDeliveryTime = ResolveHolidays(sendDate, WorkHourStart, CheckHolidays)
        Exit Function
    End If

    DeferredDeliveryTime = MessageSendDate

End Function



Private Function ResolveHolidays(CandidateDate As Date, WorkHourStart As Long, CheckHolidays As Boolean) As Date

    If Not CheckHolidays Then
        ResolveHolidays = CandidateDate
        Exit Function
    End If

    Dim resolvedDate    As Date
    Dim dayOfWeek       As Long
    Dim iterations      As Long

    resolvedDate = CandidateDate
    iterations = 0

    Do While iterations < MAX_HOLIDAY_SEARCH

        dayOfWeek = Weekday(resolvedDate, vbSunday)

        If dayOfWeek = vbSunday Then
            resolvedDate = DateAdd("d", 1, DateValue(resolvedDate)) + TimeSerial(WorkHourStart, 0, 0)
            iterations = iterations + 1

        ElseIf dayOfWeek = vbSaturday Then
            resolvedDate = DateAdd("d", 2, DateValue(resolvedDate)) + TimeSerial(WorkHourStart, 0, 0)
            iterations = iterations + 1

        ElseIf IsHoliday(resolvedDate) Then
            resolvedDate = DateAdd("d", 1, DateValue(resolvedDate)) + TimeSerial(WorkHourStart, 0, 0)
            iterations = iterations + 1

        Else
            Exit Do

        End If

    Loop

    ResolveHolidays = resolvedDate

End Function



Private Function IsHoliday(CheckDate As Date) As Boolean

    Dim outlookCalendar     As Outlook.Folder
    Dim calendarItems       As Outlook.Items
    Dim filteredItems       As Outlook.Items
    Dim checkDateOnly       As Date
    Dim filterString        As String

    IsHoliday = False

    On Error GoTo IsHoliday_Error

    checkDateOnly = DateValue(CheckDate)

    Set outlookCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set calendarItems = outlookCalendar.Items

    calendarItems.IncludeRecurrences = True
    calendarItems.Sort "[Start]"

    filterString = "[Start] >= """ & Format(checkDateOnly, "mm/dd/yyyy hh:mm AM/PM") & """" & _
                   " AND [Start] < """ & Format(checkDateOnly + 1, "mm/dd/yyyy hh:mm AM/PM") & """" & _
                   " AND [Categories] = """ & HOLIDAY_CATEGORY & """"

    Set filteredItems = calendarItems.Restrict(filterString)

    If filteredItems.Count > 0 Then
        IsHoliday = True
    End If

    Exit Function

IsHoliday_Error:
    IsHoliday = False

End Function



Private Function BuildStartOfDay(BaseDate As Date, _
                                  DaysForward As Long, _
                                  WorkHourStart As Long, _
                                  CurrentHour As Long, _
                                  CurrentMinute As Long) As Date

    Dim newDate As Date

    newDate = DateAdd("d", DaysForward, BaseDate)
    newDate = DateAdd("h", WorkHourStart - CurrentHour, newDate)
    newDate = DateAdd("n", -CurrentMinute, newDate)
    newDate = DateAdd("s", -Second(newDate), newDate)

    BuildStartOfDay = newDate

End Function


