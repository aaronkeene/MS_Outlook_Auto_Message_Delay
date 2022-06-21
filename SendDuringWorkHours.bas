Option Explicit
Const c_WorkHourStart           As Long = 7
Const c_WorkHourEnd             As Long = 19


Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    
    'This function runs on every mail item sent
    
    Dim msg                         As Outlook.mailItem

    Set msg = getActiveMessage()
    
    Call DelayMessageToNextWorkDay(msg, c_WorkHourStart, c_WorkHourEnd)        
        
End Sub
    
    
Private Sub DelayMessageToNextWorkDay(msg As Outlook.mailItem, _
                                      WorkHourStart As Long, _
                                      WorkHourEnd As Long, _
                                      Optional BypassForHighImportance As Boolean = True, _
                                      Optional UserConfirmDelay As Boolean = False)
    
    'Delay mail unit start of next workday (Monday)
        
    Dim msgSendDateTime             As Date
    Dim msgDeferredDateTime         As Date
    
    If msg Is Nothing Then
        Exit Sub
    End If
    
    If BypassForHighImportance And msg.Importance = olImportanceHigh Then
        Exit Sub
    End If
    
    msgSendDateTime = Now()
    
    msgDeferredDateTime = DeferredDeliveryTime(msgSendDateTime, WorkHourStart, WorkHourEnd)
    
    If msgDeferredDateTime < msgSendDateTime Then
        Exit Sub
    End If
    
    If Not UserConfirmDelay Then
        msg.DeferredDeliveryTime = msgDeferredDateTime
        Exit Sub
    End If
    
    If MsgBox("Do you want to delay delivery until " & msgDeferredDateTime & "?", vbYesNo, "Delay Delivery?") = vbYes Then
        msg.DeferredDeliveryTime = msgDeferredDateTime
        Exit Sub
    End If
End Sub
                        

Private Function DeferredDeliveryTime(MessageSendDate As Date, WorkHourStart As Long, WorkHourEnd As Long) As Date
    
    ' Determines the deferred delivery time based on the start and end work hours
    ' Note Sunday = 1
    
    Dim msgDay As Long
    Dim msgHour As Long
    Dim msgMinute As Long
    Dim sendDate As Date
    
    msgDay = Weekday(MessageSendDate, vbSunday)
    msgHour = Hour(MessageSendDate)
    msgMinute = Minute(MessageSendDate)
  
    ' Check if Sunday
    If msgDay = 1 Then
        sendDate = DateAdd("d", 1, MessageSendDate)
        sendDate = DateAdd("h", WorkHourStart - msgHour, sendDate)
        sendDate = DateAdd("n", -msgMinute, sendDate)
        sendDate = DateAdd("s", -Second(sendDate), sendDate)
        DeferredDeliveryTime = sendDate
        Exit Function
    End If
    
    'Check if Saturday
    If msgDay = 7 Then
        sendDate = DateAdd("d", 2, MessageSendDate)
        sendDate = DateAdd("h", WorkHourStart - msgHour, sendDate)
        sendDate = DateAdd("n", -msgMinute, sendDate)
        sendDate = DateAdd("s", -Second(sendDate), sendDate)
        DeferredDeliveryTime = sendDate
        Exit Function
    End If
    
    'Check if Friday after work hours
    If msgDay = 6 And msgHour >= WorkHourEnd Then
        sendDate = DateAdd("d", 3, MessageSendDate)
        sendDate = DateAdd("h", WorkHourStart - msgHour, sendDate)
        sendDate = DateAdd("n", -msgMinute, sendDate)
        sendDate = DateAdd("s", -Second(sendDate), sendDate)
        DeferredDeliveryTime = sendDate
        Exit Function
    End If
    
    ' Check if before work hours
    If msgHour < WorkHourStart Then
        sendDate = DateAdd("h", WorkHourStart - msgHour, MessageSendDate)
        sendDate = DateAdd("n", -msgMinute, sendDate)
        DeferredDeliveryTime = sendDate
        Exit Function
    End If

    ' Check if after work hours
    If msgHour >= WorkHourEnd Then
        sendDate = DateAdd("h", (24 + WorkHourStart) - msgHour, MessageSendDate)
        sendDate = DateAdd("n", -msgMinute, sendDate)
        sendDate = DateAdd("s", -Second(sendDate), sendDate)
        DeferredDeliveryTime = sendDate
        Exit Function
    End If

    ' No Delay
    DeferredDeliveryTime = MessageSendDate
    
End Function



Private Function getActiveMessage() As Outlook.mailItem
    
    If Application.ActiveWindow Is Outlook.Inspectors Then
        
        If Application.ActiveWindow.CurrentItem.Class = olMail Then: Set getActiveMessage = Application.ActiveInspector.CurrentItem
        Exit Function
        
    End If
    
    If Not Application.ActiveExplorer.ActiveInlineResponse Is Nothing Then
        
        Set getActiveMessage = Application.ActiveExplorer.ActiveInlineResponse
        Exit Function
        
    End If
    
End Function
