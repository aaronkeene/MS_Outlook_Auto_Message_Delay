Option Explicit

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    ' Delays messages to next working day as work start time
    
    Dim msg                         As Outlook.mailItem
    Dim msgSendDate                 As Date
    Dim msgDeferredDeliveryTime     As Date
    
    Const c_WorkHourStart   As Long = 7
    Const c_WorkHourEnd     As Long = 19
      
    Set msg = getActiveMessage()
    
    If msg Is Nothing Then
        Exit Sub
    End If
    
    ' Bypass for high importance items
    If msg.Importance = olImportanceHigh Then
        Exit Sub
    End If
      
    msgSendDate = Now()
    
    msgDeferredDeliveryTime = DeferredDeliveryTime(msgSendDate, c_WorkHourStart, c_WorkHourEnd)
                  
    If msgDeferredDeliveryTime > msgSendDate Then
        msg.DeferredDeliveryTime = msgDeferredDeliveryTime
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
    
    Dim insp                    As Outlook.Inspector
    Dim inline                  As Object
    
    If TypeOf Application.ActiveWindow Is Outlook.Inspector Then
        Set insp = Application.ActiveWindow
    End If

    If insp Is Nothing Then
        
        Set inline = Application.ActiveExplorer.ActiveInlineResponse
        
        If inline Is Nothing Then
            Exit Function
        End If
        
        Set getActiveMessage = inline
    
    Else
       
       Set insp = Application.ActiveInspector
       
       If insp.CurrentItem.Class = olMail Then
          Set getActiveMessage = insp.CurrentItem
       
       Else
         
         Exit Function
       
       End If

    End If
    
End Function
