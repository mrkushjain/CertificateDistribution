Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub Excel_to_whatsapp()
    Dim msg As String
    Dim whatsappWeb As Object
    Dim dispatchMsg As String
    Dim whatsappNum As String
    Dim lastRecipientNumber As Long
    Dim i As Long
    Const WAIT_TIME = 8
    Const COUNTRY_CODE = "91"
    Dim sendingStatus As String
    Dim sentStatus As String
    Dim errorStatus As String
    Dim onHoldStatus As String
    Dim status As String
    Dim currentDate As String
    
    
    
    
    
    sendingStatus = "Sending"
    sentStatus = "Sent"
    errorStatus = "Not exist"
    onHoldStatus = "On Hold"
    currentDate = Date & ""
    
    
    
    
    Dim picPath As String
     Dim chromePath As String
     
   
   
    
    lastRecipientNumber = Range("B" & Rows.Count).End(xlUp).Row
    For i = 2 To lastRecipientNumber
    status = Sheets("names").Range("D" & i).Value
    
    'Ignore sent and onhold entries
    If Not (status = onHoldStatus Or status = sentStatus Or status = errorStatus) Then

        Sheets("names").Range("D" & i).Value = sendingStatus
        Sheets("names").Range("E" & i).Value = currentDate
         
         
        On Error GoTo ErrorInSendingMsg:
            whatsappNum = COUNTRY_CODE & Sheets("names").Range("B" & i).Value
            
            msg = Sheets("names").Range("C" & i).Value
            dispatchMsg = "whatsapp://send?phone=" & whatsappNum & "&text=" & msg
    
    
            ShellExecute 0, vbNullString, dispatchMsg, vbNullString, vbNullString, vbNormalFocus
    
            picPath = "C:\Users\kusjain\Desktop\KushUnbackedup\iskconcert\" & Range("B" & i).Value & ".png"
            
            
            'reversing condition to not send images
            If (picPath <> "") Then
            
                    ActiveSheet.Pictures.Insert (picPath)
                    ActiveSheet.Shapes(1).Copy
                    Application.Wait (Now + TimeValue("00:00:0" & WAIT_TIME))
                    Call SendKeys("^V", True)
                    
                    ActiveSheet.Shapes(1).Delete
                      
            End If
                      
            Application.Wait (Now + TimeValue("00:00:0" & WAIT_TIME))
            Call SendKeys("{ENTER}", True)
            Sheets("names").Range("D" & i).Value = sentStatus
    End If
NextRow:
    Next i
      Exit Sub
ErrorInSendingMsg:
    
    Application.Wait (Now + TimeValue("00:00:0" & WAIT_TIME))
    Call SendKeys("{ESC}", True)
    Application.Wait (Now + TimeValue("00:00:0" & WAIT_TIME))
    Call SendKeys("{ESC}", True)
    Application.Wait (Now + TimeValue("00:00:0" & WAIT_TIME))
    Call SendKeys("{ENTER}", True)
    Sheets("names").Range("D" & i).Value = errorStatus
    

Exit Sub

       Resume NextRow
    
        

    
End Sub


