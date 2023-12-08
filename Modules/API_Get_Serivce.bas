Attribute VB_Name = "API_Get_Serivce"
Option Explicit

Function ReadFromAPI(patientID As String) As Object

    Dim URL As String
    Dim parameters As String
    
    URL = "http://xcodebackend.onrender.com/api/patient-alldaily/" + patientID
    parameters = ""
    
    ' Send Request
        
    Dim request As New WinHttpRequest
    
    request.Open "Get", URL & parameters
    'request.SetRequestHeader "[key]", "[Value]"
    
    request.Send
    
    If request.Status <> 200 Then
        MsgBox "Error: " & request.ResponseText
        Exit Function
    End If
    
    
    'Dim ReadFromAPI As Object
    Set ReadFromAPI = JsonConverter.ParseJson(request.ResponseText)
    'Debug.Print JsonConverter.ConvertToJson(response, Whitespace:=2)
       
    
End Function


