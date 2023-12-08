Attribute VB_Name = "Write_To_Page"
Option Explicit

Sub WriteToDocument(patientID As String)
    
    Dim response As Object
    Set response = ReadFromAPI(patientID)
    
    'Debug.Print JsonConverter.ConvertToJson(response, Whitespace:=2)
    
    Dim Json As Object
    Set Json = response
    
    Dim dateString As String
       
    
    ' Data
    
    'Each Patient DST Value
    Sheet1.Cells.ClearContents
    
    Sheet1.Cells(1, 1).Value = "Patient ID: " + patientID
    Sheet1.Cells(1, 3).Value = "Do NOT Edit Page Directly Contents Are Cleared After Each Search"
    Sheet1.Cells(1, 3).Font.Color = vbRed
    
    Dim patient As Dictionary
    For Each patient In Json("patient")
        
        'Debug.Print patient("patient_id")
        
        'Temporatry fix
        Dim temp As Long
        temp = 1
        
        Dim i As Long
        i = 4
        
        Dim j As Long
        j = 2
        
      
        Dim dailyData As Dictionary
        For Each dailyData In patient("dailyData")
            
            
            dateString = Left(dailyData("date"), 10)
            
            'Debug.Print (dateString)
            
        
            Sheet1.Cells(3, temp).Value = "Date: " + dateString
            
            Dim DST As Dictionary
            For Each DST In dailyData("DST")
            
                dateString = Mid(DST("timestamp"), 12, 8)
            
                'Debug.Print (dateString)
                
                Sheet1.Cells(i, temp).Value = "Time Stamp: " + dateString
                Sheet1.Cells(i, j).Value = DST("value")
                i = i + 1
            
            Next DST
            
            i = 4
            j = j + 2
            temp = temp + 2
   
        
        Next dailyData
        
 
    Next patient
    
    
    'Each Patient Speed Value
    Sheet3.Cells.ClearContents
    
    Sheet3.Cells(1, 1).Value = "Patient ID: " + patientID
    Sheet3.Cells(1, 3).Value = "Do NOT Edit Page Directly Contents Are Cleared After Each Search"
    Sheet3.Cells(1, 3).Font.Color = vbRed
    
    'Dim patient As Dictionary
    For Each patient In Json("patient")
        
        'Debug.Print patient("patient_id")
                
        'Temporatry fix
        'Dim temp As Long
        temp = 1
        
        'Dim i As Long
        i = 4
        
        'Dim j As Long
        j = 2
        
      
        'Dim dailyData As Dictionary
        For Each dailyData In patient("dailyData")
        
            dateString = Left(dailyData("date"), 10)
        
            Sheet3.Cells(3, temp).Value = "Date: " + dateString
            
            Dim SPEED As Dictionary
            For Each SPEED In dailyData("Speed")
            
                dateString = Mid(SPEED("timestamp"), 12, 8)
                
                Sheet3.Cells(i, temp).Value = "Time Stamp: " + dateString
                Sheet3.Cells(i, j).Value = SPEED("value")
                i = i + 1
            
            Next SPEED
            
            i = 4
            j = j + 2
            temp = temp + 2
   
        
        Next dailyData
        
 
    Next patient
    
    'Each Patient Asymetry Value
    Sheet4.Cells.ClearContents
    
    Sheet4.Cells(1, 1).Value = "Patient ID: " + patientID
    Sheet4.Cells(1, 3).Value = "Do NOT Edit Page Directly Contents Are Cleared After Each Search"
    Sheet4.Cells(1, 3).Font.Color = vbRed
    
    'Dim patient As Dictionary
    For Each patient In Json("patient")
        
        'Debug.Print patient("patient_id")
        
        'Temporatry fix
        'Dim temp As Long
        temp = 1
        
        'Dim i As Long
        i = 4
        
        'Dim j As Long
        j = 2
        
      
        'Dim dailyData As Dictionary
        For Each dailyData In patient("dailyData")
        
            dateString = Left(dailyData("date"), 10)
        
            Sheet4.Cells(3, temp).Value = "Date: " + dateString
            
            Dim ASYM As Dictionary
            For Each ASYM In dailyData("Asymetry")
            
                dateString = Mid(ASYM("timestamp"), 12, 8)
                
                Sheet4.Cells(i, temp).Value = "Time Stamp: " + dateString
                Sheet4.Cells(i, j).Value = ASYM("value")
                i = i + 1
            
            Next ASYM
            
            i = 4
            j = j + 2
            temp = temp + 2
   
        
        Next dailyData
        
 
    Next patient
    
    'Each Patient Speed Value
    Sheet5.Cells.ClearContents
    
    Sheet5.Cells(1, 1).Value = "Patient ID: " + patientID
    Sheet5.Cells(1, 3).Value = "Do NOT Edit Page Directly Contents Are Cleared After Each Search"
    Sheet5.Cells(1, 3).Font.Color = vbRed
    
    'Dim patient As Dictionary
    For Each patient In Json("patient")
        
        'Debug.Print patient("patient_id")
        
        'Temporatry fix
        'Dim temp As Long
        temp = 1
        
        'Dim i As Long
        i = 4
        
        'Dim j As Long
        j = 2
        
      
        'Dim dailyData As Dictionary
        For Each dailyData In patient("dailyData")
        
            dateString = Left(dailyData("date"), 10)
        
            Sheet5.Cells(3, temp).Value = "Date: " + dateString
            
            Dim STRIDE As Dictionary
            For Each STRIDE In dailyData("Stride")
            
                dateString = Mid(STRIDE("timestamp"), 12, 8)
                
                Sheet5.Cells(i, temp).Value = "Time Stamp: " + dateString
                Sheet5.Cells(i, j).Value = STRIDE("value")
                i = i + 1
            
            Next STRIDE
            
            i = 4
            j = j + 2
            temp = temp + 2
   
        
        Next dailyData
        
 
    Next patient

End Sub
Sub Button1_Click()

End Sub
