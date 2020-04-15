Attribute VB_Name = "modICD10"
Option Explicit

Private Sub TestICD()
    'Call searchICDbyName("chronic kidney disease")
    'Call searchICDbyCode("F10")
    Debug.Print getICDdesc("I12.9")
End Sub

Public Sub printICDbyName()

    Dim retval As String
    
    retval = InputBox("Enter an ICD description")
    If Len(retval) = 0 Then Exit Sub
    Call searchICDbyName(retval)

End Sub

Public Sub printICDbyCode()

    Dim retval As String
    
    retval = InputBox("Enter an ICD code")
    If Len(retval) = 0 Then Exit Sub
    Call searchICDbyCode(retval)

End Sub

Public Function getICDdesc(strICD As String) As String
    
    Dim lngTotal As Long
    
    ' JSON variables
    Dim sJSONString As String
    Dim vJSON
    Dim sState As String
    
    ' Call API
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", "https://clinicaltables.nlm.nih.gov/api/icd10cm/v3/search?maxList&sf=code&terms=" & strICD, True
        .Send
        Do Until .ReadyState = 4: DoEvents: Loop
        sJSONString = .ResponseText
    End With
    
    ' Parse JSON response
    JSON.Parse sJSONString, vJSON, sState
    ' Check response validity
    If sState = "Error" Then
        getICDdesc = "Invalid JSON response"
        Exit Function
    End If
    
    ' Extact data
    lngTotal = vJSON(0)
    
    ' Return description
    If lngTotal <> 1 Then
        getICDdesc = lngTotal & " results"
    Else
        getICDdesc = vJSON(3)(0)(1)
    End If
    
End Function

Public Sub searchICDbyName(strTerms As String)
    
    ' Response variables
    Dim lngTotal As Long
    'Dim arrICD()
    'Dim arrHeaderICD()
    'Dim aData()
    'Dim aHeader()
    
    ' JSON variables
    Dim sJSONString As String
    Dim vJSON
    Dim sState As String
    
    ' Call API
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", "https://clinicaltables.nlm.nih.gov/api/icd10cm/v3/search?maxList&sf=name&terms=" & strTerms, True
        .Send
        Do Until .ReadyState = 4: DoEvents: Loop
        sJSONString = .ResponseText
    End With
    
    ' Parse JSON response
    JSON.Parse sJSONString, vJSON, sState
    ' Check response validity
    If sState = "Error" Then
        MsgBox "Invalid JSON response", vbInformation
        Exit Sub
    End If
    
    ' Extact data
    lngTotal = vJSON(0)
    ' Convert JSON to 2D Array
    '   arrICD: 2D array but only 1st dimension has data
    '       Example: Debug.Print arrICD(0, 0)
    '   aData: 2D array; 2nd dimension determines code/name (0=Code, 1=Name)
    '       Example (Code): aData(0, 0)
    '       Example (Name): aData(0, 1)
    'JSON.ToArray vJSON(1), arrICD, arrHeaderICD
    'JSON.ToArray vJSON(3), aData, aHeader
    
    ' Let user know if results are missing or there are no results
    If lngTotal = 0 Then
        MsgBox "Search Yielded No Results", vbInformation, "Alert"
        Exit Sub
    ElseIf lngTotal > 500 Then
        MsgBox lngTotal - 500 & " results not displayed", vbInformation, "Refine Search Terms"
    End If
    ' Output to active sheet
    Call Output(ThisWorkbook.ActiveSheet, vJSON(3))
    MsgBox "Completed: " & lngTotal & " results", vbInformation
    
End Sub

Public Sub searchICDbyCode(strTerms As String)
    
    ' Response variables
    Dim lngTotal As Long
    'Dim arrICD()
    'Dim arrHeaderICD()
    'Dim aData()
    'Dim aHeader()
    
    ' JSON variables
    Dim sJSONString As String
    Dim vJSON
    Dim sState As String
    
    ' Call API
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", "https://clinicaltables.nlm.nih.gov/api/icd10cm/v3/search?maxList&sf=code&terms=" & strTerms, True
        .Send
        Do Until .ReadyState = 4: DoEvents: Loop
        sJSONString = .ResponseText
    End With
    
    ' Parse JSON response
    JSON.Parse sJSONString, vJSON, sState
    ' Check response validity
    If sState = "Error" Then
        MsgBox "Invalid JSON response", vbInformation
        Exit Sub
    End If
    
    ' Extact data
    lngTotal = vJSON(0)
    ' Convert JSON to 2D Array
    '   arrICD: 2D array but only 1st dimension has data
    '       Example: Debug.Print arrICD(0, 0)
    '   aData: 2D array; 2nd dimension determines code/name (0=Code, 1=Name)
    '       Example (Code): aData(0, 0)
    '       Example (Name): aData(0, 1)
    'JSON.ToArray vJSON(1), arrICD, arrHeaderICD
    'JSON.ToArray vJSON(3), aData, aHeader
    
    ' Let user know if results are missing or there are no results
    If lngTotal = 0 Then
        MsgBox "Search Yielded No Results", vbInformation, "Alert"
        Exit Sub
    ElseIf lngTotal > 500 Then
        MsgBox lngTotal - 500 & " results not displayed", vbInformation, "Refine Search Terms"
    End If
    ' Output to active sheet
    Call Output(ThisWorkbook.ActiveSheet, vJSON(3))
    MsgBox "Completed: " & lngTotal & " results", vbInformation
    
End Sub

Private Sub Output(oTarget As Worksheet, vJSON)
    
    Dim aData()
    Dim aHeader()
    
    ' Convert JSON to 2D Array
    JSON.ToArray vJSON, aData, aHeader
    ' Output to target worksheet range
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    With oTarget
        .Activate
        .Cells.Delete
        With .Cells(1, 1)
            .Resize(1, UBound(aHeader) - LBound(aHeader) + 1).Value = aHeader
            .Offset(1, 0).Resize( _
                    UBound(aData, 1) - LBound(aData, 1) + 1, _
                    UBound(aData, 2) - LBound(aData, 2) + 1 _
                ).Value = aData
        End With
        .Columns.AutoFit
    End With
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

