Attribute VB_Name = "modICD10_dict"
Option Explicit
Public ICD As Object
Public log As Object
Public Const SearchCode As Integer = 0
Public Const SearchName As Integer = 1

Private Sub TestQuery()
    Set log = Nothing
    Set ICD = Nothing
    ' Initialize dictionary
    Call InitializeDict(ICD)
    ' Search ICDs
    Call SearchICD(SearchName, "this is a error") ' Throw error on purpose
    Call SearchICD(SearchCode, "F10")
    Call SearchICD(SearchName, "alcohol")
    ' Print ICD and log
    Call PrintSheet(ICD, ThisWorkbook.ActiveSheet)
    'Call PrintDict(log)
    ' Show completion message
    MsgBox "Completed: " & ICD.Count & " results returned"
    Call MsgBoxDict(log)
End Sub

Public Sub SearchICD(intSearch As Integer, strTerms As String)
    
    ' Create dictionaries if they don't exist
    If ExistsDict(ICD) = False Then
        Call InitializeDict(ICD)
    End If
    If ExistsDict(log) = False Then
        Call InitializeDict(log)
    End If
    
    ' JSON variables
    Dim strAPI As String
    Dim sJSONString As String
    Dim vJSON
    Dim sState As String
    Dim oItem
    
    ' Response variables
    Dim lngTotal As Long
    'Dim arrICD()
    'Dim arrHeaderICD()
    'Dim aData()
    'Dim aHeader()
    
    ' Generate API link
    Select Case intSearch
        Case 0
            strAPI = "https://clinicaltables.nlm.nih.gov/api/icd10cm/v3/search?maxList&sf=code&terms="
        Case 1
            strAPI = "https://clinicaltables.nlm.nih.gov/api/icd10cm/v3/search?maxList&sf=name&terms="
        Case Else
            Call AddItemDict(log, "intSearch=" & intSearch & "; strTerms=" & strTerms, "Invalid search type")
            Exit Sub
    End Select
    
    ' Call API
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", strAPI & strTerms, True
        .Send
        Do Until .ReadyState = 4: DoEvents: Loop
        sJSONString = .ResponseText
    End With
    
    ' Parse JSON response
    JSON.Parse sJSONString, vJSON, sState
    ' Check response validity
    If sState = "Error" Then
        Call AddItemDict(log, "intSearch=" & intSearch & "; strTerms=" & strTerms, "Invalid JSON response")
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
        Call AddItemDict(log, "intSearch=" & intSearch & "; strTerms=" & strTerms, "Search yielded no results")
        Exit Sub
    ElseIf lngTotal > 500 Then
        Call AddItemDict(log, "intSearch=" & intSearch & "; strTerms=" & strTerms, lngTotal - 500 & " results not displayed; refine search criteria")
    End If
    
    ' Check if dictionary exists
    If ExistsDict(ICD) = False Then
        Call AddItemDict(log, "intSearch=" & intSearch & "; strTerms=" & strTerms, "Dictionary does not exist")
        Exit Sub
    End If
    ' Add all items to dictionary
    For Each oItem In vJSON(3)
        Call AddItemDict(ICD, CStr(oItem(0)), CStr(oItem(1)), True)
    Next
    
End Sub

Public Sub InitializeDict(objDict As Object)

    ' Create dictionary object (erases current dictionary)
    Set objDict = CreateObject("Scripting.Dictionary")
    objDict.CompareMode = vbTextCompare 'Not case-sensitive
    
End Sub

Public Sub AddItemDict(objDict As Object, strKey As String, _
    strItem As String, Optional boolOverwrite As Boolean = False)
    
    ' Check if dictionary exists
    If ExistsDict(objDict) = False Then
        MsgBox "Dictionary does not exist", vbInformation, "AddItemDict()"
        Exit Sub
    End If
    ' Add item to dictionary if the key doesn't exist or boolOverwrite = True
    If objDict.Exists(strKey) = True And boolOverwrite = False Then
        Exit Sub
    Else
        objDict(strKey) = strItem
    End If
    
End Sub

Public Sub RemoveItemDict(objDict As Object, strKey)

    ' Check if dictionary exists
    If ExistsDict(objDict) = False Then
        MsgBox "Dictionary does not exist", vbInformation, "RemoveItemDict()"
        Exit Sub
    End If
    ' Remove item from dictionary if the key exists
    If objDict.Exists(strKey) Then
        objDict.Remove strKey
    End If
    
End Sub

Public Sub ClearDict(objDict As Object)

    ' Check if dictionary exists and clear it if it does
    If ExistsDict(objDict) = False Then
        MsgBox "Dictionary does not exist", vbInformation, "RemoveItemDict()"
        Exit Sub
    Else
        objDict.RemoveAll
    End If
    
End Sub

Public Sub PrintSheet(objDict As Object, oTarget As Worksheet)

    ' Check if dictionary exists
    If ExistsDict(objDict) = False Then
        MsgBox "Dictionary does not exist", vbInformation, "PrintSheet()"
        Exit Sub
    End If
    ' Check if dictionary is empty
    If objDict.Count = 0 Then
        MsgBox "Dictionary is empty", vbInformation, "PrintSheet()"
        Exit Sub
    End If
    ' Write keys and items to sheet
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    With oTarget
        .Activate
        .Cells.Delete
        With .Cells(1, 1)
            .Resize(objDict.Count, 1).Value = WorksheetFunction.Transpose(objDict.keys)
            .Offset(0, 1).Resize(objDict.Count, 1).Value = WorksheetFunction.Transpose(objDict.Items)
        End With
        .Columns.AutoFit
    End With
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Public Sub PrintDict(objDict As Object)

    ' Check if dictionary exists
    If ExistsDict(objDict) = False Then
        MsgBox "Dictionary does not exist", vbInformation, "PrintDict()"
        Exit Sub
    End If
    ' Check if dictionary is empty
    If objDict.Count = 0 Then
        MsgBox "Dictionary is empty", vbInformation, "PrintDict()"
        Exit Sub
    End If
    ' Read through dictionary
    Dim k As Variant
    For Each k In objDict.keys
        ' Print key and value
        Debug.Print k, objDict(k)
    Next
    
End Sub

Public Sub MsgBoxDict(objDict As Object)

    ' Check if dictionary exists
    If ExistsDict(objDict) = False Then
        MsgBox "Dictionary does not exist", vbInformation, "MsgBoxDict()"
        Exit Sub
    End If
    ' Check if dictionary is empty
    If objDict.Count = 0 Then
        MsgBox "Dictionary is empty", vbInformation, "MsgBoxDict()"
        Exit Sub
    End If
    ' Read through dictionary
    Dim strMsg As String
    Dim k As Variant
    For Each k In objDict.keys
        ' Store key and value
        strMsg = strMsg & k & vbTab & objDict(k) & vbNewLine
    Next
    ' Show log
    formLog.Caption = "Error Log: " & objDict.Count & " result(s)"
    formLog.txtLog.Text = "[" & Now & "]" & vbNewLine & strMsg & "(end of log)"
    formLog.Show vbModeless
    
End Sub

Public Function ExistsDict(objDict As Object) As Boolean

    If objDict Is Nothing Then
        ExistsDict = False
    Else
        ExistsDict = True
    End If
    
End Function
