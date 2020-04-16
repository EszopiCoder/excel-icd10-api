VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formQueryICD 
   Caption         =   "Query ICD:"
   ClientHeight    =   5532
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4944
   OleObjectBlob   =   "formQueryICD.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formQueryICD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
    ' Set up userform
    Me.Caption = "Query ICD: " & ActiveWorkbook.Name
    listSearch.ColumnCount = 2
    listSearch.ColumnWidths = "60;170"
    listSearch.MultiSelect = fmMultiSelectExtended
    
    ' Set up sheets combo box
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        cboSheet.AddItem ws.Name
    Next ws
    cboSheet.AddItem "(Create new sheet)"
    cboSheet.Value = "Select Sheet"
    
    ' Add tooltip text
    optICD.ControlTipText = "International Classification of Disease code"
    optDesc.ControlTipText = "International Classification of Disease description"
    txtSearch.ControlTipText = "Search criteria (code or description)"
    btnAdd.ControlTipText = "Add search criteria"
    btnRemove.ControlTipText = "Remove search criteria"
    btnClear.ControlTipText = "Clear all search criteria"
    listSearch.ControlTipText = "Current search criteria"
    cboSheet.ControlTipText = "Select sheet to write to"
    btnSearch.ControlTipText = "Query search criteria and write to selected sheet"
    btnImport.ControlTipText = "Import CSV"
    btnExport.ControlTipText = "Export current search criteria to CSV"
    btnPreview.ControlTipText = "Preview current search criteria"
    
End Sub

Private Sub btnAdd_Click()
    
    If optICD.Value = False And optDesc.Value = False Then
        MsgBox "Invalid search type", vbInformation
        Exit Sub
    ElseIf Len(txtSearch.Text) = 0 Then
        MsgBox "Invalid search text", vbInformation
        Exit Sub
    Else
        ' Detect duplicates
        Dim i As Long
        For i = 0 To listSearch.ListCount - 1
            If listSearch.List(i, 0) = IIf(optICD.Value, optICD.Caption, optDesc.Caption) And _
                listSearch.List(i, 1) = txtSearch.Text Then
                    MsgBox "Duplicate search criteria detected", vbInformation
                    Exit Sub
            End If
        Next
        ' Add item to listbox
        With listSearch
            .AddItem
            .List(.ListCount - 1, 0) = IIf(optICD.Value, optICD.Caption, optDesc.Caption)
            .List(.ListCount - 1, 1) = txtSearch.Text
        End With
        ' Clear textbox and set focus
        txtSearch.Text = ""
        txtSearch.SetFocus
    End If
    
End Sub

Private Sub btnRemove_Click()
        
    If listSearch.ListCount = 0 Then
        MsgBox "No search criteria exist", vbInformation
        Exit Sub
    ElseIf listSearch.ListIndex = -1 Then
        MsgBox "No search criteria selected", vbInformation
        Exit Sub
    End If
    
    ' Get count of selected items
    Dim i As Long
    Dim ItemCount As Long
    For i = 0 To listSearch.ListCount - 1
        If listSearch.Selected(i) Then
            ItemCount = ItemCount + 1
        End If
    Next i
    
    ' Delete item if ItemCount=1 or exit sub if ItemCount=0
    If ItemCount = 1 Then
        listSearch.RemoveItem listSearch.ListIndex
        Exit Sub
    ElseIf ItemCount = 0 Then
        MsgBox "No search criteria selected", vbInformation
        Exit Sub
    End If
    
    ' Ask user to delete selected items
    If MsgBox(ItemCount & " search criteria are selected. Are you sure you want to delete them?", _
        vbYesNo + vbInformation) = vbNo Then Exit Sub
    
    ' Remove selected items
    For i = listSearch.ListCount - 1 To 0 Step -1
        If listSearch.Selected(i) Then
            listSearch.RemoveItem i
        End If
    Next i
    
End Sub

Private Sub btnClear_Click()

    If listSearch.ListCount = 0 Then
        MsgBox "No search criteria exist", vbInformation
        Exit Sub
    End If
    If MsgBox("Do you wish to clear the search criteria?", vbYesNo + vbInformation) = vbYes Then
        listSearch.Clear
    End If

End Sub

Private Sub btnSearch_Click()
    
    If listSearch.ListCount = 0 Then
        MsgBox "No search criteria exist", vbInformation
        Exit Sub
    ElseIf cboSheet.ListIndex = -1 Then
        MsgBox "Select sheet", vbInformation
        Exit Sub
    End If
    
    Me.Hide
    
    Call Search
    
    ' Add new sheet if option is selected
    If cboSheet.ListIndex = cboSheet.ListCount - 1 Then
        formLog.txtLog.Text = formLog.txtLog.Text & vbNewLine & "[" & Now & "] Adding new sheet"
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    End If
    ' Print to selected sheet
    formLog.txtLog.Text = formLog.txtLog.Text & vbNewLine & "[" & Now & "] Writing to sheet"
    Call PrintSheet(ICD, ActiveWorkbook.Worksheets(cboSheet.ListIndex + 1))
    ' Show completion message and log
    formLog.Hide
    MsgBox "Completed: " & ICD.Count & " results returned", vbInformation
    If log.Count > 0 Then Call MsgBoxDict(log)
    
End Sub

Private Sub listSearch_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With listSearch
        If optICD.Caption = .List(.ListIndex, 0) Then
            optICD.Value = True
        Else
            optDesc.Value = True
        End If
        txtSearch.Text = .List(.ListIndex, 1)
    End With
End Sub

Private Sub btnImport_Click()

    Dim strPath As String
    Dim strFilter As String
    Dim LineFromFile As String
    Dim LineItems() As String
    
    ' Open file dialog at the following path
    strFilter = modOpenDialog.OpenAddFilterItem(strFilter, "CSV (Comma delimited)", "*.csv")
    strPath = modOpenDialog.FileDialogOpen1(strPath, "Open CSV File", strFilter)
    If Len(strPath) = 0 Then Exit Sub
    
    ' Open and loop through csv file
    Open strPath For Input As #1
    Do Until EOF(1)
        Line Input #1, LineFromFile
        LineItems = Split(LineFromFile, ",")
        With listSearch
            .AddItem
            .List(.ListCount - 1, 0) = LineItems(0)
            .List(.ListCount - 1, 1) = LineItems(1)
        End With
    Loop
    Close #1

    MsgBox Dir(strPath) & " loaded successfully.", vbInformation
    
End Sub

Private Sub btnExport_Click()

    Dim strFilter As String
    Dim strPath As String
    Dim i As Long
    
    ' Check if listSearch is empty
    If listSearch.ListCount = 0 Then
        MsgBox "No search criteria exist", vbInformation
        Exit Sub
    End If
    
    ' Save file dialog at the following path
    strFilter = modSaveDialog.SaveAddFilterItem(strFilter, "CSV (Comma delimited)", "*.csv")
    strPath = modSaveDialog.FileDialogSave1("", "", "Save CSV File", strFilter)
    If Len(strPath) = 0 Then Exit Sub
    
    ' Save listbox to CSV file
    Open strPath For Output As #2
        With listSearch
            For i = 0 To .ListCount - 1
                Print #2, .List(i, 0) & "," & .List(i, 1)
            Next i
        End With
    Close #2
    
    MsgBox Dir(strPath) & " saved successfully.", vbInformation
    
End Sub

Private Sub btnPreview_Click()

    If listSearch.ListCount = 0 Then
        MsgBox "No search criteria exist", vbInformation
        Exit Sub
    End If
    
    Call Search
    
    ' Show preview
    formLog.Hide
    Call MsgBoxDict(ICD, True)
    If log.Count > 0 Then Call MsgBoxDict(log)
    
End Sub

Private Sub Search()

    ' Clear dictionaries
    formLog.Show vbModeless
    formLog.Caption = "Progress Log"
    formLog.txtLog.Text = "[" & Now & "] Clearing dictionaries"
    If ExistsDict(ICD) Then Call ClearDict(ICD)
    If ExistsDict(log) Then Call ClearDict(log)
    ' Fill dictionary
    formLog.txtLog.Text = formLog.txtLog.Text & vbNewLine & "[" & Now & "] Filling dictionary"
    Dim i As Long
    For i = listSearch.ListCount - 1 To 0 Step -1
        Call SearchICD(IIf(listSearch.List(i, 0) = optICD.Caption, 0, 1), listSearch.List(i, 1))
    Next i
    ' Let user know if there are no results
    If ICD.Count = 0 Then
        formLog.Hide
        MsgBox "Search yielded no results", vbInformation
        Call MsgBoxDict(log)
        Exit Sub
    End If

End Sub
