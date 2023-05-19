Attribute VB_Name = "m_SourceToTargetMappingDoc"
Function SQLFindFullyQualifiedMappingDelimited(SearchArray As Range, FindRange As Range, Delimiter As String, SearchColumnNumber As Integer, TableNameColumnNumber As Integer, FieldNameColumnNumber) As String

    Dim splitFindValue() As String
    Dim returnString As String
    Dim foundRow As Range
    returnString = ""
    splitFindValue = Split(FindRange.Value, Delimiter)
    Dim searchRange As Range
    
    'Debug.Print SearchArray.Address
    'Debug.Print FindRange.Address
    'Debug.Print FindRange.Value
    
    For Each searchString In splitFindValue
        Set searchRange = SearchArray.Columns(SearchColumnNumber)
        'Debug.Print searchRange.Address
        Set foundRow = searchRange.Find(searchString)
        'Debug.Print foundRow.Address
        Set foundRow = foundRow.Offset(0, -1 * (SearchColumnNumber - 1))
        'Debug.Print foundRow.Address
        Set foundRow = foundRow.Resize(1, SearchArray.Columns.Count)
        'Debug.Print foundRow.Address
        returnString = returnString & foundRow.Cells(1, TableNameColumnNumber)
        returnString = returnString & "."
        returnString = returnString & foundRow.Cells(1, FieldNameColumnNumber)
        returnString = returnString & Chr(13)
    Next searchString
        
    SQLFindFullyQualifiedMappingDelimited = returnString

End Function
Sub testing()
Attribute testing.VB_ProcData.VB_Invoke_Func = " \n14"
ExpandMappingSheet (395)
'-----------------------------------------------------
'    Dim wk As Workbook
'    Dim sh As Worksheet
'    Set wk = ActiveWorkbook
'    For Each sh In wk.Worksheets
'        If sh.Name <> "Version History" _
'            And sh.Name <> "Index" _
'            And sh.Name <> "Template" _
'            And sh.Visible = xlSheetVisible Then
'
'            If Not IsStringInColumn(sh.Name) Then
'                AddToIndex sh
'            End If
'        End If
'    Next sh
'---------------------------------------------------------
'Dim hl As hyperlink
'Set a = ActiveSheet.Hyperlinks
'For Each hl In a
'    Debug.Print hl.Range.Address; ":"""; hl.Range.Value; """<"; hl.SubAddress; ">"
'Next hl
'----------------------------------------------------------

End Sub
Sub callAddToIndex()
Attribute callAddToIndex.VB_ProcData.VB_Invoke_Func = "I\n14"
    AddToIndex ActiveSheet
End Sub
Function IsStringInColumn(inputString As String) As Boolean
    Dim wk As Workbook
    Dim idxSh As Worksheet
    Dim idxRng As Range
    Dim foundRng As Range
    Dim result As Boolean
    result = False
    
    Set wk = ActiveWorkbook
    Set idxSh = wk.Sheets("Index")
    Set idxRng = idxSh.UsedRange.Columns(4)
    Set foundRng = idxRng.Find(inputString, LookIn:=xlValues, lookat:=xlWhole)
    
    If Not foundRng Is Nothing Then
        result = True
    End If
    
    IsStringInColumn = result
    
End Function
Function FindStringInColumn(inputString As String) As Range
    Dim wk As Workbook
    Dim idxSh As Worksheet
    Dim idxRng As Range
    Dim foundRng As Range
    Dim result As Boolean
    result = False
    
    Set wk = ActiveWorkbook
    Set idxSh = wk.Sheets("Index")
    Set idxRng = idxSh.UsedRange.Columns(4)
    Set foundRng = idxRng.Find(inputString, LookIn:=xlValues, lookat:=xlWhole)
    
    Set FindStringInColumn = foundRng
        
End Function
Function SplitPascalCase(inputString As String) As String()
    Dim output As String
    Dim i As Integer
    
    output = ""
    
    'Loop through each character in the input string
    For i = 1 To Len(inputString)
        'If the current character is uppercase and not the first character, add a space before it
        If i > 1 And Mid(inputString, i, 1) Like "[A-Z]" Then
            output = output & "|"
        End If
        
        'Add the current character to the output string
        output = output & Mid(inputString, i, 1)
    Next i
    
    SplitPascalCase = Split(output, "|")
End Function
Sub CopySheet()
Attribute CopySheet.VB_ProcData.VB_Invoke_Func = "C\n14"
    ActiveWorkbook.ActiveSheet.Copy Before:=ActiveWorkbook.Sheets("Template")
End Sub
Sub AddToIndex(curSh As Worksheet)
    Dim wkbk As Workbook
    Dim idxSh As Worksheet
    'Dim curSh As Worksheet
    Dim idxRng As Range
    Dim curRng As Range
    Dim srcEntity As String
    Dim trgEntity As String
    
    Set wkbk = ActiveWorkbook
    Set idxSh = wkbk.Sheets("Index")
    'Set curSh = wkbk.ActiveSheet
    Set curRng = curSh.Cells(2, 2)
    Set idxRng = idxSh.UsedRange.End(xlDown)
    
    srcEntity = curSh.Cells(5, 2).Value
    trgEntity = curSh.Cells(5, 10).Value
    
    If IsStringInColumn(curSh.Name) Then
        Set idxRng = idxSh.Cells(FindStringInColumn(curSh.Name).Row, 1)
    Else
        idxRng.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Set idxRng = idxRng.Offset(-1, 0)
    End If
    
    If curSh.Name = "Template" Then
        idxSh.Cells(idxRng.Row, 1).Value = "NA"
        idxSh.Cells(idxRng.Row, 2).Value = "NA"
        idxSh.Cells(idxRng.Row, 3).Value = "NA"
        idxSh.Cells(idxRng.Row, 4).Value = curSh.Name
        idxSh.Hyperlinks.Add anchor:=idxSh.Cells(idxRng.Row, 4), Address:="", SubAddress:=curSh.Name & "!B2", TextToDisplay:=curSh.Name
    Else
        idxSh.Cells(idxRng.Row, 1).Value = SplitPascalCase(curSh.Name)(1)
        idxSh.Cells(idxRng.Row, 2).Value = srcEntity
        idxSh.Cells(idxRng.Row, 3).Value = trgEntity
        idxSh.Cells(idxRng.Row, 4).Value = curSh.Name
        idxSh.Hyperlinks.Add anchor:=idxSh.Cells(idxRng.Row, 4), Address:="", SubAddress:=curSh.Name & "!B2", TextToDisplay:=curSh.Name
    End If
    
    curRng.Hyperlinks.Delete
    curSh.Hyperlinks.Add anchor:=curRng, Address:="", SubAddress:= _
        "Index!D" & idxRng.Row, TextToDisplay:="Link back to Index"
    
    'idxSh.Activate
    'idxRng.Select

End Sub
Function FindRange(ByVal searchStr As String, searchRng As Range) As Range
    Set FindRange = searchRng.Find( _
        What:=searchStr, _
        LookIn:=xlValues, _
        lookat:=xlWhole _
    )
End Function
Sub AddMappingHyperlink()
Attribute AddMappingHyperlink.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim ce As Range
    Dim sh As Worksheet
    Dim searchRng As Range
    Dim linkRng As Range
    'Dim searchStr As String
    
    Set sh = ActiveSheet
    Set searchRng = sh.Range("H5")
    Set searchRng = sh.Range(searchRng, searchRng.End(xlDown))
           
    'Debug.Print searchRng.Address
               
    For Each ce In Selection.Cells
        'Debug.Print ce.Address(False, False)
        If InStr(ce.Value, ", ") > 0 Then
            strArr = Split(ce.Value, ", ")
            For Each searchStr In strArr
                If linkRng Is Nothing Then
                    Set linkRng = FindRange(searchStr, searchRng)
                Else
                    Set linkRng = Union(linkRng, FindRange(searchStr, searchRng))
                End If
            Next
        Else
            Set linkRng = FindRange(ce.Value, searchRng)
        End If
        'Debug.Print linkRng.Address
        sh.Hyperlinks.Add _
            anchor:=ce, _
            Address:="", _
            SubAddress:=linkRng.Address(False, False, xlA1, True) '"DimEmployee!B13:E13"
    Next ce
    
    Set ce = Selection
    
    ce.Cells(ce.Rows.Count, 1).Offset(1, 0).Copy
    ce.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub
Sub ExpandMappingSheet(ul As Integer)
    Dim wk As Workbook: Set wk = ActiveWorkbook
    Dim ws As Worksheet: Set ws = wk.ActiveSheet
    Dim rng As Range
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'Dim ul As Integer: ul = 650
    
    For i = 1 To ul
        Set rng = ws.Cells(i + 4, 8)
        If rng.Value = "" Then
            rng.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            rng.Offset(-1).Value = 10 * i
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
