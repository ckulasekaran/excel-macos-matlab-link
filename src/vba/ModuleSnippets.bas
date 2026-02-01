Attribute VB_Name = "ModuleSnippets"
Option Explicit

Public Function BuildSnippet(ByVal snippetId As String) As String
    Select Case snippetId
        Case "readmatrix"
            BuildSnippet = BuildReadmatrixSnippet()
        Case Else
            BuildSnippet = ""
    End Select
End Function

Private Function BuildReadmatrixSnippet() As String
    Dim varName As String
    Dim workbookPath As String
    Dim sheetName As String
    Dim rangeAddress As String
    Dim targetRange As Range

    varName = PromptVariableName()
    If varName = "" Then
        BuildReadmatrixSnippet = ""
        Exit Function
    End If

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation
        BuildReadmatrixSnippet = ""
        Exit Function
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a valid cell range first.", vbExclamation
        BuildReadmatrixSnippet = ""
        Exit Function
    End If

    Set targetRange = Selection
    If targetRange.Areas.Count > 1 Then
        MsgBox "Please select a single contiguous range. Non-contiguous ranges are not supported yet.", vbExclamation
        BuildReadmatrixSnippet = ""
        Exit Function
    End If
    workbookPath = GetWorkbookPath()
    If workbookPath = "" Then
        MsgBox "Please save the workbook to use this snippet.", vbExclamation
        BuildReadmatrixSnippet = ""
        Exit Function
    End If
    sheetName = targetRange.Worksheet.Name
    rangeAddress = targetRange.Address(False, False)

    BuildReadmatrixSnippet = varName & " = readmatrix('" & EscapeMatlabString(workbookPath) & "', " & _
        "'Sheet','" & EscapeMatlabString(sheetName) & "', 'Range','" & rangeAddress & "');"
End Function

Private Function PromptVariableName() As String
    Dim inputValue As String
    inputValue = InputBox("Variable name for readmatrix output:", "Matlab Variable")
    inputValue = Trim(inputValue)

    If inputValue = "" Then
        PromptVariableName = ""
        Exit Function
    End If

    If Not IsValidMatlabIdentifier(inputValue) Then
        MsgBox "Invalid Matlab variable name.", vbExclamation
        PromptVariableName = ""
        Exit Function
    End If

    PromptVariableName = inputValue
End Function

Private Function IsValidMatlabIdentifier(ByVal nameValue As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(nameValue) = 0 Then
        IsValidMatlabIdentifier = False
        Exit Function
    End If

    ch = Mid$(nameValue, 1, 1)
    If Not IsAlpha(ch) And ch <> "_" Then
        IsValidMatlabIdentifier = False
        Exit Function
    End If

    For i = 2 To Len(nameValue)
        ch = Mid$(nameValue, i, 1)
        If Not IsAlphaNumeric(ch) And ch <> "_" Then
            IsValidMatlabIdentifier = False
            Exit Function
        End If
    Next i

    IsValidMatlabIdentifier = True
End Function

Private Function IsAlpha(ByVal ch As String) As Boolean
    IsAlpha = ch Like "[A-Za-z]"
End Function

Private Function IsAlphaNumeric(ByVal ch As String) As Boolean
    IsAlphaNumeric = ch Like "[A-Za-z0-9]"
End Function

Private Function EscapeMatlabString(ByVal textValue As String) As String
    EscapeMatlabString = Replace(textValue, "'", "''")
End Function
