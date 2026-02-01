Attribute VB_Name = "ModulePath"
Option Explicit

Public Function GetWorkbookPath() As String
    If ActiveWorkbook Is Nothing Then
        GetWorkbookPath = ""
        Exit Function
    End If

    If ActiveWorkbook.Path = "" Then
        GetWorkbookPath = ""
    Else
        GetWorkbookPath = ActiveWorkbook.FullName
    End If
End Function
