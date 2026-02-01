Attribute VB_Name = "ModuleClipboard"
Option Explicit

Public Sub CopyToClipboard(ByVal textValue As String)
    If textValue = "" Then
        Exit Sub
    End If

    If Not SetClipboardMac(textValue) Then
        ShowClipboardFallback textValue
    End If
End Sub

Private Function SetClipboardMac(ByVal textValue As String) As Boolean
    Dim scriptText As String
    scriptText = "set the clipboard to " & AppleScriptString(textValue)
    On Error Resume Next
    MacScript scriptText
    If Err.Number <> 0 Then
        Err.Clear
        SetClipboardMac = False
    Else
        SetClipboardMac = True
    End If
    On Error GoTo 0
End Sub

Private Sub ShowClipboardFallback(ByVal textValue As String)
    Dim fallbackForm As UserFormClipboardFallback
    Set fallbackForm = New UserFormClipboardFallback
    fallbackForm.SetContent textValue, ClipboardHelpUrl
    fallbackForm.Show vbModal
End Sub

Private Function AppleScriptString(ByVal textValue As String) As String
    Dim escapedText As String
    escapedText = Replace(textValue, "\", "\\")
    escapedText = Replace(escapedText, """", "\""")
    AppleScriptString = """" & escapedText & """"
End Function
