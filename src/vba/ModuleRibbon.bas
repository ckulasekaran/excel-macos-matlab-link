Attribute VB_Name = "ModuleRibbon"
Option Explicit

Private gRibbon As IRibbonUI

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub OnSnippetAction(ByVal control As IRibbonControl)
    Dim codeText As String
    codeText = BuildSnippet("readmatrix")
    CopyToClipboard codeText
End Sub
