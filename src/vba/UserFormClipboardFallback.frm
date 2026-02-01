VERSION 5.00
Begin VB.UserForm UserFormClipboardFallback
   Caption         =   "Clipboard Access Blocked"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblMessage
      Caption         =   "Clipboard access was blocked. Copy the snippet below."
      Height          =   600
      Left            =   120
      Top             =   120
      Width           =   6960
      WordWrap        =   -1  'True
   End
   Begin VB.TextBox txtSnippet
      EnterKeyBehavior=   -1  'True
      Height          =   3000
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      Top             =   900
      Width           =   6960
   End
   Begin VB.CommandButton btnOpenGuide
      Caption         =   "Open Guide"
      Height          =   360
      Left            =   120
      Top             =   4020
      Width           =   1320
   End
   Begin VB.CommandButton btnClose
      Caption         =   "Close"
      Height          =   360
      Left            =   5760
      Top             =   4020
      Width           =   1320
   End
End
Attribute VB_Name = "UserFormClipboardFallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGuideUrl As String

Public Sub SetContent(ByVal snippetText As String, ByVal guideUrl As String)
    mGuideUrl = guideUrl
    txtSnippet.Text = snippetText
    lblMessage.Caption = "Excel could not access the clipboard (macOS restriction)." & vbCrLf & _
        "Copy the snippet below or open the guide to allow clipboard access."
    btnOpenGuide.Enabled = (mGuideUrl <> "")
End Sub

Private Sub btnOpenGuide_Click()
    If mGuideUrl <> "" Then
        ThisWorkbook.FollowHyperlink mGuideUrl
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
