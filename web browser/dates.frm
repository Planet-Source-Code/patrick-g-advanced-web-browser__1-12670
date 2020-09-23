VERSION 5.00
Begin VB.Form dates 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Dates"
   ClientHeight    =   2775
   ClientLeft      =   9720
   ClientTop       =   8220
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Rem"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "dates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
 Dim a
        List1.RemoveItem (b)
End Sub

Private Sub Form_Load()
Ontop Me
On Error Resume Next
    Call ReadList(List1, "C:\windows\dates.tmp", True)

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If List1.ListCount = 0 Then
        Kill "C:\windows\dates.tmp"
    Else
        Call WriteList(List1, "C:\windows\dates.tmp")
    End If
End Sub

Private Sub List1_DblClick()
On Error Resume Next
    iniPath$ = App.Path & "/web.dll"
 Dim GetDate As String
 Dim b
    b = List1.ListIndex
    GetDate = GetFromINI("dates", List1.list(b), iniPath$)
        MsgBox List1.list(b) & vbCrLf & "--------------------------" & vbCrLf & GetDate, vbInformation, "Set dates"
End Sub
