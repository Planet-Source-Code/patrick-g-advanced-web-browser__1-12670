VERSION 5.00
Begin VB.Form favv 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "favv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    Call WriteList(List1, "C:\windows\fav.tmp")
    Unload Me
End Sub

Private Sub Command2_Click()
Dim aa
    aa = List1.ListIndex
        List1.RemoveItem (aa)
        
End Sub

Private Sub Form_Load()
Ontop Me
On Error Resume Next
    Call ReadList(List1, "C:\windows\fav.tmp", True)
    Call WriteList(List1, "C:\windows\fav.tmp")
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call WriteList(List1, "C:\windows\fav.tmp")
End Sub

Private Sub List1_DblClick()
Dim bb
    bb = List1.ListIndex
        frmmain.Urlbox.Text = List1.list(bb)
        frmmain.Web.Navigate List1.list(bb)
        Unload Me
End Sub
