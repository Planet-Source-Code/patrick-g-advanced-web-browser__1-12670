VERSION 5.00
Begin VB.Form FrmShortcuts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shortcut"
   ClientHeight    =   750
   ClientLeft      =   12180
   ClientTop       =   2550
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image9 
      Height          =   285
      Left            =   1440
      Picture         =   "shortcuts.frx":0000
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   1080
      Picture         =   "shortcuts.frx":04B6
      Top             =   360
      Width           =   285
   End
   Begin VB.Image Image7 
      Height          =   195
      Left            =   720
      Picture         =   "shortcuts.frx":096C
      Top             =   480
      Width           =   225
   End
   Begin VB.Image Image6 
      Height          =   195
      Left            =   360
      Picture         =   "shortcuts.frx":0C1E
      Top             =   480
      Width           =   225
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   0
      Picture         =   "shortcuts.frx":0ED0
      Top             =   360
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   1080
      Picture         =   "shortcuts.frx":13C2
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   720
      Picture         =   "shortcuts.frx":18B4
      Top             =   0
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   360
      Picture         =   "shortcuts.frx":1DA6
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "shortcuts.frx":2260
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "frmshortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Ontop Me
    Me.Top = frmmain.Height / 3
End Sub


Private Sub Image1_Click()
    favv.show
End Sub

Private Sub Image2_Click()
        iniPath$ = App.Path & "/web.dll"
        starting = GetFromINI("main", "home", iniPath$)
        frmmain.Web.Navigate (starting)
End Sub

Private Sub Image3_Click()
    frmoptions.show
End Sub

Private Sub Image4_Click()
Dim subject, person
    person = InputBox("Enter email address", "email")
    subject = InputBox("Enter subject for email", "subject")
    frmmain.Web.Navigate ("mailto:" & person & "?subject=" & subject)
End Sub

Private Sub Image5_Click()
    iniPath$ = App.Path & "/web.dll"
Dim search
    search = GetFromINI("Search", "url", iniPath$)
    frmmain.Web.Navigate (search)
    frmmain.Urlbox.Text = search
    frmmain.Urlbox.AddItem (search)
End Sub

Private Sub Image6_Click()
On Error Resume Next
    frmmain.Web.GoBack
End Sub

Private Sub Image7_Click()
On Error Resume Next
    frmmain.Web.GoForward
End Sub

Private Sub Image8_Click()
On Error Resume Next
    frmmain.Web.Stop
End Sub

Private Sub Image9_Click()
On Error GoTo err
Dim history1 As Integer
        Open App.Path & "/history.htm" For Output As #2
            Print #2, "<html>" & vbCrLf & "<title>History</title>" & vbCrLf & "<font size=15 face=arial color=black>History<br></br><br></br><font size=2 face=arial color=black>" & vbCrLf
                 For history1 = 0 To frmmain.Urlbox.ListCount
            Print #2, "<a href=http://" + frmmain.Urlbox.list(history1) + ">" + frmmain.Urlbox.list(history1) + "<br>"
                Next history1
            Print #2, "</a><font size=1 face=arial color=black><br></br>End of history"
        Close #2
    frmmain.Web.Navigate (App.Path & "/history.htm")
err:
    Exit Sub
frmmain.Web.Navigate (App.Path & "/history.htm")
End Sub
