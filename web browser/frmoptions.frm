VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmoptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   Icon            =   "frmoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Delete History"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Check for update"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Save"
      Height          =   255
      Left            =   6960
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   2535
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   7560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Home Page"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Use Blank"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Use Custom"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Okay"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "HTML:"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   $"frmoptions.frx":0442
      Height          =   855
      Left            =   2760
      TabIndex        =   13
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Start With"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Defult Search Site"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    Com.Filter = "All Internet Files (*.hmt,*.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml"
    Com.ShowOpen
If Com.Filename = "" Then
    Exit Sub
Else
    Text1.Text = Com.Filename
End If
End Sub


Private Sub Command2_Click()
iniPath$ = App.Path & "/web.dll"
    entry$ = Text1.Text
     r% = WritePrivateProfileString("main", "home", entry$, iniPath$)
    entry$ = Text2.Text
     r% = WritePrivateProfileString("Search", "url", entry$, iniPath$)
    entry$ = Option1.Value
     r% = WritePrivateProfileString("main", "option1", entry$, iniPath$)
    entry$ = Option2.Value
     r% = WritePrivateProfileString("main", "option2", entry$, iniPath$)
    entry$ = Option3.Value
     r% = WritePrivateProfileString("main", "option3", entry$, iniPath$)
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
    Com.Filter = "htm (*.htm) | *.htm"
    Com.ShowSave
If Com.Filename = "" Then
    Exit Sub
Else
    Open Com.Filename & ".htm" For Output As #1
     Print #1, Text3.Text
   Close #1
End If
End Sub

Private Sub Command5_Click()
frmupdate.Show
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim msg
    msg = MsgBox("Are you sure you want to delete all history?", vbYesNo Or vbQuestion, "Delete?")
        If msg = vbYes Then
            Kill App.Path & "/history.dll"
            frmmain.Urlbox.Clear
        Else
            Exit Sub
        End If
End Sub

Private Sub Form_Load()
On Error Resume Next
iniPath$ = App.Path & "/web.dll"
Dim a, b, c, d, e As String

    a = GetFromINI("main", "home", iniPath$)
     Text1.Text = a
    b = GetFromINI("Search", "url", iniPath$)
     Text2.Text = b
    c = GetFromINI("main", "option1", iniPath$)
     Option1.Value = c
    d = GetFromINI("main", "option2", iniPath$)
     Option2.Value = d
    e = GetFromINI("main", "option3", iniPath$)
     Option3.Value = e
    List1.AddItem ("Update")
    List1.AddItem ("Web Browser")
    List1.AddItem ("Search Site")
    List1.AddItem ("HTML")
    Text3.Text = "<Html>" & vbCrLf & "<Title>Insane Programmers WeB Browser</Title>" & vbCrLf & "<Body Bgcolor=000000>" & vbCrLf & "<center>" & vbCrLf & "<font color=white face=arial size=7>" & vbCrLf & "Click <a href=http://www.insaneprogammer.f2s.com>Here</a> to visit insaneprogrammers" & vbCrLf & "</center>" & vbCrLf & "</body>" & vbCrLf & "</html>"
End Sub

Private Sub List1_Click()
Dim a
    a = List1.ListIndex
        If a = 0 Then
            Option1.Visible = False
            Option2.Visible = False
            Option3.Visible = False
            Text1.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            Label4.Visible = True
            Command1.Visible = False
            Label2.Visible = False
            Text2.Visible = False
            Text3.Visible = False
            Label5.Visible = False
            Command4.Visible = False
            Command5.Visible = True
            Command6.Visible = False
        ElseIf a = 1 Then
            Command6.Visible = True
            Option1.Visible = True
            Option2.Visible = True
            Option3.Visible = True
            Text1.Visible = True
            Label1.Visible = True
            Label2.Visible = True
            Label4.Visible = False
            Command1.Visible = True
            Label2.Visible = False
            Text2.Visible = False
            Text3.Visible = False
            Label5.Visible = False
            Command4.Visible = False
            Command5.Visible = False
        ElseIf a = 2 Then
            Command6.Visible = False
            Option1.Visible = False
            Option2.Visible = False
            Option3.Visible = False
            Text1.Visible = False
            Label4.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            Command1.Visible = False
            Text3.Visible = False
            Label5.Visible = False
            Command4.Visible = False
            Label2.Visible = True
            Command5.Visible = False
            Text2.Visible = True
        ElseIf a = 3 Then
            Command6.Visible = False
            Option1.Visible = False
            Option2.Visible = False
            Option3.Visible = False
            Label4.Visible = False
            Text1.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            Command5.Visible = False
            Command1.Visible = False
            Label2.Visible = False
            Text2.Visible = False
            Text3.Visible = True
            Label5.Visible = True
            Command4.Visible = True
End If

End Sub

Private Sub Option1_Click()
    Text1.Text = "http://www.insaneprogammer.f2s.com"
    Text1.Enabled = False
End Sub

Private Sub Option2_Click()
    Text1.Text = "about:blank"
    Text1.Enabled = False
End Sub

Private Sub Option3_Click()
    Text1.Enabled = True
    Text1.Text = "http://www.hotmail.com"
End Sub
