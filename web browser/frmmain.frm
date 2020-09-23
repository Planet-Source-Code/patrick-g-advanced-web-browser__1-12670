VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Insane Programmers Web Browser"
   ClientHeight    =   6870
   ClientLeft      =   2520
   ClientTop       =   2700
   ClientWidth     =   10110
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   10110
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   960
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":20B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":599E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7612
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9286
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":AEFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   1440
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1080
      Top             =   5520
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   480
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   0
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Urlbox 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   8655
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6615
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11818
            MinWidth        =   864
            Picture         =   "frmmain.frx":CB6E
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   9495
      ExtentX         =   16748
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":D072
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":D336
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":D786
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":DA4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":DF16
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":E41A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":E91E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":EDAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":F1DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":F6DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   926
      ButtonWidth     =   1482
      ButtonHeight    =   873
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Back"
            Key             =   "back"
            Description     =   "back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "oldsites"
                  Text            =   "Sorry was to lazy to finish this"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Forward"
            Key             =   "forward"
            Description     =   "forward"
            Object.ToolTipText     =   "&Forward"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "newsites"
                  Text            =   "Sorry was to lazy to finish this"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Stop"
            Key             =   "stop"
            Description     =   "stop"
            Object.ToolTipText     =   "&Stop"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh"
            Key             =   "refresh"
            Description     =   "refresh"
            Object.ToolTipText     =   "&Refresh"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Home"
            Key             =   "home"
            Description     =   "home"
            Object.ToolTipText     =   "&Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Favorites"
            Key             =   "fav"
            Description     =   "Favorites"
            Object.ToolTipText     =   "&Favorites"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "add"
                  Text            =   "&Add"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "view"
                  Text            =   "&View"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&History"
            Key             =   "history"
            Description     =   "history"
            Object.ToolTipText     =   "&History"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print"
            Key             =   "print"
            Description     =   "print"
            Object.ToolTipText     =   "&Print"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mail"
            Key             =   "mail"
            Description     =   "mail"
            Object.ToolTipText     =   "&Mail"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "send"
                  Text            =   "&Send Mail"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "check"
                  Text            =   "Check mail"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&About"
            Key             =   "about"
            Description     =   "about"
            Object.ToolTipText     =   "&About"
            ImageIndex      =   9
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin MSComctlLib.Toolbar Web1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   5
      Top             =   525
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   794
      ButtonWidth     =   767
      ButtonHeight    =   741
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Shortcuts"
            Description     =   "Shortcuts"
            Object.ToolTipText     =   "Shortcuts"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Calander"
            Description     =   "Calander"
            Object.ToolTipText     =   "Calander"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "source"
            Description     =   "View Source"
            Object.ToolTipText     =   "View Source"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Description     =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "icq"
            Description     =   "ICQ"
            Object.ToolTipText     =   "ICQ"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "msn"
            Description     =   "MSN"
            Object.ToolTipText     =   "MSN Messenger service"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "mp3"
            Description     =   "Mp3 player"
            Object.ToolTipText     =   "Mp3 player"
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   4
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mode 
         Caption         =   "&Work mode"
         Begin VB.Menu online 
            Caption         =   "&Online"
            Checked         =   -1  'True
         End
         Begin VB.Menu offline 
            Caption         =   "&Offline"
         End
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu search 
         Caption         =   "&Search Internet"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu show 
         Caption         =   "&Show"
         Begin VB.Menu statusbar 
            Caption         =   "&Statusbar"
            Checked         =   -1  'True
         End
         Begin VB.Menu toolbar 
            Caption         =   "&Toolbar"
            Checked         =   -1  'True
         End
         Begin VB.Menu urls 
            Caption         =   "Urls"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu stopp 
         Caption         =   "&Stop"
      End
      Begin VB.Menu refreshh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu backk 
         Caption         =   "&Go Back"
      End
      Begin VB.Menu forwardd 
         Caption         =   "&Go Forward"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu source 
         Caption         =   "&Source"
      End
      Begin VB.Menu properties 
         Caption         =   "Properties"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu helpp 
      Caption         =   "&Help"
      Begin VB.Menu help 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu about 
         Caption         =   "&About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frmabout.show
End Sub

Private Sub backk_Click()
    Web.GoBack
End Sub

Private Sub exit_Click()
Dim history As Integer
    Open App.Path & "/history.dll" For Output As #1
        For history = 1 To Urlbox.ListCount - 1
    Print #1, Urlbox.list(history)
        Next history
    Close #1
End
End Sub

Private Sub Form_Load()
Dim Msg, starting
frmabout.show
On Error Resume Next
    iniPath$ = App.Path & "/web.dll"
    starting = GetFromINI("main", "home", iniPath$)
    Web.Navigate (starting)
    StatusBar1.Panels(1).Text = "Ready."
    StatusBar1.Panels(2).Text = "Online."
    On Error GoTo err
    Open App.Path & "/history.dll" For Input As #1
        Do While Not EOF(1)
            Line Input #1, history
                Urlbox.AddItem history
            Loop
    Close #1
err:
    Exit Sub
    Close #1

End Sub
Private Sub Form_Resize()
On Error Resume Next
    Web.Width = frmmain.Width - 100
    Web.Height = frmmain.Height - 2290
    Urlbox.Width = frmmain.Width - 3000
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim history As Integer
    Open App.Path & "/history.dll" For Output As #1
        For history = 1 To Urlbox.ListCount - 1
    Print #1, Urlbox.list(history)
        Next history
    Close #1
End
End Sub

Private Sub forwardd_Click()
    Web.GoForward
End Sub



Private Sub help_Click()
Web.Navigate (App.Path & "/help/help.htm")
End Sub



Private Sub offline_Click()
If online.Checked = True Then
    online.Checked = False
    offline.Checked = True
    Web.offline = True
    StatusBar1.Panels(2).Text = "Offline"
    StatusBar1.Panels(1).Text = "You can now work offline"
End If
End Sub

Private Sub online_Click()
If offline.Checked = True Then
    online.Checked = True
    Web.offline = False
    offline.Checked = False
    StatusBar1.Panels(1).Text = ""
    StatusBar1.Panels(2).Text = "Online"
End If
End Sub

Private Sub open_Click()
On Error Resume Next
    Com.Filter = "All Internet Files (*.hmt,*.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml"
    Com.ShowOpen
If Com.Filename = "" Then
    Exit Sub
Else
    Web.Navigate (Com.Filename)
End If
End Sub

Private Sub properties_Click()
frmoptions.show
End Sub

Private Sub refreshh_Click()
    Web.Refresh
End Sub

Private Sub save_Click()
    Com.Filter = "htm (*.htm) | *.htm"
    Com.ShowSave
If Com.Filename = "" Then
    Exit Sub
Else
    Open Com.Filename For Output As #1
     Print #1, Web.Document
   Close #1
End If
End Sub

Private Sub search_Click()
    iniPath$ = App.Path & "/web.dll"
Dim search
    search = GetFromINI("Search", "url", iniPath$)
    Web.Navigate (search)
    Urlbox.Text = search
    Urlbox.AddItem (search)
End Sub


Private Sub source_Click()
On Error Resume Next
Open App.Path & "/source.tmp" For Output As #1
    Print #1, Inet1.OpenURL(Web.LocationURL)
Close #1
    Shell "C:\windows\notepad.exe " & App.Path & "/source.tmp", vbNormalFocus
    Kill App.Path & "/source.tmp"
End Sub

Private Sub statusbar_Click()
If statusbar.Checked = True Then
    statusbar.Checked = False
    StatusBar1.Visible = False
Else
    statusbar.Checked = True
    StatusBar1.Visible = True
End If
End Sub

Private Sub stopp_Click()
    Web.Stop
End Sub


Private Sub Timer1_Timer()
Unload frmabout
Me.WindowState = 2
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
    frmshortcuts.show
        frmdate.show
    Timer2.Enabled = False
    frmmain.SetFocus
End Sub

Private Sub toolbar_Click()
If toolbar.Checked = True Then
    toolbar.Checked = False
    Toolbar1.Visible = False
Else
    toolbar.Checked = True
    Toolbar1.Visible = True
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim starting
On Error Resume Next
Select Case Button.Key
    Case "back"
     Web.GoBack
     List2.AddItem (Urlbox.Text)
    Case "forward"
     Web.GoForward
    Case "stop"
     Web.Stop
    Case "refresh"
     Web.Refresh
    Case "home"
        iniPath$ = App.Path & "/web.dll"
        starting = GetFromINI("main", "home", iniPath$)
        Web.Navigate (starting)
    Case "fav"
     PopupMenu fav
    Case "print"
     Print Web.Document
    Case "mail"
     PopupMenu mail
    Case "about"
     frmabout.show
    Case "history"
     Loadhistory
End Select

End Sub
Public Sub Loadhistory()
On Error GoTo err
Dim history1 As Integer
        Open App.Path & "/history.htm" For Output As #2
            Print #2, "<html>" & vbCrLf & "<title>History</title>" & vbCrLf & "<font size=15 face=arial color=black>History<br></br><br></br><font size=2 face=arial color=black>" & vbCrLf
                 For history1 = 0 To Urlbox.ListCount
            Print #2, "<a href=http://" + Urlbox.list(history1) + ">" + Urlbox.list(history1) + "<br>"
                Next history1
            Print #2, "</a><font size=1 face=arial color=black><br></br>End of history"
        Close #2
    Web.Navigate (App.Path & "/history.htm")
err:
    Exit Sub
Web.Navigate (App.Path & "/history.htm")
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
    Case "add"
     Dim addfav
    addfav = InputBox("Enter website you wish to add to favorites", "Add", "www.hotmail.com")
        If addfav = "" Then
    Exit Sub
        Else:
        favv.List1.AddItem (addfav)
End If
    Case "view"
     favv.show
    Case "send"
     Dim subject, person
        person = InputBox("Enter email address", "email")
        subject = InputBox("Enter subject for email", "subject")
        Web.Navigate ("mailto:" & person & "?subject=" & subject)
    Case "check"
         Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
End Select
End Sub

Private Sub Urlbox_Change()
Dim b As Integer
Dim d As Integer
On Error Resume Next
    If Len(Urlbox.Text) > 6 Then
d = Len(Urlbox.Text)
    For b = 0 To Len(Urlbox.Text)
        If Left(Urlbox.list(b), d) = Left(Urlbox.Text, d) Then
            Urlbox.Text = Urlbox.list(b)
            Urlbox.SelStart = 6
            Urlbox.SelLength = Len(Urlbox.Text)
    End If
Next b
End If
End Sub

Private Sub Urlbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Web.Navigate (Urlbox.Text)
    Urlbox.AddItem (Urlbox.Text)
End If
End Sub

Private Sub urls_Click()
If urls.Checked = True Then
    urls.Checked = False
    Urlbox.Visible = False
Else
    urls.Checked = True
    Urlbox.Visible = True
End If
End Sub

Private Sub viewfav_Click()
    favv.show
End Sub

Private Sub Web_DocumentComplete(ByVal pDisp As Object, Url As Variant)
    StatusBar1.Panels(1).Text = "Document Finished."
End Sub


Private Sub Web_DownloadBegin()
    StatusBar1.Panels(1).Text = "Opening Page....."
End Sub

Private Sub Web_DownloadComplete()
    StatusBar1.Panels(1).Text = "Download Finished..."
End Sub

Private Sub Web_FileDownload(Cancel As Boolean)
    StatusBar1.Panels(1).Text = "Beginning Download....."
End Sub


Private Sub Web_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
    frmmain.Caption = Web.LocationName
    StatusBar1.Panels(1).Text = Web.LocationURL
    Urlbox.Text = Web.LocationURL
End Sub



Private Sub Web_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    Stat.show
    Stat.ProgressBar1.Max = ProgressMax
    Stat.ProgressBar1.Value = Progress
        If Progress = 0 Then
            Stat.Hide
        Else:
            Stat.show
        End If
End Sub

Private Sub Web_StatusTextChange(ByVal Text As String)
StatusBar1.Panels(1).Text = Text
End Sub

Private Sub Web1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Shortcuts"
     frmshortcuts.show
    Case "Calander"
     frmdate.show
    Case "source"
     On Error Resume Next
Open App.Path & "/source.tmp" For Output As #1
    Print #1, Inet1.OpenURL(Web.LocationURL)
Close #1
    Shell "C:\windows\notepad.exe " & App.Path & "/source.tmp", vbNormalFocus
    Kill App.Path & "/source.tmp"
    Case "Search"
         iniPath$ = App.Path & "/web.dll"
Dim search
    search = GetFromINI("Search", "url", iniPath$)
    Web.Navigate (search)
    Urlbox.Text = search
    Urlbox.AddItem (search)
    Case "icq"
On Error Resume Next
     Shell "C:\program files\icq\icq.exe", vbNormalFocus
    Case "msn"
On Error Resume Next
     Shell "C:\Program Files\Messenger\msmsgs.exe", vbNormalFocus
    Case "mp3"
On Error Resume Next
     Shell App.Path & "/mp3/mp3.exe", vbNormalFocus
     End Select
End Sub
