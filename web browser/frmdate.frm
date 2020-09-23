VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date"
   ClientHeight    =   2385
   ClientLeft      =   11430
   ClientTop       =   6330
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComCtl2.MonthView Date 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24510465
      CurrentDate     =   36844
   End
End
Attribute VB_Name = "frmdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Date_DateClick(ByVal DateClicked As Date)
Text1.Text = DateClicked
Dim Msg, SetDate, Info, OtherDate
Info = DateClicked
OtherDate = Date
    iniPath$ = App.Path & "/web.dll"
        Msg = MsgBox("Would you like to set something for " & DateClicked & " ?", vbYesNo Or vbQuestion, "Set date?")
            If Msg = vbYes Then
                SetDate = InputBox("Enter what you would like to set for this date", "Info")
        entry$ = SetDate
            r% = WritePrivateProfileString("dates", Text1.Text, entry$, iniPath$)
            dates.show
            dates.List1.AddItem (DateClicked)
            Call WriteList(dates.List1, "C:\windows\dates.tmp")
            dates.Hide
            End If
End Sub

Private Sub Date_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        dates.show
    End If
End Sub

Private Sub Form_Load()
    Me.Top = frmmain.Height / 2
    Ontop Me
End Sub
