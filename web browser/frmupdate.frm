VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmupdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   9000
      Left            =   3120
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      Max             =   20
      TickStyle       =   3
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   120
      Picture         =   "frmupdate.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   240
      Width           =   335
   End
   Begin VB.Label Label1 
      Caption         =   "Checking for updates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Slider1.Value < 20 Then
    Slider1.Value = Slider1.Value + 1
   Else:
    Timer1.Enabled = False
    Timer2.Enabled = True
End If

End Sub


Private Sub Timer2_Timer()
If Slider1.Value > 0 Then
    Slider1.Value = Slider1.Value - 1
    Else:
    Timer2.Enabled = False
    Timer1.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
MsgBox "No update found" & vbCrLf & "Latest version: 5.7b", vbInformation, "No update found"
Unload Me
End Sub
