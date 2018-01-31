VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form splash 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   661
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   2
      OLEDropMode     =   1
      Max             =   105
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   10800
      Top             =   4080
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                industry management system"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label lblstat 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblstatus 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BorderColor     =   &H00000000&
      Height          =   6015
      Left            =   -360
      Top             =   -120
      Width           =   12135
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

End Sub


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
lblstatus.Caption = "Loading..Plz Wait.."
lblstat.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Enabled = False
Unload Me
Form1.Show
End If

End Sub
