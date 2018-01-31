VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "form1"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox password 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox username 
      DataField       =   "&H80000012&"
      Height          =   405
      Left            =   4680
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "                 Admin  Login"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      FillColor       =   &H00E0E0E0&
      Height          =   4815
      Left            =   0
      Top             =   0
      Width           =   11655
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim user As String
Dim pass As String
user = "admin"
pass = "admin"
If (user = username.Text And pass = password.Text) Then
MsgBox ("login successful")
Form2.Show
Else
MsgBox ("login failed")
End If
End Sub



