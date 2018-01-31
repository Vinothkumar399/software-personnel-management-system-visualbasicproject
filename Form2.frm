VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   12870
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "NewEntry"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000080&
      Caption         =   "Exit"
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
      Left            =   9720
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Search"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Entry"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Industry management system"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form4.Show
Unload Me
End Sub

Private Sub Command3_Click()
Form5.Show
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
