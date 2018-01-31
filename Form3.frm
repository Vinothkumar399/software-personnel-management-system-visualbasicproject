VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   13380
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc ado 
      Height          =   330
      Left            =   6840
      Top             =   5760
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Casetools\newentry.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Casetools\newentry.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "newentry"
      Caption         =   "Ado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
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
      Left            =   7440
      TabIndex        =   15
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Entry"
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
      Left            =   7440
      TabIndex        =   14
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtphoneno 
      DataField       =   "phoneno"
      DataSource      =   "Ado"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox txtname 
      DataField       =   "name"
      DataSource      =   "Ado"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtage 
      DataField       =   "age"
      DataSource      =   "Ado"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtdesignation 
      DataField       =   "designation"
      DataSource      =   "Ado"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "address"
      DataSource      =   "Ado"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtemailid 
      DataField       =   "email"
      DataSource      =   "Ado"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtid 
      DataField       =   "id"
      DataSource      =   "ado"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "   NEW employee  registration"
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
      Left            =   3960
      TabIndex        =   16
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label7 
      Caption         =   "Phone no"
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
      Left            =   480
      TabIndex        =   6
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "E-Mail Id"
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
      Left            =   480
      TabIndex        =   5
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Address"
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
      Left            =   480
      TabIndex        =   4
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Designation"
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
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Age"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Id"
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
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ado.Recordset.Fields("id") = txtid.Text
ado.Recordset.Fields("name") = txtname.Text
ado.Recordset.Fields("age") = txtage.Text
ado.Recordset.Fields("designation") = txtdesignation.Text
ado.Recordset.Fields("address") = txtaddress.Text
ado.Recordset.Fields("email") = txtemailid.Text
ado.Recordset.Fields("phoneno") = txtphoneno.Text
ado.Recordset.Update
MsgBox ("Added successfully"), vbInformation
End Sub

Private Sub Command2_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
ado.Recordset.AddNew

End Sub

