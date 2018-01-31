VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form5"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   14565
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtid 
      DataField       =   "id"
      DataSource      =   "search"
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton gobtn 
      Caption         =   "Go"
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
      Left            =   13680
      TabIndex        =   24
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox idsearch 
      Height          =   495
      Left            =   11400
      TabIndex        =   23
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9120
      TabIndex        =   22
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton updatebtn 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   21
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton delbtn 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton clearbtn 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton addbtn 
      Caption         =   "Add "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   18
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton prevtbtn 
      Caption         =   "Previous"
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
      Left            =   7080
      TabIndex        =   17
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton nxtbtn 
      Caption         =   "Next"
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
      Left            =   7080
      TabIndex        =   16
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton lastbtn 
      Caption         =   "Last"
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
      Left            =   7080
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton firstbtn 
      Caption         =   "First"
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
      Left            =   7080
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtemail 
      DataField       =   "email"
      DataSource      =   "search"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox txtphone 
      DataField       =   "phoneno"
      DataSource      =   "search"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtname 
      DataField       =   "name"
      DataSource      =   "search"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtage 
      DataField       =   "age"
      DataSource      =   "search"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtdesignation 
      DataField       =   "designation"
      DataSource      =   "search"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "address"
      DataSource      =   "search"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   4320
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc search 
      Height          =   495
      Left            =   10560
      Top             =   6840
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from newentry"
      Caption         =   "search"
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
   Begin VB.Label Label10 
      Caption         =   "search by id and name"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   25
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "address"
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
      Left            =   960
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "e-mail id"
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
      Left            =   960
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "phone no"
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
      Left            =   960
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "age"
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
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "designation"
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
      Left            =   960
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "name"
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
      Left            =   960
      TabIndex        =   2
      Top             =   2160
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
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "          search ENTRY"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addbtn_Click()
search.Recordset.AddNew
End Sub

Private Sub clearbtn_Click()
txtid.Text = ""
txtname.Text = ""
txtage.Text = ""
txtdesignation.Text = ""
txtaddress.Text = ""
txtemail.Text = ""
txtphone.Text = ""
End Sub


Private Sub Command1_Click()
Form2.Show
Unload Me
End Sub

Private Sub delbtn_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "Delete Record Confirmation")
If confirmation = vbYes Then
search.Recordset.Delete
MsgBox "Record has been deleted successfully", vbInformation, "message"
Else
MsgBox "Record Not Delete!!!", vbInformation, "Message"
End If

End Sub

Private Sub firstbtn_Click()
search.Recordset.MoveFirst

End Sub

Private Sub gobtn_Click()
search.RecordSource = "select * from newentry where id ='" + idsearch.Text + "'or Name='" + idsearch.Text + "'"
search.Refresh
If search.Recordset.EOF Then
MsgBox "Record Not Found,Please Try any other Id", vbInformation
Else
search.Caption = search.RecordSource
End If
End Sub

Private Sub lastbtn_Click()
search.Recordset.MoveLast

End Sub

Private Sub nxtbtn_Click()
search.Recordset.MoveNext

End Sub

Private Sub prevtbtn_Click()
search.Recordset.MovePrevious

End Sub



Private Sub updatebtn_Click()
search.Recordset.AddNew
search.Recordset.Fields("id") = txtid.Text
search.Recordset.Fields("name") = txtname.Text
search.Recordset.Fields("age") = txtage.Text
search.Recordset.Fields("designation") = txtdesignation.Text
search.Recordset.Fields("address") = txtaddress.Text
search.Recordset.Fields("email") = txtemail.Text
search.Recordset.Fields("phoneno") = txtphone.Text
search.Recordset.Update
MsgBox "Data is Successfully Updated ", vbInformation, "Message"
End Sub
