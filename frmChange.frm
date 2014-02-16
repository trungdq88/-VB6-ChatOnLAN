VERSION 5.00
Begin VB.Form frmChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Infomation"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin ChatOnLan.UniTextBox lblErorr 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      Text            =   ""
      Locked          =   -1  'True
      BorderStyle     =   0
      Alignment       =   2
   End
   Begin ChatOnLan.Label Label10 
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   3000
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmChange.frx":0000
      Left            =   2040
      List            =   "frmChange.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin ChatOnLan.Label Label9 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label Label8 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label Label7 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label Label6 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label Label5 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label Label4 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cha As Boolean

Private Sub Combo1_Change()
Cha = True
End Sub

Private Sub Form_Load()
Label10.Caption = "Thông tin cá nhân"
Label1.Caption = ChrW$(&H54) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E)
Label2.Caption = ChrW$(&H4D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75)
Label3.Caption = ChrW$(&H58) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75)
Label4.Caption = ChrW$(&H54) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E)
Label5.Caption = ChrW$(&H54) & ChrW$(&H75) & ChrW$(&H1ED5) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E)
Label6.Caption = ChrW$(&H43) & ChrW$(&HE2) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1ECF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HED) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74)
Label7.Caption = ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H1EA3) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EDD) & ChrW$(&H69)
Label8.Caption = "  " & ChrW$(&H58) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H6E)
Label9.Caption = "      " & ChrW$(&H54) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&HE1) & ChrW$(&H74)

'**********************************888
GetInfo
End Sub

Private Sub Label8_Click()
If Text2 = Text3 Then

If Cha = False Then
lblErorr.Text = ChrW$(&H4B) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H1EF1) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1ED5) & ChrW$(&H69) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H1EDF) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA1) & ChrW$(&H69)
Else
ThayDoi
lblErorr.Text = ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1ED5) & ChrW$(&H69) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H1EDF) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA1) & ChrW$(&H69)
End If

Else
lblErorr.Text = ChrW$(&H4D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HFA) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H21)
End If
End Sub

Private Sub Label9_Click()
Unload Me
End Sub

Private Sub Text2_Change()
Cha = True
End Sub

Private Sub Text3_Change()
Cha = True
End Sub

Private Sub Text4_Change()
Cha = True
End Sub

Private Sub Text5_Change()
Cha = True
End Sub

Private Sub Text6_Change()
Cha = True
End Sub
Private Sub ThayDoi()
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)

Set rs = db.OpenRecordset("Select * from Chat where USER = '" & frmMem.Label6.Caption & "'")

rs.Edit
rs("PASS") = Text2
rs("NAME") = Text4
rs("OLD") = Text5
rs("ANQUES") = Text6
rs("QUES") = Combo1.Text
rs.Update

End Sub

Private Sub GetInfo()
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)

Set rs = db.OpenRecordset("Select * from Chat where USER = '" & frmMem.Label6.Caption & "'")

Text1 = rs("USER")
Text2 = rs("PASS")
Text3 = rs("PASS")
Text4 = rs("NAME")
Text5 = rs("OLD")
Text6 = rs("ANQUES")
Combo1.Text = rs("QUES")
Cha = False
End Sub
