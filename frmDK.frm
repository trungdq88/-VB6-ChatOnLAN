VERSION 5.00
Begin VB.Form frmDK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
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
   ScaleHeight     =   4815
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   5400
      TabIndex        =   18
      Top             =   600
      Width           =   1935
   End
   Begin ChatOnLan.UniTextBox lblErorr 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   3840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
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
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label Label9 
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label Label8 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
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
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   3480
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmDK.frx":0000
      Left            =   2040
      List            =   "frmDK.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Width           =   2655
   End
   Begin ChatOnLan.Label Label7 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   1695
      _ExtentX        =   2990
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
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Tai Khoan"
   End
End
Attribute VB_Name = "frmDK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long
Dim OK As Boolean


Private Sub Form_Load()
Label1.Caption = ChrW$(&H54) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H4B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E)
Label2.Caption = ChrW$(&H4D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H4B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75)
Label3.Caption = ChrW$(&H4E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75)
Label4.Caption = ChrW$(&H54) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H3A)
Label5.Caption = ChrW$(&H54) & ChrW$(&H75) & ChrW$(&H1ED5) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H3A)
Label6.Caption = ChrW$(&H43) & ChrW$(&HE2) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1ECF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HED) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&HF2) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&HFA) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H71) & ChrW$(&H75) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75)
Label7.Caption = ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H1EA3) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EDD) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H3A)
Label8.Caption = ChrW$(&H42) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&HFD)
Label9.Caption = "   " & ChrW$(&H110) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&HFD)
Label10.Caption = "      Thoát"
Combo1.ListIndex = 0
OK = True
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)

Max = rs.RecordCount

If rs.RecordCount = 0 Then
Exit Sub
Else
rs.MoveFirst

List1.Clear

For i = 1 To Max
    List1.AddItem rs!User
    rs.MoveNext
Next i
List1.ListIndex = 0
End If
End Sub

Private Sub Label10_Click()
Unload Me
End Sub

Private Sub Label9_Click()
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)

Max = rs.RecordCount

If rs.RecordCount = 0 Then
Exit Sub
Else
rs.MoveFirst

List1.Clear

For i = 1 To Max
    List1.AddItem rs!User
    rs.MoveNext
Next i
List1.ListIndex = 0
End If

Check

If OK = True Then
If Text1.Text = "" Then
lblErorr.Text = ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H4B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E) & " !"
ElseIf Text2.Text <> Text3.Text Or Text2.Text = "" Then
lblErorr.Text = ChrW$(&H4D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HFA) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H21)
ElseIf Text6.Text = "" Then
lblErorr.Text = ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H1EA3) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EDD) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HE2) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1ECF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HED) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H21)
ElseIf Len(Text1.Text) < 5 Then
lblErorr.Text = ChrW$(&H54) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H71) & ChrW$(&H75) & ChrW$(&HE1) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H1EAF) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H21)
ElseIf Len(Text2.Text) < 7 Then
lblErorr.Text = ChrW$(&H4D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H71) & ChrW$(&H75) & ChrW$(&HE1) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H1EAF) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H21)
Else

lblErorr.Text = ""
Register
End If
End If



End Sub
Private Sub Register()

Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)

rs.AddNew
rs("USER") = Text1.Text
rs("PASS") = Text2.Text
rs("NAME") = Text4.Text
rs("OLD") = Text5.Text
rs("QUES") = Combo1.Text
rs("ANQUES") = Text6.Text
rs("ADMIN") = "no"
rs("ONLINE") = "no"
rs("MSG") = "no"
rs.Update
rs.MoveFirst

lblErorr.Text = ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&HFD) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&HE0) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H21) & ChrW$(&H20) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H21)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.ListIndex = 0

End Sub

Private Sub Check()

Dim i
For i = 0 To List1.ListCount - 1
If Text1.Text = List1.List(i) Then
OK = False
lblErorr.Text = ChrW$(&H54) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H1EE3) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H1EED) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EE5) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21)
End If
Next i
End Sub
