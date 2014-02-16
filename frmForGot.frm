VERSION 5.00
Begin VB.Form frmForGot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fotgoten Password"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
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
   ScaleHeight     =   2955
   ScaleWidth      =   4785
   StartUpPosition =   1  'CenterOwner
   Begin ChatOnLan.UniTextBox lblErorr 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
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
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmForGot.frx":0000
      Left            =   1800
      List            =   "frmForGot.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   2895
   End
   Begin ChatOnLan.Label Label6 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BackColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label Label5 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BackColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label Label4 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
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
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
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
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
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
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmForGot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace


Private Sub Form_Load()
Label1.Caption = ChrW$(&H51) & ChrW$(&H75) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H3F)
Label2.Caption = ChrW$(&H54) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E)
Label3.Caption = ChrW$(&H43) & ChrW$(&HE2) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1ECF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HED) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74)
Label4.Caption = ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H1EA3) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EDD) & ChrW$(&H69)
Label5.Caption = "   " & ChrW$(&H54) & ChrW$(&HEC) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75)
Label6.Caption = "          Thoát"
Combo1.ListIndex = 0

End Sub

Private Sub Label5_Click()
If Text1.Text = "" Then
lblErorr.Text = ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E)
ElseIf Text2.Text = "" Then
lblErorr.Text = ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H1EA3) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EDD) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HE2) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1ECF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HED) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H21)
Else
lblErorr.Text = ""
Search
End If
End Sub

Private Sub Label6_Click()
Unload Me
End Sub
Private Sub Search()
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
On Error GoTo loi
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & Text1.Text & "'")
If rs("ANQUES") = Text2.Text And rs("QUES") = Combo1.Text Then
lblErorr.Text = ChrW$(&H4D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&HE0) & ChrW$(&H20) & ChrW$(&H3A) & " : " & rs("PASS")
End If

loi:
If lblErorr.Text = "" Then
lblErorr.Text = ChrW$(&H53) & ChrW$(&H61) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H21)
End If
End Sub
