VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin ChatOnLan.UniTextBox lblErorr 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   1920
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin ChatOnLan.Label label1 
      Height          =   375
      Index           =   7
      Left            =   5040
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label label1 
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label label1 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label label1 
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin ChatOnLan.Label label1 
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   5535
   End
   Begin ChatOnLan.Label label1 
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label label1 
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label label1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   873
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace

Dim OK As Boolean


Function FileExists(sFile As String) As Boolean
 On Error Resume Next
 FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function
Private Sub Form_Load()
If FileExists("Z:\CSDL.MDB") = False Then
UniMsgBox "Không tìm th" & ChrW$(&H1EA5) & "y c" & ChrW$(&H1A1) & " s" & ChrW$(&H1EDF) & " d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u ! Xin th" & ChrW$(&H1EED) & " l" & ChrW$(&H1EA1) & "i sau vài giây !"
End
Else


Label1(0).Caption = ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H1EE3) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&HE0) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&HF2) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H20)
Label1(1).Caption = ChrW$(&H54) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E)
Label1(2).Caption = ChrW$(&H4D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75)
Label1(3).Caption = "       Thoát"
Label1(4).Caption = "   " & ChrW$(&H110) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H70)
Label1(5).Caption = ChrW$(&H4E) & ChrW$(&H1EBF) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H3A)
Label1(6).Caption = "    " & ChrW$(&H110) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&HFD)
Label1(7).Caption = "     " & ChrW$(&H51) & ChrW$(&H75) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75) & " ?"
'COnnect
Set db = OpenDatabase("Z:\CSDL.MDB", False, False, ";PWD=881817258")

OK = False
End If
End Sub

Private Sub Label1_Click(Index As Integer)
If Index = 3 Then End
If Index = 6 Then frmDK.Show
If Index = 7 Then frmForGot.Show
If Index = 4 Then
    If Text1.Text = "" Or Text2.Text = "" Then
        lblErorr.Text = ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EB7) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75)
    Else
    Login
    End If
End If
End Sub
Private Sub Login()
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
If rs("USER") = Text1.Text And rs("PASS") = Text2.Text Then
lblErorr.Text = ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H1ED1) & ChrW$(&H69) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H75) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&HF2) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EDD) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H69) & ChrW$(&HE2) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H2E) & ChrW$(&H2E) & ChrW$(&H2E)
    If rs("ADMIN") = "yes" Then
        frmAdminFirst.Show
        Else
            If rs("ONLINE") = "no" Then
                Online
                frmMem.Show
                frmMem.Label4.Caption = ChrW$(&H58) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&HE0) & ChrW$(&H6F) & ChrW$(&H20) & rs("NAME")
                frmMem.Label6.Caption = rs("USER")
            Else
                If UniMsgBox(ChrW$(&H54) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H1EE3) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H1EED) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EE5) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H75) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&HF2) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1EED) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H61) & ChrW$(&H75) & ChrW$(&H21), vbRetryCancel, "Thông báo") = vbRetry Then
                    Login
                Else
                    End
                End If
            End If
    End If
Unload Me
End If

loi:
If lblErorr.Text = "" Then
lblErorr.Text = ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H65) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HE0) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H1EB7) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAD) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EA9) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H21)
End If
End Sub

Private Sub Text2_Click()
lblErorr.Text = ""
End Sub

Private Sub Text1_Click()
lblErorr.Text = ""
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Label1_Click (4)
End If
End Sub
Private Sub Online()


Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Room", dbOpenTable)
rs.AddNew
rs("ListUser") = Text1.Text
rs.Update

Set rs = db.OpenRecordset("Chat", dbOpenTable)

Set rs = db.OpenRecordset("Select * from Chat where User = '" & Text1.Text & "'")
rs.Edit
rs("ONLINE") = "yes"
rs("MSG") = "no"
rs.Update

End Sub
