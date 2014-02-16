VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAdmin.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   7800
      Top             =   6600
   End
   Begin ChatOnLan.UniTextBox Text 
      Height          =   255
      Left            =   4440
      TabIndex        =   38
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      BackColor       =   -2147483641
      Text            =   ""
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000007&
      Caption         =   "Edit"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   37
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H008080FF&
      Height          =   285
      Index           =   6
      Left            =   5280
      TabIndex        =   36
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H008080FF&
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   35
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H008080FF&
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   34
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H008080FF&
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   33
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H008080FF&
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   32
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H008080FF&
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   31
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H008080FF&
      Height          =   285
      Index           =   0
      Left            =   5280
      TabIndex        =   30
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Height          =   255
      Left            =   6720
      TabIndex        =   27
      Top             =   5160
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Height          =   255
      Left            =   6720
      TabIndex        =   26
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   25
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   24
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kick"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   19
      Top             =   4080
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4020
      ItemData        =   "frmAdmin.frx":2D85C
      Left            =   360
      List            =   "frmAdmin.frx":2D85E
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Timer Timer2 
      Left            =   4680
      Top             =   6840
   End
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   6720
   End
   Begin VB.Label txtChat 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   6135
      Left            =   8040
      TabIndex        =   39
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Send 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   4560
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SendKeys"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   23
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Shell"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   22
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Msgbox"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   21
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   18
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   17
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   16
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   15
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   14
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   12
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   10
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Messenger      :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin mode     :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer         :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Question       :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Online         :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Age            :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name           :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password       :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name      :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Dim m_lAlpha

Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Dim KiemTra
Dim Li As String
Dim Lo As String




Private Sub Check1_Click()
If Check1.Value = 1 Then
OpenEdit
Else
CloseEdit
End If
End Sub

Private Sub Command2_Click()

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & List1.List(List1.ListIndex) & "'")
rs.Delete
Dim xua
xua = List1.ListIndex
Display
If xua <> 0 Then
List1.ListIndex = xua - 1
End If
End Sub

Private Sub Command3_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & List1.List(List1.ListIndex) & "'")
rs.Edit
rs("MSG") = "kck"
rs.Update

End Sub

Private Sub Command5_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & List1.List(List1.ListIndex) & "'")
rs.Edit
rs("MSG") = Text.Text & "msg"
rs.Update
If Not rs("MSG") = vbNullString Then Label2(8) = rs("MSG") Else Label2(8) = "----"

End Sub

Private Sub Command6_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & List1.List(List1.ListIndex) & "'")
rs.Edit
rs("MSG") = Text1(1).Text & "sll"
rs.Update
If Not rs("MSG") = vbNullString Then Label2(8) = rs("MSG") Else Label2(8) = "----"

End Sub

Private Sub Command7_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & List1.List(List1.ListIndex) & "'")
rs.Edit
rs("MSG") = Text1(2).Text & "sdk"
rs.Update
If Not rs("MSG") = vbNullString Then Label2(8) = rs("MSG") Else Label2(8) = "----"

End Sub

Private Sub Form_Load()
KiemTra = 1
'**************************************************
'Ho tro mo dan Form *******************************
    Dim lStyle As Long
    lStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, lStyle
    SetLayeredWindowAttributes Me.hWnd, 0, 0, LWA_ALPHA
    Timer1.Interval = 25
    Timer2.Interval = 25
    Timer2.Enabled = False
    Timer1.Enabled = True
'**************************************************
'**************************************************
Display
For hi = 0 To 6
Text2(hi).Visible = False
Next hi

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Cancel = True
        Timer2.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & List1.List(List1.ListIndex) & "'")
Label2(0) = rs("USER")
Label2(1) = rs("PASS")
If Not rs("NAME") = vbNullString Then Label2(2) = rs("NAME") Else Label2(2) = "----"
If Not rs("OLD") = vbNullString Then Label2(3) = rs("OLD") Else Label2(3) = "----"
If Not rs("ONLINE") = vbNullString Then Label2(4) = rs("ONLINE") Else Label2(4) = "----"
If Not rs("ONLINE") = vbNullString Then Label2(5) = rs("QUES") Else Label2(5) = "----"
If Not rs("ANQUES") = vbNullString Then Label2(6) = rs("ANQUES") Else Label2(6) = "----"
If Not rs("ADMIN") = vbNullString Then Label2(7) = rs("ADMIN") Else Label2(7) = "----"
If Not rs("MSG") = vbNullString Then Label2(8) = rs("MSG") Else Label2(8) = "----"

Send.Caption = "Send To " & List1.List(List1.ListIndex)

End Sub

Private Sub Timer1_Timer()
    m_lAlpha = m_lAlpha + 15
    If (m_lAlpha > 255) Then
        m_lAlpha = 255
        Timer1.Enabled = False
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, m_lAlpha, LWA_ALPHA
    End If
End Sub

Private Sub Timer2_Timer()
    m_lAlpha = m_lAlpha - 15
    If (m_lAlpha < 0) Then
        m_lAlpha = 0
        Unload Me
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, m_lAlpha, LWA_ALPHA
    End If
End Sub


Private Sub Display()
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


Private Sub OpenEdit()
List1.Enabled = False
For hi = 0 To 6
Text2(hi).Visible = True
Next hi
Text2(0) = Label2(1)
Text2(1) = Label2(2)
Text2(2) = Label2(3)
Text2(3) = Label2(5)
Text2(4) = Label2(6)
Text2(5) = Label2(7)
Text2(6) = Label2(8)
End Sub

Private Sub CloseEdit()
List1.Enabled = True
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & List1.List(List1.ListIndex) & "'")
rs.Edit
rs("PASS") = Text2(0)
rs("NAME") = Text2(1)
rs("OLD") = Text2(2)
rs("QUES") = Text2(3)
rs("ANQUES") = Text2(4)
rs("ADMIN") = Text2(5)
rs("MSG") = Text2(6)
rs.Update

Label2(0) = rs("USER")
Label2(1) = rs("PASS")
If Not rs("NAME") = vbNullString Then Label2(2) = rs("NAME") Else Label2(2) = "----"
If Not rs("OLD") = vbNullString Then Label2(3) = rs("OLD") Else Label2(3) = "----"
If Not rs("ONLINE") = vbNullString Then Label2(4) = rs("ONLINE") Else Label2(4) = "----"
If Not rs("ONLINE") = vbNullString Then Label2(5) = rs("QUES") Else Label2(5) = "----"
If Not rs("ANQUES") = vbNullString Then Label2(6) = rs("ANQUES") Else Label2(6) = "----"
If Not rs("ADMIN") = vbNullString Then Label2(7) = rs("ADMIN") Else Label2(7) = "----"
If Not rs("MSG") = vbNullString Then Label2(8) = rs("MSG") Else Label2(8) = "----"


For hi = 0 To 6
Text2(hi).Visible = False
Next hi

End Sub

Private Sub Timer3_Timer()

If KiemTra = 1 Then
Lo = ReadIniFile("Z:\ChatLog.log", "ChatRoom", "Log", "No Data")
txtChat.Caption = txtChat.Caption & Lo & vbCrLf
End If
Li = ReadIniFile("Z:\ChatLog.log", "ChatRoom", "Log", "No Data")
If Li = Lo Then
KiemTra = 0
Else
KiemTra = 1
End If
End Sub

Private Sub txtChat_Click()
txtChat.Caption = ""
End Sub
