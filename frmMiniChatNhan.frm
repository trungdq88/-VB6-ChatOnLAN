VERSION 5.00
Begin VB.Form frmMiniChatNhan 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4575
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   735
      Left            =   4800
      TabIndex        =   1
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   3960
   End
End
Attribute VB_Name = "frmMiniChatNhan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lo As String
Dim li As String
Dim KiemTra
Dim MM As String
Dim Nhan As String
Dim Gui As String
Private Sub cmdSend_Click()

If Text1.Text <> "" Then
WriteIniFile "Z:\ChatLog.log", MM, "Text", frmChat.txtName.Text & " : " & Text1.Text
Text1.Text = ""
End If



End Sub



Private Sub Form_Load()


Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & frmChat.txtName.Text & "'")

MM = rs("MSG")



KiemTra = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = ("Z:\CSDL.MDB")
    PwdString = "881817258"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Chat", dbOpenTable)
Set rs = db.OpenRecordset("Select * from Chat where USER = '" & frmChat.txtName.Text & "'")
rs.Edit
rs("MSG") = "no"
rs.Update


End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
cmdSend_Click
End If
End Sub

Private Sub Timer1_Timer()
If KiemTra = 1 Then
lo = ReadIniFile("Z:\ChatLog.log", MM, "Text", "")
Text2.Text = Text2.Text & lo & vbCrLf
Text2.SelStart = Len(Text2.Text)
End If
li = ReadIniFile("Z:\ChatLog.log", MM, "Text", "")
If li = lo Then
KiemTra = 0
Else
KiemTra = 1
End If
End Sub

