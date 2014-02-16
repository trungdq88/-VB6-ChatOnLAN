VERSION 5.00
Begin VB.Form frmMem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
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
   ScaleHeight     =   5055
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin ChatOnLan.Label Label5 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChatOnLan.Label Label4 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
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
   Begin ChatOnLan.Label Label3 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   12640511
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
   Begin ChatOnLan.Label Label2 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   12640511
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
   Begin ChatOnLan.Label Label1 
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   12640511
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Label6"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   6015
   End
End
Attribute VB_Name = "frmMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace

Dim X

Private Sub Form_Load()
Label1.Caption = " " & ChrW$(&H56) & ChrW$(&HE0) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H50) & ChrW$(&H68) & ChrW$(&HF2) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&HE1) & ChrW$(&H74)
Label2.Caption = ChrW$(&H54) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1ED5) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H69) & ChrW$(&H6E)
Label3.Caption = " " & ChrW$(&H54) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H72) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H6F) & ChrW$(&HE0) & ChrW$(&H69)
Label4.Caption = ChrW$(&H58) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&HE0) & ChrW$(&H6F) & ChrW$(&H20)
Label5.Caption = ChrW$(&H43) & ChrW$(&H68) & ChrW$(&HFA) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1ED9) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H75) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H1EBB)

End Sub




Private Sub Form_Unload(Cancel As Integer)

Offline
frmMain.Show
End Sub

Private Sub Label1_Click()

frmChat.Show
frmChat.txtName.Text = Label6.Caption
frmChat.Caption = "Chat - " & Label6.Caption
Me.Hide
End Sub

Private Sub Label2_Click()
frmChange.Show
End Sub

Private Sub Label3_Click()
Unload Me
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
Set rs = db.OpenRecordset("Select * from Room where ListUser = '" & Label6.Caption & "'")
rs.Delete

Set rs = db.OpenRecordset("Chat", dbOpenTable)

Set rs = db.OpenRecordset("Select * from Chat where User = '" & Label6.Caption & "'")
rs.Edit
rs("ONLINE") = "no"
rs("MSG") = "no"
rs.Update

End Sub
