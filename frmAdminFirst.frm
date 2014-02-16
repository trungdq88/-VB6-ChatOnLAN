VERSION 5.00
Begin VB.Form frmAdminFirst 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   Picture         =   "frmAdminFirst.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Left            =   2040
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   1800
   End
End
Attribute VB_Name = "frmAdminFirst"
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

Private Sub Form_Click()
If Timer1.Enabled = False Then
Timer2.Enabled = True
End If
End Sub

Private Sub Form_Load()
'**************************************************
'Ho tro mo dan Form *******************************
    Dim lStyle As Long
    lStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, lStyle
    SetLayeredWindowAttributes Me.hWnd, 0, 0, LWA_ALPHA
    Timer1.Interval = 100
    Timer2.Interval = 50
    Timer2.Enabled = False
    Timer1.Enabled = True
'**************************************************
'**************************************************

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Cancel = True
        Timer2.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAdmin.Show
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

