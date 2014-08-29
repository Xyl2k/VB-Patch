VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   453
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pSp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   0
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   0
      Top             =   0
      Width           =   6795
   End
   Begin VB.Timer tU 
      Interval        =   35
      Left            =   6360
      Top             =   3000
   End
   Begin VB.Shape s 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00716B64&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      Top             =   3000
      Width           =   435
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Sub Form_Click()
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'Form1 always visible
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_Click
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
    Me.Caption = "Credits"
    MakeTransparent Me.hwnd, 240
    vTop = Me.pSp.Height
        
 CrdLines = Split(" °°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°" & vbNewLine _
                & " ±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±" & vbNewLine _
                & " ²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²" & vbNewLine _
                & " ²²ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿²²" & vbNewLine _
                & " ²²³                                                                   ³²²" & vbNewLine _
                & " ²²³    ÛÛÛ  ÛÛÛ  ÛÛÛ ÛÛÛ  ÛÛÛ       ÛÛÛ  ÛÛÛÛÛÛÛ   ÛÛÛÛÛÛ   ÛÛÛ       ³²²" & vbNewLine _
                & " ²²³    ÛÛÛ  ÛÛÛ  ÛÛÛ ÛÛÛ  ÛÛÛ       ÛÛÛ  ÛÛÛÛÛÛÛ  ÛÛÛÛÛÛÛÛ  ÛÛÛ       ³²²" & vbNewLine _
                & " ²²³    ÛÛ²  ²ÛÛ  ÛÛ² ²ÛÛ  ÛÛ²       ÛÛ²    ÛÛ²    ÛÛ²  ÛÛÛ  ÛÛ²       ³²²" & vbNewLine _
                & " ²²³    ²Û²  Û²²  ²Û² Û²²  ²Û²       ²Û²    ²Û²    ²Û²  Û²Û  ²Û²       ³²²" & vbNewLine _
                & " ²²³     ²ÛÛ²Û²    ²Û²Û²   Û²²       ²²Û    Û²²    Û²Û  ²Û²  Û²²       ³²²" & vbNewLine _
                & " ²²³      Û²²²      Û²²²   ²²²       ²²²    ²²²    ²Û²  ²²²  ²²²       ³²²" & vbNewLine _
                & " ²²³     ²± ±²²     ²²±    ²²±       ²²±    ²²±    ²²±  ²²²  ²²±       ³²²" & vbNewLine _
                & " ²²³    ±²±  ²±²    ±²±     ±²±      ±²±    ±²±    ±²±  ²±²   ±²±      ³²²" & vbNewLine _
                & " ²²³     ±±  ±±±     ±±     ±± ±±±±   ±±     ±±    ±±±±± ±±   ±± ±±±±  ³²²" & vbNewLine _
                & " ²²³     ±   ±±      ±     ± ±± ± ±  ±       ±      ± ±  ±   ± ±± ± ±  ³²²" & vbNewLine _
                & " ²²³                                                                   ³²²" & vbNewLine _
                & " ²²³        BlaBlaBlaBla                                               ³²²" & vbNewLine _
                & " ²²³                                                                   ³²²" & vbNewLine _
                & " ²²³  This nice scroller was coded by: ata (NOHK)                      ³²²" & vbNewLine _
                & " ²²³  SFX.: kenet!ribbon.99, Original By Scavenger!synergy             ³²²" & vbNewLine _
                & " ²²³                                                                   ³²²" & vbNewLine _
                & " ²²ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ²²" & vbNewLine _
                & " ²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²" & vbNewLine _
                & " ±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±" & vbNewLine _
                & " °°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°" & vbNewLine, vbNewLine)

End Sub

Private Sub Form_Resize()
    s.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Private Sub l_Click(Index As Integer)
    Call Form_Click
End Sub
Private Sub Picture1_Click()
    Call Form_Click
End Sub
Private Sub pSp_Click()
    Call Form_Click
End Sub
Private Sub tU_Timer()
    Dim X As Integer
    Dim nTop As Long
    Me.pSp.Cls
    nTop = vTop
    For X = 0 To UBound(CrdLines)
        'if the 'top' is inside the picturebox then draw
        If nTop > -50 And nTop < pSp.Height Then SendCredits pSp, CrdLines(X), 1, nTop, RGB(161, 255, 66), RGB(161, 66, 255), RGB(255, 161, 66), 1 / 6
        nTop = nTop + pSp.TextHeight(CrdLines(X))
    Next X
    'Reloading at the end of the file
    If vTop + 20 < -pSp.TextHeight("A") * UBound(CrdLines) Then vTop = pSp.Height
    vTop = vTop - 0.6
End Sub
