VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "XYLITOL PROUDLY PRESENTS :"
   ClientHeight    =   4425
   ClientLeft      =   -1140
   ClientTop       =   4185
   ClientWidth     =   6630
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":E4F5
   Picture         =   "Form1.frx":E7FF
   ScaleHeight     =   4425
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label CRC32 
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   250
      Left            =   5160
      Top             =   1320
      Width           =   1000
   End
   Begin VB.Image BackupOk 
      Height          =   105
      Left            =   3795
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":746BD
      Top             =   4000
      Width           =   135
   End
   Begin VB.Image BackupNo 
      Height          =   105
      Left            =   3790
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":747C3
      Top             =   4000
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image SoundOff 
      Height          =   105
      Left            =   5160
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":748C9
      Top             =   4000
      Width           =   135
   End
   Begin VB.Image ExitMouseUp 
      Height          =   225
      Left            =   5880
      Picture         =   "Form1.frx":749CF
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image ExitMouseDown 
      Height          =   225
      Left            =   5880
      Picture         =   "Form1.frx":74CA5
      Top             =   360
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image ExitMouseMove 
      Height          =   225
      Left            =   5880
      Picture         =   "Form1.frx":74F7B
      Top             =   600
      Width           =   210
   End
   Begin VB.Image PatchMouseUp 
      Height          =   225
      Left            =   4800
      Picture         =   "Form1.frx":75251
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image PatchMouseDown 
      Height          =   225
      Left            =   4800
      Picture         =   "Form1.frx":75B7B
      Top             =   360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image PatchMouseMove 
      Height          =   225
      Left            =   4800
      Picture         =   "Form1.frx":764A5
      Top             =   600
      Width           =   750
   End
   Begin VB.Label Date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09/06/2010"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3430
      TabIndex        =   3
      Top             =   3500
      Width           =   870
   End
   Begin VB.Label Author 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Xylitol"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3680
      TabIndex        =   2
      Top             =   3090
      Width           =   405
   End
   Begin VB.Label AppTarget 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CrackMe.exe"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3650
      TabIndex        =   1
      Top             =   2310
      Width           =   945
   End
   Begin VB.Label AppUrl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://xylitol.free.fr/"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3300
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   2685
      Width           =   1365
   End
   Begin VB.Image SoundPlay 
      Height          =   105
      Left            =   5160
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":76DCF
      Top             =   4000
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'     ____  ___       .__   .__   __          .__
'     \   \/  /___.__.|  |  |__|_/  |_  ____  |  |
'      \     /<   |  ||  |  |  |\   __\/  _ \ |  |
'      /     \ \___  ||  |__|  | |  | (  <_> )|  |__
'     /___/\  \/ ____||____/|__| |__|  \____/ |____/
'           \_/\/                               2k10
' Thanks to salazar, Robert Gainor, ata, Ed Preston, [x]sp!d3r, DonGkeY, KKR & all hardworking sceners

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim bytRegion(3199) As Byte
Dim nBytes As Long
Dim Offst(0 To 1) As Long 'Change If needed
Dim DataP(0 To 1) As Byte 'Change If needed
Dim Fhand As Integer
Dim i As Integer
Dim p As Integer
Dim Answer As Integer
Dim Temp As String
' Variable to hold the instance of the CRC32 Class
Private objCRC32 As clsCRC32
Const CHUNK_SIZE = 2048


Private Sub Form_Load()
 Set objCRC32 = New clsCRC32
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'Form1 always visible
'GFX Initialisation
PatchMouseMove.Top = 4170
PatchMouseDown.Top = 4170
PatchMouseUp.Top = 4170
ExitMouseMove.Top = 1650
ExitMouseDown.Top = 1650
ExitMouseUp.Top = 1650
'Check if the patch run already
 If App.PrevInstance = True Then
  End
 End If
'Play chiptune
  Call playXM("credits screen.xm", 102)
'Image background transparency By Robert Gainor
    Dim rgnMain As Long
    nBytes = 3200
    LoadBytes
    rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
    SetWindowRgn Me.hwnd, rgnMain, True
End Sub
'Move code
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
AppUrl.FontUnderline = False
PatchMouseMove.Visible = True
PatchMouseDown.Visible = False
PatchMouseUp.Visible = False
ExitMouseMove.Visible = True
ExitMouseDown.Visible = False
ExitMouseUp.Visible = False
End Sub
Private Sub AppTarget_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
    End Sub
Private Sub Author_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
    End Sub
    Private Sub Date_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
    End Sub
    
Private Sub AppURL_Click()
'Website
ShellExecute hwnd, "Open", "http://xylitol.free.fr/", "", App.Path, 1 'change
End Sub

Private Sub AppURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AppUrl.FontUnderline = True
End Sub

Function FileExists(dzaFile As String) As Boolean
On Error Resume Next
Err.Clear
GetAttr dzaFile
FileExists = (Err = 0)
End Function

'Exit Button
Private Sub ExitMouseMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ExitMouseMove.Visible = False
ExitMouseDown.Visible = True
End Sub
Private Sub ExitMouseDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ExitMouseDown.Visible = False
ExitMouseUp.Visible = True
End Sub
Private Sub ExitMouseDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Bye :')
ExitMouseUp.Visible = False
ExitMouseMove.Visible = True
    Call stopXM("credits screen.xm")
     Set objCRC32 = Nothing
    End
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) 'Press Escape ? bye !
vbKeyEscape:
            Call stopXM("credits screen.xm")
             Set objCRC32 = Nothing
            End
End Sub



Private Sub Image1_Click()
'SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'Form1 normal
SetWindowPos Form2.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'Form2 always visible
    Load Form2
    Form2.Left = Form1.Left + (Form1.ScaleWidth / 0.31)
    Form2.Top = Form1.Top + (Form1.ScaleWidth / 0.32)
    Call Form2.Show(1)
End Sub



'Patch Button
Private Sub PatchMouseMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PatchMouseMove.Visible = False
PatchMouseDown.Visible = True
End Sub
Private Sub PatchMouseDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PatchMouseDown.Visible = False
PatchMouseUp.Visible = True
End Sub

Private Sub PatchMouseDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PatchMouseUp.Visible = False
PatchMouseMove.Visible = True
If Not FileExists("CrackMe.exe") Then   'First check to see if the target file exists

GoTo NOexist:
NOexist:
Answer = MsgBox("      [CrackMe.exe] not found ! " & vbCrLf & _
                "                                " & vbCrLf & _
                "      Search  ?                 ", _
   vbQuestion + vbYesNo, "Search ?")
   If Answer = vbYes Then

CommonDialog1.InitDir = "C:\"
CommonDialog1.DialogTitle = "Searching for CrackMe.exe"
CommonDialog1.FileName = ""
CommonDialog1.Filter = "CrackMe.exe|*.exe;*.exe"
CommonDialog1.ShowOpen
End If
If CommonDialog1.FileName = "" Then Exit Sub
If CommonDialog1.FileTitle <> "CrackMe.exe" Then
MsgBox "Please Give the correct File To patch !", vbCritical, "Error"
Exit Sub
End If
If FileLen("CrackMe.exe") <> 20480 Then 'File size
  MsgBox "The patch is for CrackMe.exe only !!", vbCritical, "Error"
Exit Sub
End If
Else: GoTo Exist:
Exist:
If FileLen("CrackMe.exe") <> 20480 Then 'File size
    MsgBox "The patch is for CrackMe.exe only !!", vbCritical, "Error"
Exit Sub
End If
End If
CRC32.Caption = CRCFromFile(App.Path & "/" & AppTarget) 'CRC32 Check
If CRC32.Caption = "2C667B30" Then
Else
MsgBox "CrackMe.exe bad CRC32 Check, The file was already patched ?", vbCritical, "Error"
Exit Sub
End If
If BackupOk.Visible = True Then 'Do you want make a backup ?
FileCopy App.Path & "/" & AppTarget, App.Path & "/" & AppTarget + ".bak"
Else
End If
Fhand = FreeFile

'Offset patch: 00401595
'              00401596
Offst(0) = &H1595
Offst(1) = &H1596

'Fill the array of data

DataP(0) = &H90 '75 Replaced by 90 (nop)
DataP(1) = &H90 '16 Replaced by 90 (nop)

Open "CrackMe.exe" For Binary As Fhand 'Change this

   For i = 0 To 1 'Change this


   Put Fhand, Offst(i) + 1, DataP(i)
   Next i

Close Fhand
MsgBox "CrackMe.exe Patched successfuly !", vbOKOnly + vbInformation + vbApplicationModal, "Done !"
CommonDialog1.FileName = ""
End Sub

' Generated by: "A Transparent Form Maker" By Robert Gainor
' https://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50405&lngWId=1
Private Sub LoadBytes()
bytRegion(0) = 32
bytRegion(4) = 1
bytRegion(8) = 198
bytRegion(12) = 96
bytRegion(13) = 12
bytRegion(16) = 20
bytRegion(24) = 185
bytRegion(25) = 1
bytRegion(28) = 37
bytRegion(29) = 1
bytRegion(32) = 118
bytRegion(40) = 120
bytRegion(44) = 1
bytRegion(48) = 117
bytRegion(52) = 1
bytRegion(56) = 121
bytRegion(60) = 2
bytRegion(64) = 106
bytRegion(68) = 2
bytRegion(72) = 107
bytRegion(76) = 3
bytRegion(80) = 114
bytRegion(84) = 2
bytRegion(88) = 121
bytRegion(92) = 3
bytRegion(96) = 105
bytRegion(100) = 3
bytRegion(104) = 108
bytRegion(108) = 4
bytRegion(112) = 109
bytRegion(116) = 3
bytRegion(120) = 123
bytRegion(124) = 4
bytRegion(128) = 105
bytRegion(132) = 4
bytRegion(136) = 126
bytRegion(140) = 5
bytRegion(144) = 105
bytRegion(148) = 5
bytRegion(152) = 128
bytRegion(156) = 6
bytRegion(160) = 129
bytRegion(164) = 5
bytRegion(168) = 131
bytRegion(172) = 6
bytRegion(176) = 103
bytRegion(180) = 6
bytRegion(184) = 104
bytRegion(188) = 7
bytRegion(192) = 105
bytRegion(196) = 6
bytRegion(200) = 132
bytRegion(204) = 7
bytRegion(208) = 102
bytRegion(212) = 7
bytRegion(216) = 132
bytRegion(220) = 8
bytRegion(224) = 103
bytRegion(228) = 8
bytRegion(232) = 132
bytRegion(236) = 9
bytRegion(240) = 100
bytRegion(244) = 9
bytRegion(248) = 134
bytRegion(252) = 10
bytRegion(256) = 99
bytRegion(260) = 10
bytRegion(264) = 135
bytRegion(268) = 11
bytRegion(272) = 98
bytRegion(276) = 11
bytRegion(280) = 136
bytRegion(284) = 12
bytRegion(288) = 97
bytRegion(292) = 12
bytRegion(296) = 138
bytRegion(300) = 13
bytRegion(304) = 96
bytRegion(308) = 13
bytRegion(312) = 139
bytRegion(316) = 14
bytRegion(320) = 96
bytRegion(324) = 14
bytRegion(328) = 140
bytRegion(332) = 15
bytRegion(336) = 96
bytRegion(340) = 15
bytRegion(344) = 141
bytRegion(348) = 17
bytRegion(352) = 95
bytRegion(356) = 17
bytRegion(360) = 141
bytRegion(364) = 18
bytRegion(368) = 94
bytRegion(372) = 18
bytRegion(376) = 141
bytRegion(380) = 19
bytRegion(384) = 93
bytRegion(388) = 19
bytRegion(392) = 141
bytRegion(396) = 20
bytRegion(400) = 93
bytRegion(404) = 20
bytRegion(408) = 142
bytRegion(412) = 21
bytRegion(416) = 93
bytRegion(420) = 21
bytRegion(424) = 143
bytRegion(428) = 24
bytRegion(432) = 93
bytRegion(436) = 24
bytRegion(440) = 142
bytRegion(444) = 26
bytRegion(448) = 94
bytRegion(452) = 26
bytRegion(456) = 142
bytRegion(460) = 27
bytRegion(464) = 94
bytRegion(468) = 27
bytRegion(472) = 143
bytRegion(476) = 29
bytRegion(480) = 93
bytRegion(484) = 29
bytRegion(488) = 143
bytRegion(492) = 30
bytRegion(496) = 92
bytRegion(500) = 30
bytRegion(504) = 143
bytRegion(508) = 33
bytRegion(512) = 93
bytRegion(516) = 33
bytRegion(520) = 143
bytRegion(524) = 34
bytRegion(528) = 94
bytRegion(532) = 34
bytRegion(536) = 142
bytRegion(540) = 35
bytRegion(544) = 94
bytRegion(548) = 35
bytRegion(552) = 141
bytRegion(556) = 36
bytRegion(560) = 95
bytRegion(564) = 36
bytRegion(568) = 141
bytRegion(572) = 38
bytRegion(576) = 94
bytRegion(580) = 38
bytRegion(584) = 141
bytRegion(588) = 39
bytRegion(592) = 89
bytRegion(596) = 39
bytRegion(600) = 140
bytRegion(604) = 40
bytRegion(608) = 87
bytRegion(612) = 40
bytRegion(616) = 140
bytRegion(620) = 41
bytRegion(624) = 86
bytRegion(628) = 41
bytRegion(632) = 139
bytRegion(636) = 42
bytRegion(640) = 85
bytRegion(644) = 42
bytRegion(648) = 145
bytRegion(652) = 43
bytRegion(656) = 85
bytRegion(660) = 43
bytRegion(664) = 146
bytRegion(668) = 52
bytRegion(672) = 84
bytRegion(676) = 52
bytRegion(680) = 147
bytRegion(684) = 53
bytRegion(688) = 84
bytRegion(692) = 53
bytRegion(696) = 148
bytRegion(700) = 54
bytRegion(704) = 71
bytRegion(708) = 54
bytRegion(712) = 150
bytRegion(716) = 55
bytRegion(720) = 70
bytRegion(724) = 55
bytRegion(728) = 152
bytRegion(732) = 56
bytRegion(736) = 69
bytRegion(740) = 56
bytRegion(744) = 153
bytRegion(748) = 57
bytRegion(752) = 68
bytRegion(756) = 57
bytRegion(760) = 155
bytRegion(764) = 58
bytRegion(768) = 65
bytRegion(772) = 58
bytRegion(776) = 156
bytRegion(780) = 59
bytRegion(784) = 63
bytRegion(788) = 59
bytRegion(792) = 161
bytRegion(796) = 60
bytRegion(800) = 60
bytRegion(804) = 60
bytRegion(808) = 162
bytRegion(812) = 61
bytRegion(816) = 59
bytRegion(820) = 61
bytRegion(824) = 163
bytRegion(828) = 62
bytRegion(832) = 57
bytRegion(836) = 62
bytRegion(840) = 166
bytRegion(844) = 63
bytRegion(848) = 56
bytRegion(852) = 63
bytRegion(856) = 168
bytRegion(860) = 64
bytRegion(864) = 55
bytRegion(868) = 64
bytRegion(872) = 170
bytRegion(876) = 65
bytRegion(880) = 53
bytRegion(884) = 65
bytRegion(888) = 171
bytRegion(892) = 66
bytRegion(896) = 52
bytRegion(900) = 66
bytRegion(904) = 173
bytRegion(908) = 67
bytRegion(912) = 51
bytRegion(916) = 67
bytRegion(920) = 174
bytRegion(924) = 68
bytRegion(928) = 51
bytRegion(932) = 68
bytRegion(936) = 175
bytRegion(940) = 69
bytRegion(944) = 50
bytRegion(948) = 69
bytRegion(952) = 176
bytRegion(956) = 70
bytRegion(960) = 50
bytRegion(964) = 70
bytRegion(968) = 177
bytRegion(972) = 71
bytRegion(976) = 50
bytRegion(980) = 71
bytRegion(984) = 178
bytRegion(988) = 72
bytRegion(992) = 50
bytRegion(996) = 72
bytRegion(1000) = 179
bytRegion(1004) = 73
bytRegion(1008) = 50
bytRegion(1012) = 73
bytRegion(1016) = 178
bytRegion(1020) = 74
bytRegion(1024) = 50
bytRegion(1028) = 74
bytRegion(1032) = 175
bytRegion(1036) = 75
bytRegion(1040) = 176
bytRegion(1044) = 74
bytRegion(1048) = 177
bytRegion(1052) = 75
bytRegion(1056) = 51
bytRegion(1060) = 75
bytRegion(1064) = 174
bytRegion(1068) = 76
bytRegion(1072) = 52
bytRegion(1076) = 76
bytRegion(1080) = 175
bytRegion(1084) = 77
bytRegion(1088) = 53
bytRegion(1092) = 77
bytRegion(1096) = 175
bytRegion(1100) = 78
bytRegion(1104) = 52
bytRegion(1108) = 78
bytRegion(1112) = 176
bytRegion(1116) = 79
bytRegion(1120) = 52
bytRegion(1124) = 79
bytRegion(1128) = 177
bytRegion(1132) = 80
bytRegion(1136) = 52
bytRegion(1140) = 80
bytRegion(1144) = 178
bytRegion(1148) = 81
bytRegion(1152) = 51
bytRegion(1156) = 81
bytRegion(1160) = 179
bytRegion(1164) = 82
bytRegion(1168) = 51
bytRegion(1172) = 82
bytRegion(1176) = 180
bytRegion(1180) = 83
bytRegion(1184) = 51
bytRegion(1188) = 83
bytRegion(1192) = 181
bytRegion(1196) = 84
bytRegion(1200) = 52
bytRegion(1204) = 84
bytRegion(1208) = 182
bytRegion(1212) = 85
bytRegion(1216) = 51
bytRegion(1220) = 85
bytRegion(1224) = 182
bytRegion(1228) = 86
bytRegion(1232) = 51
bytRegion(1236) = 86
bytRegion(1240) = 183
bytRegion(1244) = 87
bytRegion(1248) = 51
bytRegion(1252) = 87
bytRegion(1256) = 184
bytRegion(1260) = 88
bytRegion(1264) = 50
bytRegion(1268) = 88
bytRegion(1272) = 184
bytRegion(1276) = 89
bytRegion(1280) = 50
bytRegion(1284) = 89
bytRegion(1288) = 184
bytRegion(1292) = 90
bytRegion(1296) = 84
bytRegion(1297) = 1
bytRegion(1300) = 89
bytRegion(1304) = 158
bytRegion(1305) = 1
bytRegion(1308) = 90
bytRegion(1312) = 50
bytRegion(1316) = 90
bytRegion(1320) = 185
bytRegion(1324) = 91
bytRegion(1328) = 84
bytRegion(1329) = 1
bytRegion(1332) = 90
bytRegion(1336) = 158
bytRegion(1337) = 1
bytRegion(1340) = 91
bytRegion(1344) = 50
bytRegion(1348) = 91
bytRegion(1352) = 186
bytRegion(1356) = 92
bytRegion(1360) = 84
bytRegion(1361) = 1
bytRegion(1364) = 91
bytRegion(1368) = 158
bytRegion(1369) = 1
bytRegion(1372) = 92
bytRegion(1376) = 50
bytRegion(1380) = 92
bytRegion(1384) = 187
bytRegion(1388) = 93
bytRegion(1392) = 84
bytRegion(1393) = 1
bytRegion(1396) = 92
bytRegion(1400) = 158
bytRegion(1401) = 1
bytRegion(1404) = 93
bytRegion(1408) = 49
bytRegion(1412) = 93
bytRegion(1416) = 188
bytRegion(1420) = 94
bytRegion(1424) = 84
bytRegion(1425) = 1
bytRegion(1428) = 93
bytRegion(1432) = 158
bytRegion(1433) = 1
bytRegion(1436) = 94
bytRegion(1440) = 49
bytRegion(1444) = 94
bytRegion(1448) = 71
bytRegion(1449) = 1
bytRegion(1452) = 95
bytRegion(1456) = 84
bytRegion(1457) = 1
bytRegion(1460) = 94
bytRegion(1464) = 158
bytRegion(1465) = 1
bytRegion(1468) = 95
bytRegion(1472) = 50
bytRegion(1476) = 95
bytRegion(1480) = 73
bytRegion(1481) = 1
bytRegion(1484) = 96
bytRegion(1488) = 84
bytRegion(1489) = 1
bytRegion(1492) = 95
bytRegion(1496) = 158
bytRegion(1497) = 1
bytRegion(1500) = 96
bytRegion(1504) = 50
bytRegion(1508) = 96
bytRegion(1512) = 74
bytRegion(1513) = 1
bytRegion(1516) = 97
bytRegion(1520) = 84
bytRegion(1521) = 1
bytRegion(1524) = 96
bytRegion(1528) = 158
bytRegion(1529) = 1
bytRegion(1532) = 97
bytRegion(1536) = 51
bytRegion(1540) = 97
bytRegion(1544) = 76
bytRegion(1545) = 1
bytRegion(1548) = 98
bytRegion(1552) = 84
bytRegion(1553) = 1
bytRegion(1556) = 97
bytRegion(1560) = 158
bytRegion(1561) = 1
bytRegion(1564) = 98
bytRegion(1568) = 51
bytRegion(1572) = 98
bytRegion(1576) = 77
bytRegion(1577) = 1
bytRegion(1580) = 99
bytRegion(1584) = 84
bytRegion(1585) = 1
bytRegion(1588) = 98
bytRegion(1592) = 158
bytRegion(1593) = 1
bytRegion(1596) = 99
bytRegion(1600) = 51
bytRegion(1604) = 99
bytRegion(1608) = 79
bytRegion(1609) = 1
bytRegion(1612) = 100
bytRegion(1616) = 84
bytRegion(1617) = 1
bytRegion(1620) = 99
bytRegion(1624) = 158
bytRegion(1625) = 1
bytRegion(1628) = 100
bytRegion(1632) = 52
bytRegion(1636) = 100
bytRegion(1640) = 81
bytRegion(1641) = 1
bytRegion(1644) = 101
bytRegion(1648) = 84
bytRegion(1649) = 1
bytRegion(1652) = 100
bytRegion(1656) = 158
bytRegion(1657) = 1
bytRegion(1660) = 101
bytRegion(1664) = 53
bytRegion(1668) = 101
bytRegion(1672) = 82
bytRegion(1673) = 1
bytRegion(1676) = 102
bytRegion(1680) = 84
bytRegion(1681) = 1
bytRegion(1684) = 101
bytRegion(1688) = 158
bytRegion(1689) = 1
bytRegion(1692) = 102
bytRegion(1696) = 53
bytRegion(1700) = 102
bytRegion(1704) = 158
bytRegion(1705) = 1
bytRegion(1708) = 103
bytRegion(1712) = 54
bytRegion(1716) = 103
bytRegion(1720) = 158
bytRegion(1721) = 1
bytRegion(1724) = 104
bytRegion(1728) = 53
bytRegion(1732) = 104
bytRegion(1736) = 158
bytRegion(1737) = 1
bytRegion(1740) = 105
bytRegion(1744) = 54
bytRegion(1748) = 105
bytRegion(1752) = 158
bytRegion(1753) = 1
bytRegion(1756) = 106
bytRegion(1760) = 55
bytRegion(1764) = 106
bytRegion(1768) = 158
bytRegion(1769) = 1
bytRegion(1772) = 107
bytRegion(1776) = 54
bytRegion(1780) = 107
bytRegion(1784) = 158
bytRegion(1785) = 1
bytRegion(1788) = 109
bytRegion(1792) = 54
bytRegion(1796) = 109
bytRegion(1800) = 166
bytRegion(1801) = 1
bytRegion(1804) = 112
bytRegion(1808) = 54
bytRegion(1812) = 112
bytRegion(1816) = 165
bytRegion(1817) = 1
bytRegion(1820) = 113
bytRegion(1824) = 53
bytRegion(1828) = 113
bytRegion(1832) = 164
bytRegion(1833) = 1
bytRegion(1836) = 114
bytRegion(1840) = 53
bytRegion(1844) = 114
bytRegion(1848) = 163
bytRegion(1849) = 1
bytRegion(1852) = 115
bytRegion(1856) = 53
bytRegion(1860) = 115
bytRegion(1864) = 162
bytRegion(1865) = 1
bytRegion(1868) = 116
bytRegion(1872) = 53
bytRegion(1876) = 116
bytRegion(1880) = 161
bytRegion(1881) = 1
bytRegion(1884) = 117
bytRegion(1888) = 53
bytRegion(1892) = 117
bytRegion(1896) = 160
bytRegion(1897) = 1
bytRegion(1900) = 118
bytRegion(1904) = 53
bytRegion(1908) = 118
bytRegion(1912) = 159
bytRegion(1913) = 1
bytRegion(1916) = 119
bytRegion(1920) = 53
bytRegion(1924) = 119
bytRegion(1928) = 158
bytRegion(1929) = 1
bytRegion(1932) = 120
bytRegion(1936) = 53
bytRegion(1940) = 120
bytRegion(1944) = 157
bytRegion(1945) = 1
bytRegion(1948) = 121
bytRegion(1952) = 52
bytRegion(1956) = 121
bytRegion(1960) = 156
bytRegion(1961) = 1
bytRegion(1964) = 122
bytRegion(1968) = 52
bytRegion(1972) = 122
bytRegion(1976) = 155
bytRegion(1977) = 1
bytRegion(1980) = 123
bytRegion(1984) = 52
bytRegion(1988) = 123
bytRegion(1992) = 154
bytRegion(1993) = 1
bytRegion(1996) = 124
bytRegion(2000) = 52
bytRegion(2004) = 124
bytRegion(2008) = 165
bytRegion(2009) = 1
bytRegion(2012) = 127
bytRegion(2016) = 51
bytRegion(2020) = 127
bytRegion(2024) = 164
bytRegion(2025) = 1
bytRegion(2028) = 128
bytRegion(2032) = 51
bytRegion(2036) = 128
bytRegion(2040) = 162
bytRegion(2041) = 1
bytRegion(2044) = 129
bytRegion(2048) = 51
bytRegion(2052) = 129
bytRegion(2056) = 160
bytRegion(2057) = 1
bytRegion(2060) = 130
bytRegion(2064) = 51
bytRegion(2068) = 130
bytRegion(2072) = 159
bytRegion(2073) = 1
bytRegion(2076) = 131
bytRegion(2080) = 51
bytRegion(2084) = 131
bytRegion(2088) = 157
bytRegion(2089) = 1
bytRegion(2092) = 132
bytRegion(2096) = 51
bytRegion(2100) = 132
bytRegion(2104) = 156
bytRegion(2105) = 1
bytRegion(2108) = 133
bytRegion(2112) = 51
bytRegion(2116) = 133
bytRegion(2120) = 154
bytRegion(2121) = 1
bytRegion(2124) = 134
bytRegion(2128) = 51
bytRegion(2132) = 134
bytRegion(2136) = 163
bytRegion(2137) = 1
bytRegion(2140) = 135
bytRegion(2144) = 50
bytRegion(2148) = 135
bytRegion(2152) = 163
bytRegion(2153) = 1
bytRegion(2156) = 140
bytRegion(2160) = 49
bytRegion(2164) = 140
bytRegion(2168) = 163
bytRegion(2169) = 1
bytRegion(2172) = 148
bytRegion(2176) = 48
bytRegion(2180) = 148
bytRegion(2184) = 163
bytRegion(2185) = 1
bytRegion(2188) = 154
bytRegion(2192) = 47
bytRegion(2196) = 154
bytRegion(2200) = 163
bytRegion(2201) = 1
bytRegion(2204) = 161
bytRegion(2208) = 46
bytRegion(2212) = 161
bytRegion(2216) = 163
bytRegion(2217) = 1
bytRegion(2220) = 166
bytRegion(2224) = 45
bytRegion(2228) = 166
bytRegion(2232) = 163
bytRegion(2233) = 1
bytRegion(2236) = 170
bytRegion(2240) = 44
bytRegion(2244) = 170
bytRegion(2248) = 163
bytRegion(2249) = 1
bytRegion(2252) = 175
bytRegion(2256) = 43
bytRegion(2260) = 175
bytRegion(2264) = 163
bytRegion(2265) = 1
bytRegion(2268) = 178
bytRegion(2272) = 43
bytRegion(2276) = 178
bytRegion(2280) = 162
bytRegion(2281) = 1
bytRegion(2284) = 179
bytRegion(2288) = 42
bytRegion(2292) = 179
bytRegion(2296) = 161
bytRegion(2297) = 1
bytRegion(2300) = 180
bytRegion(2304) = 42
bytRegion(2308) = 180
bytRegion(2312) = 160
bytRegion(2313) = 1
bytRegion(2316) = 181
bytRegion(2320) = 42
bytRegion(2324) = 181
bytRegion(2328) = 159
bytRegion(2329) = 1
bytRegion(2332) = 182
bytRegion(2336) = 42
bytRegion(2340) = 182
bytRegion(2344) = 157
bytRegion(2345) = 1
bytRegion(2348) = 183
bytRegion(2352) = 42
bytRegion(2356) = 183
bytRegion(2360) = 156
bytRegion(2361) = 1
bytRegion(2364) = 184
bytRegion(2368) = 41
bytRegion(2372) = 184
bytRegion(2376) = 155
bytRegion(2377) = 1
bytRegion(2380) = 185
bytRegion(2384) = 41
bytRegion(2388) = 185
bytRegion(2392) = 154
bytRegion(2393) = 1
bytRegion(2396) = 186
bytRegion(2400) = 41
bytRegion(2404) = 186
bytRegion(2408) = 152
bytRegion(2409) = 1
bytRegion(2412) = 187
bytRegion(2416) = 40
bytRegion(2420) = 187
bytRegion(2424) = 151
bytRegion(2425) = 1
bytRegion(2428) = 188
bytRegion(2432) = 40
bytRegion(2436) = 188
bytRegion(2440) = 150
bytRegion(2441) = 1
bytRegion(2444) = 189
bytRegion(2448) = 40
bytRegion(2452) = 189
bytRegion(2456) = 149
bytRegion(2457) = 1
bytRegion(2460) = 193
bytRegion(2464) = 39
bytRegion(2468) = 193
bytRegion(2472) = 149
bytRegion(2473) = 1
bytRegion(2476) = 196
bytRegion(2480) = 38
bytRegion(2484) = 196
bytRegion(2488) = 149
bytRegion(2489) = 1
bytRegion(2492) = 201
bytRegion(2496) = 37
bytRegion(2500) = 201
bytRegion(2504) = 149
bytRegion(2505) = 1
bytRegion(2508) = 205
bytRegion(2512) = 36
bytRegion(2516) = 205
bytRegion(2520) = 149
bytRegion(2521) = 1
bytRegion(2524) = 210
bytRegion(2528) = 35
bytRegion(2532) = 210
bytRegion(2536) = 149
bytRegion(2537) = 1
bytRegion(2540) = 214
bytRegion(2544) = 34
bytRegion(2548) = 214
bytRegion(2552) = 149
bytRegion(2553) = 1
bytRegion(2556) = 218
bytRegion(2560) = 33
bytRegion(2564) = 218
bytRegion(2568) = 149
bytRegion(2569) = 1
bytRegion(2572) = 222
bytRegion(2576) = 32
bytRegion(2580) = 222
bytRegion(2584) = 149
bytRegion(2585) = 1
bytRegion(2588) = 226
bytRegion(2592) = 31
bytRegion(2596) = 226
bytRegion(2600) = 149
bytRegion(2601) = 1
bytRegion(2604) = 232
bytRegion(2608) = 30
bytRegion(2612) = 232
bytRegion(2616) = 149
bytRegion(2617) = 1
bytRegion(2620) = 236
bytRegion(2624) = 29
bytRegion(2628) = 236
bytRegion(2632) = 149
bytRegion(2633) = 1
bytRegion(2636) = 242
bytRegion(2640) = 28
bytRegion(2644) = 242
bytRegion(2648) = 149
bytRegion(2649) = 1
bytRegion(2652) = 246
bytRegion(2656) = 27
bytRegion(2660) = 246
bytRegion(2664) = 149
bytRegion(2665) = 1
bytRegion(2668) = 251
bytRegion(2672) = 26
bytRegion(2676) = 251
bytRegion(2680) = 151
bytRegion(2681) = 1
bytRegion(2684) = 252
bytRegion(2688) = 26
bytRegion(2692) = 252
bytRegion(2696) = 154
bytRegion(2697) = 1
bytRegion(2700) = 253
bytRegion(2704) = 26
bytRegion(2708) = 253
bytRegion(2712) = 157
bytRegion(2713) = 1
bytRegion(2716) = 254
bytRegion(2720) = 26
bytRegion(2724) = 254
bytRegion(2728) = 160
bytRegion(2729) = 1
bytRegion(2732) = 255
bytRegion(2736) = 26
bytRegion(2740) = 255
bytRegion(2744) = 164
bytRegion(2745) = 1
bytRegion(2749) = 1
bytRegion(2752) = 25
bytRegion(2757) = 1
bytRegion(2760) = 167
bytRegion(2761) = 1
bytRegion(2764) = 1
bytRegion(2765) = 1
bytRegion(2768) = 25
bytRegion(2772) = 1
bytRegion(2773) = 1
bytRegion(2776) = 170
bytRegion(2777) = 1
bytRegion(2780) = 2
bytRegion(2781) = 1
bytRegion(2784) = 25
bytRegion(2788) = 2
bytRegion(2789) = 1
bytRegion(2792) = 173
bytRegion(2793) = 1
bytRegion(2796) = 3
bytRegion(2797) = 1
bytRegion(2800) = 25
bytRegion(2804) = 3
bytRegion(2805) = 1
bytRegion(2808) = 177
bytRegion(2809) = 1
bytRegion(2812) = 4
bytRegion(2813) = 1
bytRegion(2816) = 25
bytRegion(2820) = 4
bytRegion(2821) = 1
bytRegion(2824) = 180
bytRegion(2825) = 1
bytRegion(2828) = 5
bytRegion(2829) = 1
bytRegion(2832) = 24
bytRegion(2836) = 5
bytRegion(2837) = 1
bytRegion(2840) = 183
bytRegion(2841) = 1
bytRegion(2844) = 6
bytRegion(2845) = 1
bytRegion(2848) = 24
bytRegion(2852) = 6
bytRegion(2853) = 1
bytRegion(2856) = 185
bytRegion(2857) = 1
bytRegion(2860) = 8
bytRegion(2861) = 1
bytRegion(2864) = 23
bytRegion(2868) = 8
bytRegion(2869) = 1
bytRegion(2872) = 185
bytRegion(2873) = 1
bytRegion(2876) = 12
bytRegion(2877) = 1
bytRegion(2880) = 22
bytRegion(2884) = 12
bytRegion(2885) = 1
bytRegion(2888) = 184
bytRegion(2889) = 1
bytRegion(2892) = 15
bytRegion(2893) = 1
bytRegion(2896) = 21
bytRegion(2900) = 15
bytRegion(2901) = 1
bytRegion(2904) = 184
bytRegion(2905) = 1
bytRegion(2908) = 18
bytRegion(2909) = 1
bytRegion(2912) = 20
bytRegion(2916) = 18
bytRegion(2917) = 1
bytRegion(2920) = 184
bytRegion(2921) = 1
bytRegion(2924) = 19
bytRegion(2925) = 1
bytRegion(2928) = 20
bytRegion(2932) = 19
bytRegion(2933) = 1
bytRegion(2936) = 183
bytRegion(2937) = 1
bytRegion(2940) = 21
bytRegion(2941) = 1
bytRegion(2944) = 21
bytRegion(2948) = 21
bytRegion(2949) = 1
bytRegion(2952) = 183
bytRegion(2953) = 1
bytRegion(2956) = 22
bytRegion(2957) = 1
bytRegion(2960) = 40
bytRegion(2961) = 1
bytRegion(2964) = 22
bytRegion(2965) = 1
bytRegion(2968) = 128
bytRegion(2969) = 1
bytRegion(2972) = 23
bytRegion(2973) = 1
bytRegion(2976) = 41
bytRegion(2977) = 1
bytRegion(2980) = 23
bytRegion(2981) = 1
bytRegion(2984) = 128
bytRegion(2985) = 1
bytRegion(2988) = 24
bytRegion(2989) = 1
bytRegion(2992) = 42
bytRegion(2993) = 1
bytRegion(2996) = 24
bytRegion(2997) = 1
bytRegion(3000) = 127
bytRegion(3001) = 1
bytRegion(3004) = 25
bytRegion(3005) = 1
bytRegion(3008) = 43
bytRegion(3009) = 1
bytRegion(3012) = 25
bytRegion(3013) = 1
bytRegion(3016) = 127
bytRegion(3017) = 1
bytRegion(3020) = 26
bytRegion(3021) = 1
bytRegion(3024) = 44
bytRegion(3025) = 1
bytRegion(3028) = 26
bytRegion(3029) = 1
bytRegion(3032) = 126
bytRegion(3033) = 1
bytRegion(3036) = 27
bytRegion(3037) = 1
bytRegion(3040) = 45
bytRegion(3041) = 1
bytRegion(3044) = 27
bytRegion(3045) = 1
bytRegion(3048) = 125
bytRegion(3049) = 1
bytRegion(3052) = 28
bytRegion(3053) = 1
bytRegion(3056) = 46
bytRegion(3057) = 1
bytRegion(3060) = 28
bytRegion(3061) = 1
bytRegion(3064) = 125
bytRegion(3065) = 1
bytRegion(3068) = 29
bytRegion(3069) = 1
bytRegion(3072) = 47
bytRegion(3073) = 1
bytRegion(3076) = 29
bytRegion(3077) = 1
bytRegion(3080) = 124
bytRegion(3081) = 1
bytRegion(3084) = 30
bytRegion(3085) = 1
bytRegion(3088) = 48
bytRegion(3089) = 1
bytRegion(3092) = 30
bytRegion(3093) = 1
bytRegion(3096) = 124
bytRegion(3097) = 1
bytRegion(3100) = 31
bytRegion(3101) = 1
bytRegion(3104) = 49
bytRegion(3105) = 1
bytRegion(3108) = 31
bytRegion(3109) = 1
bytRegion(3112) = 123
bytRegion(3113) = 1
bytRegion(3116) = 32
bytRegion(3117) = 1
bytRegion(3120) = 50
bytRegion(3121) = 1
bytRegion(3124) = 32
bytRegion(3125) = 1
bytRegion(3128) = 123
bytRegion(3129) = 1
bytRegion(3132) = 33
bytRegion(3133) = 1
bytRegion(3136) = 51
bytRegion(3137) = 1
bytRegion(3140) = 33
bytRegion(3141) = 1
bytRegion(3144) = 122
bytRegion(3145) = 1
bytRegion(3148) = 34
bytRegion(3149) = 1
bytRegion(3152) = 52
bytRegion(3153) = 1
bytRegion(3156) = 34
bytRegion(3157) = 1
bytRegion(3160) = 121
bytRegion(3161) = 1
bytRegion(3164) = 35
bytRegion(3165) = 1
bytRegion(3168) = 53
bytRegion(3169) = 1
bytRegion(3172) = 35
bytRegion(3173) = 1
bytRegion(3176) = 121
bytRegion(3177) = 1
bytRegion(3180) = 36
bytRegion(3181) = 1
bytRegion(3184) = 54
bytRegion(3185) = 1
bytRegion(3188) = 36
bytRegion(3189) = 1
bytRegion(3192) = 120
bytRegion(3193) = 1
bytRegion(3196) = 37
bytRegion(3197) = 1
End Sub

Private Sub BackupOk_Click()
BackupOk.Visible = False
BackupNo.Visible = True
End Sub
Private Sub BackupNo_Click()
BackupNo.Visible = False
BackupOk.Visible = True
End Sub

Private Sub SoundOff_Click()
Call BASSMOD_MusicPause
SoundOff.Visible = False
SoundPlay.Visible = True
End Sub
Private Sub SoundPlay_Click()
Call BASSMOD_MusicPlay
SoundPlay.Visible = False
SoundOff.Visible = True
End Sub

'Or a system like:
' Picture1(0) - Play
' Picture1(1) - Pause
'Private Sub Picture1_Click(Index As Integer) 'Play/Pause the chiptune
'    Select Case Index
'        Case 0: Call BASSMOD_MusicPlay
'        Case 1: Call BASSMOD_MusicPause
'    End Select
'End Sub


' ----------------------------
' Support Routines
' ----------------------------

' Class method accepts byte arrays, so we will need to read the file,
' turn it into a byte array, and pass to the method.  The return value
' is numeric but we want to display it as text.  We will have to convert
' the return value before returning the result.
Private Function CRCFromFile(ByVal strFilePath As String) As String
    Dim bArrayFile() As Byte
    Dim lngCRC32 As Long

    Dim lngChunkSize As Long
    Dim lngSize As Long

    lngSize = FileLen(strFilePath)
    lngChunkSize = CHUNK_SIZE

    If lngSize <> 0 Then

        ' Read byte array from file
        Open strFilePath For Binary Access Read As #1

        Do While Seek(1) < lngSize

            If (lngSize - Seek(1)) > lngChunkSize Then
                ' Process data in chunks. Chunky!
                Do While Seek(1) < (lngSize - lngChunkSize)
                    ReDim bArrayFile(lngChunkSize - 1)
                    Get #1, , bArrayFile()
                    lngCRC32 = objCRC32.CRC32(lngCRC32, bArrayFile, lngChunkSize - 1)
                Loop
            Else
                ' Blast it at them
                ReDim bArrayFile(lngSize - Seek(1))
                Get #1, , bArrayFile()
                
                lngCRC32 = objCRC32.CRC32(lngCRC32, bArrayFile, UBound(bArrayFile))
            End If

        Loop

        Close #1

        ' Everyone expects to view checksums in Hex strings.  Add buffer zeros if
        ' needed by smaller values.
        CRCFromFile = Right$("00000000" & Hex$(lngCRC32), 8)
    Else
        ' File of zero bytes has a CRC of 0
        CRCFromFile = "00000000"
    End If
End Function

