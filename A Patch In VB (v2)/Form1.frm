VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "XYLITOL PROUDLY PRESENTS :"
   ClientHeight    =   5715
   ClientLeft      =   -1095
   ClientTop       =   4500
   ClientWidth     =   4590
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":08CA
   ScaleHeight     =   5715
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3360
      Top             =   6000
   End
   Begin VB.PictureBox Picture1 
      Height          =   6525
      Left            =   4800
      ScaleHeight     =   6465
      ScaleWidth      =   6615
      TabIndex        =   38
      Top             =   0
      Width           =   6680
      Begin VB.TextBox Text3 
         BackColor       =   &H00000000&
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
         Height          =   5775
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Text            =   "Form1.frx":0BD4
         Top             =   0
         Width           =   6615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   5880
         Width           =   6375
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   120
         Top             =   5880
         Width           =   6375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cool"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   6000
         Width           =   6375
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000005&
      Height          =   5715
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   0
      Width           =   4600
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4335
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "About"
            ForeColor       =   &H80000008&
            Height          =   2390
            Left            =   240
            TabIndex        =   23
            Top             =   2550
            Visible         =   0   'False
            Width           =   2655
            Begin VB.PictureBox Picture4 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H000000FF&
               Height          =   2115
               Left            =   120
               ScaleHeight     =   141
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   161
               TabIndex        =   24
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Status"
            ForeColor       =   &H80000008&
            Height          =   2390
            Left            =   240
            TabIndex        =   4
            Top             =   2550
            Width           =   2655
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   29
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "result:"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   18
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label20 
               BackColor       =   &H80000005&
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   17
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label19 
               BackColor       =   &H80000005&
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   16
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label18 
               BackColor       =   &H80000005&
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   15
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label17 
               BackColor       =   &H80000005&
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   14
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "access:"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               Caption         =   "write:"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               Caption         =   "bytes:"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               Caption         =   "filesize:"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               Caption         =   "filename:"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   600
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Info"
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   240
            TabIndex        =   3
            Top             =   1560
            Width           =   3855
            Begin VB.Label APPTARGET 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "xyli crkMe 5.exe"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1320
               TabIndex        =   36
               Top             =   220
               Width           =   2415
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "bytes"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   2760
               TabIndex        =   35
               Top             =   480
               Width           =   975
            End
            Begin VB.Label APPSIZE 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "510 976"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1320
               TabIndex        =   34
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Filesize:"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   120
               TabIndex        =   33
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Target:"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   120
               TabIndex        =   32
               Top             =   225
               Width           =   1095
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000005&
               X1              =   1320
               X2              =   0
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000005&
               X1              =   120
               X2              =   3720
               Y1              =   480
               Y2              =   480
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            Height          =   1095
            Left            =   360
            Picture         =   "Form1.frx":0BDC
            ScaleHeight     =   1035
            ScaleWidth      =   3555
            TabIndex        =   2
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "15-10-07"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   42
            Top             =   5040
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Cracked by Xylitol"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   5040
            Width           =   2655
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   31
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "NFO"
            Height          =   255
            Left            =   3000
            TabIndex        =   30
            Top             =   4150
            Width           =   1095
         End
         Begin VB.Shape Shape8 
            Height          =   375
            Left            =   3000
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label APPTITLE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "XylitolCrkMe 5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   26
            Top             =   3600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Left            =   3000
            TabIndex        =   25
            Top             =   3690
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Command3 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   22
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   21
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Command2 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   20
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Command1 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   19
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "About"
            Height          =   255
            Left            =   3000
            TabIndex        =   8
            Top             =   3690
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Restore"
            Height          =   255
            Left            =   3000
            TabIndex        =   7
            Top             =   3200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Crack"
            Height          =   255
            Left            =   3000
            TabIndex        =   6
            Top             =   2720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3000
            TabIndex        =   5
            Top             =   4640
            Width           =   1095
         End
         Begin VB.Shape Command4 
            Height          =   375
            Left            =   3000
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Shape Shape7 
            Height          =   375
            Left            =   3000
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Shape Shape6 
            Height          =   375
            Left            =   3000
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Shape Shape2 
            Height          =   375
            Left            =   3000
            Top             =   2640
            Width           =   1095
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Option Explicit

Dim Offst(0 To 1) As Long 'change
Dim DataP(0 To 1) As Byte 'change
Dim Fhand As Integer
Dim i As Integer
Dim p As Integer
Dim Answer As Integer
Dim Temp As String

'------------------------------------

Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nSize As Long, ByVal lpBuffer As String) As Long

Dim X As Long
Dim TempPfad As String
Dim SystemPfad As String

'-------------------
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
     Const DT_BOTTOM As Long = &H8
     Const DT_CALCRECT As Long = &H400
     Const DT_CENTER As Long = &H1
     Const DT_EXPANDTABS As Long = &H40
     Const DT_EXTERNALLEADING As Long = &H200
     Const DT_LEFT As Long = &H0
     Const DT_NOCLIP As Long = &H100
     Const DT_NOPREFIX As Long = &H800
     Const DT_RIGHT As Long = &H2
     Const DT_SINGLELINE As Long = &H20
     Const DT_TABSTOP As Long = &H80
     Const DT_TOP As Long = &H0
     Const DT_VCENTER As Long = &H4
     Const DT_WORDBREAK As Long = &H10
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'START EDIT CODE
Const ScrollText As String = "* * * * * * * * * * * * * * *" & _
                    vbCrLf & "XyliCrackMe 5 Patch" & _
                    vbCrLf & "* * * * * * * * * * * * * * *" & _
                    vbCrLf & " By Xylitol" & vbCrLf & _
                             "                      " & vbCrLf & _
                             "my respects fly out to:" & _
                    vbCrLf & "" & _
                    vbCrLf & "ByZiNc" & _
                    vbCrLf & "Maxtreme" & _
                    vbCrLf & "Niszczyciel" & _
                    vbCrLf & "Ypogeios" & _
                    vbCrLf & "Mr Fawkes" & _
                    vbCrLf & "gORDon_vdLg" & _
                    vbCrLf & "Sancho" & _
                    vbCrLf & "" & _
                    vbCrLf & "And" & _
                    vbCrLf & "All hardworking sceners" & _
                             ""
'STOP EDIT CODE

Dim EndingFlag As Boolean
'-----------------




Function FileExists(dzaFile As String) As Boolean
On Error Resume Next
Err.Clear
GetAttr dzaFile
FileExists = (Err = 0)

End Function

Private Sub APPTITLE_Click()

'START EDIT CODE
ShellExecute hwnd, "Open", "http://xylitol.free.fr/", "", App.Path, 1 'change
'STOP EDIT CODE

End Sub

Private Sub Command1_Click()
Label24.Visible = True
Label11.Visible = True
Label23.Visible = False
Label22.Visible = False
Frame3.Visible = True
Frame4.Visible = False

'START EDIT CODE
If Not FileExists("xyli crkMe 5.exe") Then   'first check to see if the target file exists
'STOP EDIT CODE

GoTo NOexist:
NOexist:
Label17.Caption = "ERROR"
Label18.Caption = "ERROR"
Label19.Caption = "ERROR"
Label25.Caption = "ERROR"
Label20.Caption = "ERROR"
Label21.Caption = "ABORTED"

'START EDIT CODE
Answer = MsgBox("      [xyli crkMe 5.exe] not found ! " & vbCrLf & _
              "                                          " & vbCrLf & _
              "        Make a Search  ?                 ", _
   vbQuestion + vbYesNo, "Search ?") 'change
   If Answer = vbYes Then
'STOP EDIT CODE

CommonDialog1.InitDir = "C:\"

'START EDIT CODE
CommonDialog1.DialogTitle = "Searching for xyli crkMe 5.exe" 'change
'STOP EDIT CODE

CommonDialog1.FileName = ""

'START EDIT CODE
CommonDialog1.Filter = "xyli crkMe 5.exe|*.exe;*.exe" 'change
'STOP EDIT CODE
CommonDialog1.ShowOpen
End If
If CommonDialog1.FileName = "" Then Exit Sub

'START EDIT CODE
If CommonDialog1.FileTitle <> "xyli crkMe 5.exe" Then 'change
'STOP EDIT CODE

Label25.Caption = "ERROR"
Label17.Caption = "ERROR"
Label18.Caption = "-"
Label19.Caption = "-"
Label20.Caption = "-"
Label21.Caption = "ABORTED"
MsgBox "Please Give the correct File To patch !", vbCritical, "Error"
Exit Sub
End If

'START EDIT CODE
If FileLen("xyli crkMe 5.exe") <> 510976 Then 'change
'STOP EDIT CODE

Label17.Caption = "OK"
Label25.Caption = "OK"
Label18.Caption = "ERROR"
Label19.Caption = "-"
Label20.Caption = "-"
Label21.Caption = "ABORTED"

'START EDIT CODE
  MsgBox "The patch is for xylitol's crackMe 5 only !!", vbCritical, "Error" 'change
'STOP EDIT CODE
Exit Sub
End If

Else: GoTo Exist:

Exist:
'then see if the size matches; if not, inform user and abort

'START EDIT CODE
If FileLen("xyli crkMe 5.exe") <> 510976 Then
'STOP EDIT CODE

Label25.Caption = "OK"
Label17.Caption = "OK"
Label18.Caption = "ERROR"
Label19.Caption = "-"
Label20.Caption = "-"
Label21.Caption = "ABORTED"

'START EDIT CODE
    MsgBox "The patch is for xylitol's crackMe 5 only !!", vbCritical, "Error"
'STOP EDIT CODE
   
Exit Sub
End If
End If
FileCopy App.Path & "/" & APPTARGET, App.Path & "/" & APPTARGET + ".bak"
Fhand = FreeFile

'          /!\
'/!\ START EDIT CODE /!\
'          /!\

Offst(0) = &H66D61
Offst(1) = &H66D62

'fill the array of data

DataP(0) = &H90
DataP(1) = &H90

Open "xyli crkMe 5.exe" For Binary As Fhand 'change

   For i = 0 To 1 'change
'STOP EDIT CODE

   Put Fhand, Offst(i) + 1, DataP(i)
   Next i

Close Fhand
        Label17.Caption = "OK"
        Label25.Caption = "OK"
        Label18.Caption = "OK"
        Label19.Caption = "OK"
        Label20.Caption = "OK"
        Label21.Caption = "CRACKED"
MsgBox "Patch Done successfuly !", vbOKOnly, "Patch Done !"
CommonDialog1.FileName = ""
End Sub



Private Sub Command2_Click()
        Label24.Visible = True
        Label11.Visible = True
        Label23.Visible = False
        Label22.Visible = False
Frame3.Visible = True
Frame4.Visible = False
If Not FileExists("xyli crkMe 5.exe") Then   'first check to see if the target file exists
GoTo NOexist:
NOexist:
        Label17.Caption = "ERROR"
        Label18.Caption = "ERROR"
        Label19.Caption = "ERROR"
        Label25.Caption = "ERROR"
        Label20.Caption = "ERROR"
        Label21.Caption = "ABORTED"
Answer = MsgBox("      [xyli crkMe 5.exe] not found ! " & vbCrLf & _
              "                                          " & vbCrLf & _
              "        Make a Search  ?                 ", _
   vbQuestion + vbYesNo, "Search?") 'change
   If Answer = vbYes Then

CommonDialog1.InitDir = "C:\Program Files"
CommonDialog1.DialogTitle = "Searching for xyli crkMe 5.exe" 'change
CommonDialog1.FileName = ""
CommonDialog1.Filter = "xyli crkMe 5.exe|*.exe;*.exe" 'change
CommonDialog1.ShowOpen
End If
If CommonDialog1.FileName = "" Then Exit Sub
If CommonDialog1.FileTitle <> "xyli crkMe 5.exe" Then 'change
        Label17.Caption = "ERROR"
        Label25.Caption = "ERROR"
        Label18.Caption = "-"
        Label19.Caption = "-"
        Label20.Caption = "-"
        Label21.Caption = "ABORTED"
MsgBox "Please Give the correct File To patch !", vbCritical, "Error"
Exit Sub
End If
If FileLen("xyli crkMe 5.exe") <> 510976 Then 'change
        Label17.Caption = "OK"
        Label25.Caption = "OK"
        Label18.Caption = "ERROR"
        Label19.Caption = "-"
        Label20.Caption = "-"
        Label21.Caption = "ABORTED"
  MsgBox "The patch is for xylitol's crackMe 5 only !!", vbCritical, "Error" 'change
Exit Sub
End If

Else: GoTo Exist:

Exist:
'then see if the size matches; if not, inform user and abort
If FileLen("xyli crkMe 5.exe") <> 510976 Then
          Label17.Caption = "OK"
          Label25.Caption = "OK"
          Label18.Caption = "ERROR"
          Label19.Caption = "-"
          Label20.Caption = "-"
          Label21.Caption = "ABORTED"
MsgBox "The patch is for xylitol's crackMe 5 only !!", vbCritical, "Error"
   
Exit Sub
End If
End If

              
Fhand = FreeFile

              Offst(0) = &H66D61 'offset to line
              Offst(1) = &H66D62 'offset to line

'fill the array of data

              DataP(0) = &H75 'code to change
              DataP(1) = &HC 'code to change

Open "xyli crkMe 5.exe" For Binary As Fhand 'change

   For i = 0 To 1 'exemple Offst(5) = For i = 0 To 5  ;-)
   Put Fhand, Offst(i) + 1, DataP(i)
   Next i

Close Fhand
          Label25.Caption = "OK" 'change text to: OK
          Label17.Caption = "OK" 'change text to: OK
          Label18.Caption = "OK" 'change text to: OK
          Label19.Caption = "OK" 'change text to: OK
          Label20.Caption = "OK" 'change text to: OK
          Label21.Caption = "RESTORED" 'change text to: RESTORED
    MsgBox "Patch Done successfuly !", vbOKOnly, "Patch Done !" 'code for show patch done
CommonDialog1.FileName = ""
End Sub



Private Sub Command3_Click()
uFMOD_PlaySong 0, 0, 0
End
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF& 'color = red
          Label1.ForeColor = &H80000008 'color = black
          Label22.ForeColor = &H80000008 'color = black
          Label23.ForeColor = &H80000008 'color = black
          Label2.ForeColor = &H80000008 'color = black
          Label24.ForeColor = &H80000008 'color = black
          Label26.ForeColor = &H80000008 'color = black
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &H80000008 'color = black
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80000008 'color = black
          Label1.ForeColor = &H80000008 'color = black
          Label26.ForeColor = &H80000008 'color = black
          Label22.ForeColor = &H80000008 'color = black
          Label3.ForeColor = &H80000008 'color = black
          Label11.ForeColor = &H80000008 'color = black
End Sub

Private Sub Label10_Click()
End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label22.ForeColor = &HFF& 'color = red
          Label1.ForeColor = &H80000008 'color = black
          Label12.ForeColor = &H80000008 'color = black
          Label24.ForeColor = &H80000008 'color = black
          Label3.ForeColor = &H80000008 'color = black
          Label2.ForeColor = &H80000008 'color = black
          Label26.ForeColor = &H80000008 'color = black
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &HFF& 'color = red
          Label1.ForeColor = &H80000008 'color = black
          Label12.ForeColor = &H80000008 'color = black
          Label22.ForeColor = &H80000008 'color = black
          Label23.ForeColor = &H80000008 'color = black
          Label3.ForeColor = &H80000008 'color = black
          Label2.ForeColor = &H80000008 'color = black
          Label26.ForeColor = &H80000008 'color = black
End Sub

Private Sub Label27_Click()
Form1.Width = 6735 ' New dim
Form1.Height = 6915 ' New dim
Form1.Caption = "NFO" ' change  title of this patch to NFO
    Picture1.Width = 6680 ' New dim
    Picture1.Height = 6525 ' New dim
    Picture1.Top = 0 ' New dim
    Picture1.Left = 0 ' New dim
Timer1.Enabled = True 'timer1 = actived
End Sub

Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label26.ForeColor = &HFF& 'color = red
          Label1.ForeColor = &H80000008 'color = black
          Label12.ForeColor = &H80000008 'color = black
          Label22.ForeColor = &H80000008 'color = black
          Label3.ForeColor = &H80000008 'color = black
          Label11.ForeColor = &H80000008 'color = black
          Label2.ForeColor = &H80000008 'color = black
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF& 'color = red
          Label2.ForeColor = &H80000008 'color = black
          Label26.ForeColor = &H80000008 'color = black
          Label22.ForeColor = &H80000008 'color = black
          Label3.ForeColor = &H80000008 'color = black
          Label11.ForeColor = &H80000008 'color = black
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF& 'color = red
          Label1.ForeColor = &H80000008 'color = black
          Label26.ForeColor = &H80000008 'color = black
          Label22.ForeColor = &H80000008 'color = black
          Label3.ForeColor = &H80000008 'color = black
          Label11.ForeColor = &H80000008 'color = black
End Sub


Private Sub Form_Load()

If App.PrevInstance = True Then
    MsgBox "Patch is already running...", vbExclamation, "Operation : Terminated"
    End
    
End If
Picture4.ForeColor = &H0& 'color = red
Picture4.FontSize = 11 'size police = 11
Picture4.Font = "Arial" 'police type = arial


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub Label23_Click()
Label24.Visible = True
Label11.Visible = True
Label23.Visible = False
Label22.Visible = False
Frame4.Visible = False
End Sub

Private Sub Label24_Click()
Label24.Visible = False
Label11.Visible = False
Label23.Visible = True
Label22.Visible = True
Frame4.Visible = True
RunMain
End Sub



Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 30
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long
Dim RectHeight As Long
Me.Refresh
rt = DrawText(Picture4.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)
If rt = 0 Then
    MsgBox "Err0r: Impossible to Scroll", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = Picture4.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = Picture4.ScaleWidth

    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + Picture4.ScaleHeight
End If

Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                   
        Picture4.Cls
        
        DrawText Picture4.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
    
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        

        If DrawingRect.Top < -(RectHeight) Then
            DrawingRect.Top = Picture4.ScaleHeight ' code for scroll
            DrawingRect.Bottom = RectHeight + Picture4.ScaleHeight
        End If
        
        Picture4.Refresh 'refresh the piture4
        
        LastFrameTime = GetTickCount()
        
    End If
    
    DoEvents
Loop ' continu
Unload Me ' unload
Set Form1 = Nothing
End Sub



Private Sub Label9_Click()
Form1.Height = 6120 ' new dim
Form1.Width = 4680 ' new dim
Form1.Caption = "XYLITOL PROUDLY PRESENTS :" ' change  title of this patch to XYLITOL PROUDLY PRESENTS :
Picture1.Width = 6680 ' new dim
Picture1.Height = 6525 ' new dim
Picture1.Left = 4800 ' new dim
Timer1.Enabled = True ' Activate timer1
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF& 'Red color
End Sub


Public Sub centerform(frm As Form)
'Code for center to screen
frm.Top = Screen.Height / 2 - frm.Height / 2
frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub


Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &H80000008 'color = black
End Sub

Private Sub Timer1_Timer()
centerform Me 'code for re center the form
Timer1.Enabled = False 'code for desactivate the timer
End Sub
