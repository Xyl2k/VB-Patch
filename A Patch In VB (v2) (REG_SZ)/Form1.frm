VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "XYLITOL PROUDLY PRESENTS :"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Height          =   1935
      Left            =   4800
      ScaleHeight     =   1875
      ScaleWidth      =   3555
      TabIndex        =   11
      Top             =   0
      Width           =   3615
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get Current Values"
         Height          =   255
         Left            =   1320
         TabIndex        =   38
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "email:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "key:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000005&
      Height          =   5715
      Left            =   5280
      ScaleHeight     =   5655
      ScaleWidth      =   4545
      TabIndex        =   2
      Top             =   7800
      Width           =   4600
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   4335
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Left            =   3000
            TabIndex        =   9
            Top             =   3690
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   8
            Top             =   3600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Shape Shape8 
            Height          =   375
            Left            =   3000
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "NFO"
            Height          =   255
            Left            =   3000
            TabIndex        =   7
            Top             =   4150
            Width           =   1095
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   6
            Top             =   4080
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5595
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   0
      Width           =   4620
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4335
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            Caption         =   "About"
            ForeColor       =   &H80000008&
            Height          =   2025
            Left            =   240
            TabIndex        =   39
            Top             =   2670
            Visible         =   0   'False
            Width           =   2655
            Begin VB.PictureBox Picture4 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   1635
               Left            =   120
               ScaleHeight     =   109
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   161
               TabIndex        =   40
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            Caption         =   "Status"
            ForeColor       =   &H80000008&
            Height          =   2025
            Left            =   240
            TabIndex        =   21
            Top             =   2670
            Width           =   2655
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   29
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "result:"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   27
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label18 
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   26
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label17 
               Caption         =   "-"
               Height          =   255
               Left            =   1320
               TabIndex        =   25
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "access:"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "write:"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "filename:"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            Caption         =   "Info"
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   240
            TabIndex        =   16
            Top             =   1680
            Width           =   3855
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Filesize:"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Target:"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   120
               TabIndex        =   19
               Top             =   225
               Width           =   1095
            End
            Begin VB.Label APPTARGET 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Register"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1320
               TabIndex        =   20
               Top             =   220
               Width           =   2415
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "bytes"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   2760
               TabIndex        =   18
               Top             =   480
               Width           =   975
            End
            Begin VB.Label APPSIZE 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "n/a"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1320
               TabIndex        =   17
               Top             =   480
               Width           =   1455
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
            Index           =   1
            Left            =   360
            Picture         =   "Form1.frx":0000
            ScaleHeight     =   1035
            ScaleWidth      =   3555
            TabIndex        =   15
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   47
            Top             =   3360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   42
            Top             =   3840
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Left            =   3000
            TabIndex        =   41
            Top             =   3920
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   4
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   5
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   10
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   32
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
            Height          =   255
            Left            =   3000
            TabIndex        =   34
            Top             =   4400
            Width           =   1095
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "About"
            Height          =   255
            Left            =   3000
            TabIndex        =   37
            Top             =   3920
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Unpatch"
            Height          =   255
            Left            =   3000
            TabIndex        =   36
            Top             =   3450
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Patch"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3000
            TabIndex        =   35
            Top             =   2850
            Width           =   1095
         End
         Begin VB.Shape Shape4 
            Height          =   375
            Left            =   3000
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Shape Shape3 
            Height          =   375
            Left            =   3000
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Shape Shape2 
            Height          =   375
            Left            =   3000
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Shape Shape1 
            Height          =   375
            Left            =   3000
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "08-12-07"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   31
            Top             =   4800
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Cracked by Xylitol"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   4800
            Width           =   2655
         End
         Begin VB.Label APPTITLE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "XyliRegMe v1"
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
            TabIndex        =   14
            Top             =   0
            Width           =   4095
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim c As New cRegistry

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
                    vbCrLf & "XyliRegMe 1 Patch" & _
                    vbCrLf & "* * * * * * * * * * * * * * *" & _
                    vbCrLf & " By Xylitol" & vbCrLf & _
                             "                      " & vbCrLf & _
                             "my respects fly out to:" & _
                    vbCrLf & "" & _
                    vbCrLf & "ByZiNc" & _
                    vbCrLf & "Maxtreme" & _
                    vbCrLf & "Niszczyciel" & _
                    vbCrLf & "BytePlayer" & _
                    vbCrLf & "FLeXuS_GReeN" & _
                    vbCrLf & "gORDon_vdLg" & _
                    vbCrLf & "Sancho" & _
                    vbCrLf & "" & _
                    vbCrLf & "And" & _
                    vbCrLf & "All hardworking sceners" & _
                             ""
'STOP EDIT CODE

Dim EndingFlag As Boolean
'-----------------


Private Sub APPTITLE_Click()

'START EDIT CODE
ShellExecute hwnd, "Open", "http://google.fr/", "", App.Path, 1 'change
'STOP EDIT CODE

End Sub

Private Sub Command1_Click()
With c
'START EDIT CODE
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\XyliRegMe" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "Mail" 'nom d'un key (lol)
         Text2.Text = .Value 'attribuer la valeur du clé au texte (affichage)
        .ValueKey = "unlockKey"  '//
         Text3.Text = .Value
        .ValueKey = "Name"   '//
         Text1.Text = .Value
'STOP EDIT CODE
         End With
         If Text3.Text = "" And Text2.Text = "" Then
         Text3.Text = "Unregistred..." 'A bon ?
         Text2.Text = "Unregistred..." 'A bon ?
         Text1.Text = "Unregistred..." 'A bon ?
         End If
End Sub





Private Sub Form_Load()
        If uFMOD_PlaySong(1, 0, XM_RESOURCE) <> 0 Then
        End If
        With c
'START EDIT CODE
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\XyliRegMe" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "Mail"
         Text2.Text = .Value 'attribuer la valeur du clé au texte (affichage)
        .ValueKey = "unlockKey"  '//
         Text3.Text = .Value
        .ValueKey = "Name"   '//
         Text1.Text = .Value
'STOP EDIT CODE
         End With
         If Text3.Text = "" And Text2.Text = "" Then
         Text3.Text = "Unregistred..." 'A bon ?
         Text2.Text = "Unregistred..." 'A bon ?
         Text1.Text = "Unregistred..." 'A bon ?
         End If
End Sub

Private Sub Form_Terminate()
End 'close this app
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = &H80000012 'Black color
Label11.ForeColor = &H80000012 'Black color
Label12.ForeColor = &H80000012 'Black color
Label13.ForeColor = &H80000012 'Black color
Label10.ForeColor = &H80000012 'Black color
Label29.ForeColor = &H80000012 'Black color
End Sub



Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = &H80000012 'Black color
Label11.ForeColor = &H80000012 'Black color
Label12.ForeColor = &H80000012 'Black color
Label13.ForeColor = &H80000012 'Black color
Label10.ForeColor = &H80000012 'Black color
Label29.ForeColor = &H80000012 'Black color
End Sub

Private Sub Label19_Click()
Form2.Show 'parametre pour changer de fenetre et por désactivé le bazrd sur la premiere
Form1.Enabled = False
Frame4.Visible = False
Label29.Visible = False
Label30.Visible = False
Label24.Visible = True
Label13.Visible = True
Label25.Caption = "OK"
Label31.Visible = False
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFF& 'Red color
Label11.ForeColor = &H80000012 'Black color
Label12.ForeColor = &H80000012 'Black color
Label13.ForeColor = &H80000012 'Black color
Label14.ForeColor = &H80000012 'Black color
Label29.ForeColor = &H80000012 'Black color
End Sub

Private Sub Label20_Click()
If Text2.Text = "Unregistred..." Then 'email
Label25.Caption = "OK"
Label18.Caption = "OK"
Label17.Caption = "BAD"
Label21.Caption = "ALREADY UNPATCHED"
Label31.Visible = True
Else

With c
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\XyliRegMe"
        .DeleteKey   'Suppression de la clé en haut
         End With
         
Frame4.Visible = False
Label29.Visible = False
Label30.Visible = False
Label24.Visible = True
Label13.Visible = True
Label25.Caption = "OK"
Label18.Caption = "OK"
Label17.Caption = "OK"
Label21.Caption = "UNPATCHED"
Label31.Visible = True
         Text3.Text = "Unregistred..." 'A bon ?
         Text2.Text = "Unregistred..." 'A bon ?
         Text1.Text = "Unregistred..." 'A bon ?
End If
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &HFF& 'Red color
Label10.ForeColor = &H80000012 'Black color
Label12.ForeColor = &H80000012 'Black color
Label13.ForeColor = &H80000012 'Black color
Label14.ForeColor = &H80000012 'Black color
Label29.ForeColor = &H80000012 'Black color
End Sub


Private Sub Label24_Click()
Frame4.Visible = True
Label24.Visible = False
Label13.Visible = False
Label29.Visible = True
Label30.Visible = True
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

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = &HFF& 'Red color
Label11.ForeColor = &H80000012 'Black color
Label12.ForeColor = &H80000012 'Black color
Label10.ForeColor = &H80000012 'Black color
Label29.ForeColor = &H80000012 'Black color
Label14.ForeColor = &H80000012 'Black color
End Sub

Private Sub Label28_Click()
End 'close this app
End Sub

Private Sub Label28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = &HFF& 'Red color
Label11.ForeColor = &H80000012 'Black color
Label12.ForeColor = &H80000012 'Black color
Label13.ForeColor = &H80000012 'Black color
Label10.ForeColor = &H80000012 'Black color
Label29.ForeColor = &H80000012 'Black color
End Sub

Private Sub Label30_Click()
Label29.Visible = False
Label30.Visible = False
Label24.Visible = True
Label13.Visible = True
Frame4.Visible = False
End Sub

Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label29.ForeColor = &HFF& 'Red color
Label14.ForeColor = &H80000012 'Black color
Label11.ForeColor = &H80000012 'Black color
Label12.ForeColor = &H80000012 'Black color
Label13.ForeColor = &H80000012 'Black color
Label10.ForeColor = &H80000012 'Black color
End Sub

Private Sub Label31_Click()
Label25.Caption = "OK"
Label18.Caption = "OK"
Label17.Caption = "BAD"
Label21.Caption = "ALREADY UNPATCHED" 'dommage !
End Sub

Private Sub Label31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &HFF& 'Red color
Label10.ForeColor = &H80000012 'Black color
Label12.ForeColor = &H80000012 'Black color
Label13.ForeColor = &H80000012 'Black color
Label14.ForeColor = &H80000012 'Black color
Label29.ForeColor = &H80000012 'Black color
End Sub

