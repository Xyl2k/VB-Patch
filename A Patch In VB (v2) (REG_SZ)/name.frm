VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patch"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   570
         Width           =   735
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   2760
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Enter your name please:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New cRegistry

Private Sub Label3_Click()
If Text1.Text = "" Then
MsgBox "Enter a name please", 48 + 0, "Hey you !" 'Alert message
Else
Form1.Label18.Caption = "OK"
Form1.Label17.Caption = "OK"
Form1.Label21.Caption = "PATCHED"
With c
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\XyliRegMe" 'the directory to the key
        .ValueType = REG_SZ 'you can choose REG_DWORD if you want to add numbers only
        .ValueKey = "Mail" 'the first key name
        .Value = "" & Text1 & "@fbi.gov" 'the above key value(serialNum)
        .ValueKey = "unlockKey" 'the second key name
        .Value = "A4D8-S7D9-A7X1-SQ7X" 'password
        .ValueKey = "Name" 'the third key name
        .Value = Text1.Text
End With

With c
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\XyliRegMe" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "Mail"
         Form1.Text2.Text = .Value 'attribuer la valeur du clé au texte (affichage)
        .ValueKey = "unlockKey"  '//
         Form1.Text3.Text = .Value
        .ValueKey = "Name"   '//
         Form1.Text1.Text = .Value
         End With
         If Form1.Text3.Text = "" And Form1.Text2.Text = "" Then
         Form1.Text3.Text = "Unregistred..." 'A bon ?
         Form1.Text2.Text = "Unregistred..." 'A bon ?
         Form1.Text1.Text = "Unregistred..." 'A bon ?
         End If
         
Form1.Enabled = True 'parametre pour changer de fenetre et por désactivé le bazrd sur la premiere
Unload Me 'dégage
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF& 'Red color
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80000012 'Black color
End Sub
