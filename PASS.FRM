VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Killer"
   ClientHeight    =   1410
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "PASS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Enable File Killer"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "-=<File Killer>=-"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Label9.FontSize = 14
If Text1.Text = "password" Then
Form1.Show
Unload Me
Label9.Caption = "Password is valid..."
Else
Label9.Caption = "Try Again"
Text1.Text = ""
End If

End Sub
Private Sub Label9_DblClick()
    
    Form1.Show
Unload Me
    
End Sub

Private Sub Text1_Change()

If Text1.Text = "password" Then
Label9.FontSize = 10
Label9.Caption = "Password is valid..."
Command1.SetFocus
Else
End If

End Sub

Private Sub Text1_Click()
Label9.FontSize = 14
Label9.Caption = "-=<File Killer>=-"
End Sub
