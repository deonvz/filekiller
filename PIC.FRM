VERSION 5.00
Begin VB.Form pic 
   BackColor       =   &H00000000&
   Caption         =   "File Killer"
   ClientHeight    =   3990
   ClientLeft      =   270
   ClientTop       =   1995
   ClientWidth     =   4455
   ClipControls    =   0   'False
   Icon            =   "pic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2520
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      Picture         =   "pic.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      Picture         =   "pic.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   720
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "-=<File Killer>=-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mnuright 
      Caption         =   "Right"
      Begin VB.Menu mnufull 
         Caption         =   "FullScreen"
      End
      Begin VB.Menu mnunormal 
         Caption         =   "Normal"
      End
      Begin VB.Menu mnumouse 
         Caption         =   "Hide Mouse Cursor"
      End
      Begin VB.Menu mnuslide 
         Caption         =   "Slide Show"
         Begin VB.Menu mnu2 
            Caption         =   "2 Sec"
         End
         Begin VB.Menu mnu4 
            Caption         =   "4 Sec"
         End
         Begin VB.Menu mnu6 
            Caption         =   "6 Sec"
         End
         Begin VB.Menu mnu8 
            Caption         =   "8 Sec"
         End
         Begin VB.Menu mnuown 
            Caption         =   "Own Timing"
         End
         Begin VB.Menu mnuline2 
            Caption         =   "-"
         End
         Begin VB.Menu mnustop 
            Caption         =   "Stop"
         End
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim c As String
Dim b As Integer

If pic.WindowState = 2 Then
    pic.WindowState = 0
End If

If Form1.File1.ListIndex = 0 Then
    Form1.File1.ListIndex = Form1.File1.ListCount - 1
Else
    Form1.File1.ListIndex = Form1.File1.ListIndex - 1
End If

    c = Form1.Dir1.Path & "\" & Form1.File1.FileName
    Form1.Text3.Text = Form1.Dir1.Path & "\" & Form1.File1.FileName
    Form1.File1.Pattern = Form1.Text1.Text

If Form1.File1.Pattern = Form1.Text1.Text Then
    Module1.show
     Form1.Image1.Picture = LoadPicture(c)
     Form1.Image2.Picture = Form1.Image1.Picture
     Form1.Label7.Caption = "Width=" & Form1.Image2.Width
     Form1.Label8.Caption = "Height=" & Form1.Image2.Height
     Form1.Label4.Caption = "Current File : " & (Form1.File1.ListIndex + 1) & "/" & Form1.File1.ListCount
 Else
End If

    Image1.Picture = Form1.Image2.Picture
    Image1.Width = Form1.Image2.Width
    Image1.Height = Form1.Image2.Height
    pic.Height = Form1.Image2.Height + 1000
    pic.Width = Form1.Image2.Width + 100
    Command1.SetFocus

If Form1.Image2.Width >= 10000 Then
    pic.Left = 0
    pic.Top = 0
ElseIf Form1.Image2.Height >= 7650 Then
    pic.Top = 0
Else
pic.WindowState = 0
End If

pic.Caption = "File Killer :" & Form1.Label4.Caption & " Filesize: " & Form1.Label11.Caption

End Sub

Private Sub Command2_Click()
Dim c As String
Dim b As Integer

If pic.WindowState = 2 Then
pic.WindowState = 0
End If

If Form1.File1.ListIndex = Form1.File1.ListCount - 1 Then
    Form1.File1.ListIndex = 0
Else
    Form1.File1.ListIndex = Form1.File1.ListIndex + 1
End If

    c = Form1.Dir1.Path & "\" & Form1.File1.FileName
    Form1.Text3.Text = Form1.Dir1.Path & "\" & Form1.File1.FileName
    Form1.File1.Pattern = Form1.Text1.Text
    
If Form1.File1.Pattern = Form1.Text1.Text Then
    Module1.show
     Form1.Image1.Picture = LoadPicture(c)
     Form1.Image2.Picture = Form1.Image1.Picture
     Form1.Label7.Caption = "Width=" & Form1.Image2.Width
     Form1.Label8.Caption = "Height=" & Form1.Image2.Height
     Form1.Label4.Caption = "Current File : " & (Form1.File1.ListIndex + 1) & "/" & Form1.File1.ListCount
 Else
End If

    Image1.Picture = Form1.Image2.Picture
    Image1.Width = Form1.Image2.Width
    Image1.Height = Form1.Image2.Height
    pic.Height = Form1.Image2.Height
    pic.Width = Form1.Image2.Width

    Command2.SetFocus
    
If Image1.Width >= 10000 Then
    pic.Left = 0
    pic.Top = 0
ElseIf Image1.Height >= 7650 Then
    pic.Top = 0
Else
    pic.WindowState = 0
End If

pic.Caption = "File Killer :" & Form1.Label4.Caption & " Filesize: " & Form1.Label11.Caption
 

End Sub


Private Sub Command3_Click()
    Kill (Form1.Dir1.Path + "\" + Form1.File1.FileName)
    Form1.File1.Refresh
    Command2.SetFocus
    
    CurrentX = Command2.Left
    CurrentY = Command2.Top
    
End Sub

Private Sub Form_Load()

    Image1.Stretch = False
    Label9.Caption = Form1.File1.FileName
    Label1.Caption = Trim(Form1.Image2.Width) & "x" & Trim(Form1.Image2.Height)
    Image1.Height = Form1.Image2.Height
    mnuright.Visible = False
    '''''''''''''''
    pic.Height = Form1.Image2.Height + 1000
    pic.Width = Form1.Image2.Width + 100
    ''''''''''''''''''
If Image1.Width >= 10200 Then
    pic.Left = 0
    pic.Top = 0
ElseIf Image1.Height >= 7680 Then
    pic.Top = 0
Else
End If

End Sub



Private Sub Form_Resize()
    
    If Form1.WindowState = 2 Then
    Else
    Image1.Stretch = True
    Label9.Caption = Form1.File1.FileName
    Label1.Caption = Trim(Form1.Image2.Width) & "x" & Trim(Form1.Image2.Height)
    
    Image1.Width = pic.ScaleWidth
    Image1.Height = pic.ScaleHeight
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Form1.Visible = True

End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
        PopupMenu mnuright
    End If
    
End Sub

Private Sub mnu2_Click()

Timer1.Interval = 2000
Timer1.Enabled = True
mnu2.Checked = True
mnu4.Checked = False
mnu6.Checked = False
mnu8.Checked = False
mnuown.Checked = False

End Sub

Private Sub mnu4_Click()

Timer1.Interval = 4000
Timer1.Enabled = True
mnu2.Checked = False
mnu4.Checked = True
mnu6.Checked = False
mnu8.Checked = False
mnuown.Checked = False

End Sub

Private Sub mnu6_Click()

Timer1.Interval = 6000
Timer1.Enabled = True
mnu2.Checked = False
mnu4.Checked = False
mnu6.Checked = True
mnu8.Checked = False
mnuown.Checked = False

End Sub

Private Sub mnu8_Click()

Timer1.Interval = 8000
Timer1.Enabled = True
mnu2.Checked = False
mnu4.Checked = False
mnu6.Checked = False
mnu8.Checked = True
mnuown.Checked = False

End Sub

Private Sub mnudelete_Click()

    Kill (Form1.Dir1.Path + "\" + Form1.File1.FileName)
    Form1.File1.Refresh
    
End Sub

Private Sub mnufull_Click()

    
    pic.WindowState = 2
   ' Image1.Stretch = True
   ' Image1.Width = pic.Width - 100
   ' Image1.Height = pic.Height - 1000
    
    'Take this out if it is making A fuck up =(
    Image1.Height = Form1.Image2.Height
    Image1.Width = Form1.Image2.Width
    pic.BorderStyle = 0
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Shape1.Visible = False
    Label1.Visible = False
    Label9.Visible = False
    
    
End Sub

Private Sub mnumouse_Click()

If mnumouse.Checked = True Then
mnumouse.Checked = False
Else
mnumouse.Checked = True
End If

End Sub

Private Sub mnunormal_Click()

    Image1.Stretch = False
    'pic.WindowState = 0
    
    'Take this shit out is it messes up =(
    pic.BorderStyle = 2
        Command1.Visible = True
    Command2.Visible = True
    Command3.Visible = True
    Shape1.Visible = True
    Label1.Visible = True
    Label9.Visible = True
     pic.WindowState = 0
     pic.Height = Form1.Image2.Height
    pic.Width = Form1.Image2.Width
     Image1.Height = Form1.Image2.Height
    Image1.Width = Form1.Image2.Width
    
    If Image1.Width >= 10000 Then
    pic.Left = 0
    pic.Top = 0
ElseIf Image1.Height >= 7650 Then
    pic.Top = 0
    End If
    
End Sub

Private Sub mnuown_Click()

Dialog.show
mnu2.Checked = False
mnu4.Checked = False
mnu6.Checked = False
mnu8.Checked = False
mnuown.Checked = True

End Sub

Private Sub mnustop_Click()

Timer1.Enabled = False
mnu2.Checked = False
mnu4.Checked = False
mnu6.Checked = False
mnu8.Checked = False
mnuown.Checked = False

End Sub

Private Sub Timer1_Timer()

Module1.slide

End Sub
