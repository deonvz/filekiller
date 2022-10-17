VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Killer"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5280
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FF8080&
   Icon            =   "file killer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox Check6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Delete by File Height and Width"
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2280
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   5055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "OS"
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   2280
      TabIndex        =   20
      Top             =   3120
      Width           =   2895
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Win NT"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Win 95/98"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "Enjoy this proggie !!! "
      Top             =   3840
      Width           =   5055
   End
   Begin VB.CheckBox Check5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "<"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   ">"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Quit"
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   7680
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Text            =   "5"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Delete files by Filesize"
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "NOt Working Yet , I stress Yet!!.Delete files in the selected extention that are not equal to the following file size in Kb"
      Top             =   1260
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Text            =   "*.jpg"
      Top             =   600
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Delete all but these files"
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Delete all files in the directory except the following files . "
      Top             =   300
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "NO Preview Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   7177
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   29
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Designed by Deon van Zyl"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   6960
      UseMnemonic     =   0   'False
      Width           =   2295
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
      Left            =   2400
      TabIndex        =   27
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "W"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "H"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Info :"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2400
      TabIndex        =   22
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2895
      Left            =   2280
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "File size in Kb"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "File extention"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   4080
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   120
      MouseIcon       =   "file killer.frx":0442
      MousePointer    =   14  'Arrow and Question
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Menu mnuright 
      Caption         =   "Right Click"
      Begin VB.Menu mnuDlete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filesize
Private Sub Check1_Click()
    If Check1.Value = 1 Then
    Text1.Enabled = True
    File1.Enabled = True
    File1.Pattern = Text1.Text
     ElseIf Check1.Value = 0 Then
    Text1.Enabled = False
    File1.Enabled = False
    File1.Pattern = "*.*"
    End If
    
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
    Text2.Enabled = True
    Check3.Enabled = True
    Check4.Enabled = True
    Check5.Enabled = True
    ElseIf Check1.Value = 0 Then
    Text2.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    Check5.Enabled = False
    End If
End Sub

Private Sub Check6_Click()
    If Check6.Value = 1 Then
    File1.Enabled = True
    Text4.Enabled = True
    Label7.Visible = True
    Label8.Visible = True
    File1.Pattern = Text1.Text
        Text5.Enabled = True
        Else
        File1.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Label7.Visible = False
    Label8.Visible = False
    File1.Pattern = "*.*"
    End If
End Sub

Private Sub Command1_Click()
    Dim a As String
    Dim b As Integer
    Dim c As String
    Dim d As Integer
    Dim e As String
    Dim f As String
    File1.Enabled = False
    Label4.Caption = "Busy!!!"
    If Option1.Value = True Then
    a = "c:\windows\temp"
    ElseIf Option2.Value = True Then
    a = "c:\temp"
    End If
     e = Dir1.Path
     f = File1.List(b)
     MousePointer = 11
     
     
    
     
     
''''''''''''''''''''''''Running to make file sort go in to Filelistbox'''''
    ' Initialize for search, then perform recursive search.
    Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
    Dim result As Integer
      ' Check what the user did last.
        ' Update dir1.Path if it is different from the currently
        ' selected directory, otherwise perform the search.
        If Dir1.Path <> Dir1.List(Dir1.ListIndex) Then
            Dir1.Path = Dir1.List(Dir1.ListIndex)
            Exit Sub         ' Exit so user can take a look before searching.
        End If
    
        ' Continue with the search.
        File1.Pattern = Text1.Text
        FirstPath = Dir1.Path
        DirCount = Dir1.ListCount
        'Start recursive direcory search.
        NumFiles = 0                       ' Reset found files indicator.
        File1.Path = Dir1.Path
        MousePointer = 0
        '''''''''''''''''''''''''''''''''''''''End of Test'''''''''''''
        If Check1.Value = 0 Then
        File1.Pattern = "*.*"
        Label4.Caption = "Files deleted by size not File Sort , Done!!"
        End If
        
        MousePointer = 0
        Label3.Caption = "Files Found:" & File1.ListCount
        
'''''''''''''''''''''''Finished running all the commands of show so that a module isn`t needed'''''''''
'''''''''''''''''''''''''''''Cool Stuff 2 Know (do not Delete)'''''''''''''
'MkDir "c:\woop"
'RmDir "c:\woop"
'kill "c:\cleaner.ack"
'FileCopy "c:\cleaner.bat", "d:\cleaner.bat"
'filemove c:\games.jpg"
'Shell "c:\windows\notepad.exe"
'File1.Pattern = Text1.Text the file1.pattern ,pattern means what sort of files (like *.avi) must be displayed in the Filelistbox
'''''''''''''''''''''''''''''''End of cool stuff'''''''''''''''
''''''''''''''''This must copy the selected files to temp and move it back after everything has been deleted in the Dir
    b = 0
    d = 0
 
    If Check1.Value = 1 Then 'moving Files to temp Dir
    Do Until d = File1.ListCount
    c = Dir1.Path & "\" & File1.List(b)
    e = Dir1.Path

    FileCopy c, a & "\" & File1.List(b)
    
    Kill c
    b = b + 1
    d = d + 1
    MousePointer = 11
    Label4.Caption = "Busy with " & b & "\" & File1.ListCount
    Loop
         Label4.Caption = "Moving to temp Done..."
    Kill e & "\*.*"
    Label4.Caption = "Deleting all unwanted files in Directory"
    End If
    
    If d = File1.ListCount Then 'Moving files back from temp Dir
    d = 0
     b = 0
    File1.Pattern = Text1.Text
    Do Until d = File1.ListCount
    Dir1.Path = a
    File1.Pattern = Text1.Text
    c = Dir1.Path & "\" & File1.List(b)
    FileCopy c, e & "\" & File1.List(b)
    Kill c
    b = b + 1
    d = d + 1
    MousePointer = 11
    Label4.Caption = "Busy with " & b & "\" & File1.ListCount
       Loop
       Dir1.Path = e
    Label4.Caption = "Moving to source Directory done..."
    MousePointer = 0
     Label4.Caption = "Files deleted by File Sort , Done!!"
    End If
    
    
 If Check2.Value = 1 Then 'Deleting Files by size
 b = 0
 d = 0
 Dim kb As Integer
 Dim c1
 
 'Do Until d = File1.ListCount
c1 = Dir1.Path & "\" & File1.List(b)
kb = Val(Text2.Text)
    If Val(Label11.Caption) < kb Then
    Kill c1
    
        b = b + 1
         d = d + 1
         MousePointer = 11
        
       ElseIf Val(Label11.Caption) >= kb Then
    b = b + 1
      d = d + 1
      
     MousePointer = 11
  
     MousePointer = 0
     Label4.Caption = "Files deleted by Size, Done!!"
    Else
    MousePointer = 0
   'Loop
 End If
     End If
     
     
     If Check6.Value = 1 Then 'Deleting Files by Height and width
     b = 0
     d = 0
     File1.Pattern = Text1.Text
    Do Until d = File1.ListCount
   
    File1.Pattern = Text1.Text
    c = Dir1.Path & "\" & File1.List(b)
    Image2.Picture = LoadPicture(c)
    If Image2.Width <= Val(Text4.Text) Or Image2.Height <= Val(Text5.Text) Then
         Kill (e & "\" & File1.List(b))
        End If
    b = b + 1
    d = d + 1
    Loop
    Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.List(0))
     File1.ListIndex = 0
       Label4.Caption = "Files deleted by height and width!!!"
     End If
     File1.Refresh
     Image1.Picture = LoadPicture("")
     Image2.Picture = LoadPicture("")
     Beep
     File1.Enabled = True
     Form1.Refresh
End Sub

Private Sub Command2_Click()

    End
    
End Sub

Private Sub Command3_Click()

    Module1.show

    
End Sub



Private Sub Command4_Click()

Debug.Print Image1.Picture

End Sub

Private Sub Dir1_Change()
    
    Text3.Text = Dir1.Path
    File1.Path = Dir1.Path
    File1.Pattern = "*.*"
    Label3.Caption = ""
    Label4.Caption = ""
End Sub

Private Sub Drive1_Change()
   
    On Error GoTo DriveError
     Dir1.Path = Drive1.Drive
     Exit Sub
     
DriveError:
     MsgBox "Device Not Ready!", vbExclamation, "Error"
     Drive1.Drive = Dir1.Path
     Exit Sub
     
     
End Sub



Private Sub File1_Click()

Dim c As String
Dim b As Integer

     c = Dir1.Path & "\" & File1.FileName
    Text3.Text = Dir1.Path & "\" & File1.FileName
      File1.Pattern = Text1.Text


    'Get file size of file
  
  filesize = Str(FileLen(c))
  Label4.Caption = "Current File : " & (File1.ListIndex + 1) & "/" & File1.ListCount
    
If filesize >= 1000000 Then
    filesize = Val(filesize) / 1000000
    Label11.Caption = "Filesize:" & Int(filesize) & "Mb"
   
ElseIf filesize >= 1000 Then
    filesize = Val(filesize) / 1000
    Label11.Caption = "Filesize:" & Int(filesize) & "Kb"
    
    Else
    Label11.Caption = "Filesize:" & Int(filesize) & "Bytes"
     
    End If
    'End of getting filesize
    

 On Error GoTo PicError
 
 
     Label12.Visible = False
     Image1.Visible = True

If File1.Pattern = Text1.Text Then
    Module1.show
    Image1.Picture = LoadPicture(c)
    Image2.Picture = Image1.Picture
    Label7.Caption = "Width=" & Image2.Width
    Label8.Caption = "Height=" & Image2.Height
    
    End If
    
     Exit Sub
     
PicError:
    
    Label7.Caption = ""
    Label8.Caption = ""
    Image1.Visible = False
     Label12.Visible = True
     
     
     Exit Sub

   
    
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        
        PopupMenu mnuright
    End If
End Sub

Private Sub Form_Load()
    Text1.Enabled = False
    Text2.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    Check5.Enabled = False
    Drive1.Drive = "c"
    File1.Enabled = False
    mnuright.Visible = False
    
End Sub

Private Sub Image1_Click()

    If File1.ListIndex = -1 Then
    Image1.Enabled = False
    Else
    Image1.Enabled = True
    pic.Image1.Picture = Image1.Picture
    pic.show
    Form1.Visible = False
    End If
''''''''''''''''''
If Image1.Width >= 10200 Then
    pic.Left = 0
    pic.Top = 0
ElseIf Image1.Height >= 7680 Then
    pic.Top = 0
Else
End If

End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
        
        PopupMenu mnuright
    End If
    
End Sub



Private Sub mnuDlete_Click()

    Kill (Dir1.Path + "\" + File1.FileName)
    File1.Refresh
    Label3.Caption = "Files Found:" & File1.ListCount
    Label4.Caption = "Current File : " & (File1.ListIndex + 1) & "/" & File1.ListCount
    
End Sub

