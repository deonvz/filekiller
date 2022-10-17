Attribute VB_Name = "Module1"
Public Function show()
 MousePointer = 11
 ' Created by Deon van Zyl
    
    ''''''''''''''''''''''Testing''''''''''''''''''''
    ' Initialize for search, then perform recursive search.
    Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
    Dim result As Integer
      ' Check what the user did last.
        
    
        ' Update dir1.Path if it is different from the currently
        ' selected directory, otherwise perform the search.
        If Form1.Dir1.Path <> Form1.Dir1.List(Form1.Dir1.ListIndex) Then
            Form1.Dir1.Path = Form1.Dir1.List(Form1.Dir1.ListIndex)
           ' Exit Sub         ' Exit so user can take a look before searching.
        End If
    
        ' Continue with the search.
       
    
    
    
        Form1.File1.Pattern = Form1.Text1.Text
        FirstPath = Form1.Dir1.Path
        DirCount = Form1.Dir1.ListCount
    
        ' Start recursive direcory search.
        NumFiles = 0                       ' Reset found files indicator.
        Form1.File1.Path = Form1.Dir1.Path
    
        MousePointer = 0
       
        '''''''''''''''''''''''''''''''''''''''End of Test'''''''''''''
        If Form1.Check1.Value = 0 And Form1.Check6.Value = 0 Then
        Form1.File1.Pattern = "*.*"
        End If
        MousePointer = 0
        Form1.Label3.Caption = "Files Found:" & Form1.File1.ListCount
End Function

Public Function slide()
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

    pic.Image1.Picture = Form1.Image2.Picture
    pic.Image1.Width = Form1.Image2.Width
    pic.Image1.Height = Form1.Image2.Height
    pic.Height = Form1.Image2.Height
    pic.Width = Form1.Image2.Width

    
If pic.Image1.Width >= 10000 Then
    pic.Left = 0
    pic.Top = 0
ElseIf pic.Image1.Height >= 7650 Then
    pic.Top = 0
Else
    pic.WindowState = 0
End If
pic.Caption = "File Killer :" & Form1.Label4.Caption & " Filesize: " & Form1.Label11.Caption
End Function
