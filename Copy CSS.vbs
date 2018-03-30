Dim FSO, Src

SRC = "C:\WinCDS\Matt's Work\Help\help.css"
SRC2 = "C:\WinCDS\Matt's Work\Help\Backing.bmp"
SRC3 = "C:\WinCDS\Matt's Work\Help\bulBlu04.png"
SRC4 = "C:\WinCDS\Matt's Work\Help\bulYel04.png"
SRC5 = "C:\WinCDS\Matt's Work\Help\bulGre04.png"
Set FSO = CreateObject("Scripting.FileSystemObject")

On Error Resume Next


Sub DoDir(Str)
'Msgbox "Dir: " & Str
	Dim X, FC, F1
	Dim D
  
	If Right(Str,1) <> "\" Then Str = Str & "\"
	D = Str & "help.css"
	If UCase(SRC) <> UCase(D) Then
          FSO.CopyFile SRC,D, True
'          FSO.CopyFile SRC2,D, True
'          FSO.CopyFile SRC3,D, True
'          FSO.CopyFile SRC4,D, True
'          FSO.CopyFile SRC5,D, True
        End If

  Set X = FSO.GetFolder(Str)
  Set FC = X.SubFolders
  For Each F1 In FC
		DoDir Str & F1.Name
  Next
End Sub

DoDir "C:\WinCDS\Matt's Work\Help\"

MsgBox "Help.css copied to all sub-folders.",,"Operation Complete"