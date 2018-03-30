Set WshShell = WScript.CreateObject("WScript.Shell")

If WScript.Arguments.length = 0 Then
Set ObjShell = CreateObject("Shell.Application")
ObjShell.ShellExecute "wscript.exe", """" & _
WScript.ScriptFullName & """" &_
 " RunAsAdministrator", , "runas", 1
Else

Dim FSO, Target, L, objFile
Set FSO = CreateObject("Scripting.FileSystemObject")

'On Error Resume Next

'Set objFile = CreateObject("DSOFile.OleDocumentProperties")
'objFile.Open("C:\WinCDS\Matt's Work\Help\WinCDS.chm")

'objFile.SummaryProperties.Title = "WinCDS Help File"
'objFile.SummaryProperties.ApplicationName = "WinCDS Help File"
'objFile.SummaryProperties.Author = "Custom Design Software"
'objFile.SummaryProperties.Company = "Custom Design Software"
'objFile.SummaryProperties.Version = 10
'objFile.Save
'objFile.Close

'Err.Clear



on error resume next
  Dim Ds(12)
  DS( 0) = "C:\CDSData\InventFX\"
  DS( 1) = "C:\WinCDs\WinCDS-Setup\CDSData\InventFX\"
  DS( 2) = ""
  DS( 3) = ""
  DS( 4) = ""
  DS( 5) = ""
  DS( 6) = "I:\InventFX\"
  DS( 7) = "I:\WinCDs\WinCDS-Setup\CDSData\InventFX\"
  DS( 8) = ""
  DS( 9) = ""
  DS(10) = ""
  DS(11) = ""
  DS(12) = "\\webserver1\inetpub\wwwcds\webupdate\"


  For Each L in DS
    Target = L
	If Target<>"" Then
      FSO.CopyFile "C:\WinCDS\WinCDS\Help\WinCDS.chm", Target, True
      If Err.Number <> 0 Then 
        MsgBox "Failed to copy CHM to " & Target, vbInformation, "Failed"
        Err.Clear
      End If
	End If
  Next

  MsgBox "Help files distributed.", vbinformation, "Done"

end if