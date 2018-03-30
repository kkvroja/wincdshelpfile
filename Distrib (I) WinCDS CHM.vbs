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
  Dim Ds(7)
  DS(0) = "C:\WinCDS\Matt's Work\"
  DS(1) = "C:\Program Files\Microsoft Visual Studio\VB98\"
  DS(2) = "C:\Program Files\WinCDS\"
  DS(3) = "C:\Documents And Settings\Jerry Katz\Desktop\"
  DS(4) = "C:\CDS Customer Order\Support\"
  DS(5) = "C:\CDS Customer Order Station (Paul)\Support\"
  DS(6) = "C:\CDS Demo\Support\"
  DS(7) = "\\dc1\inetpub\wwwcds\\WebUpdate\"


  For Each L in DS
    Target = L
    FSO.CopyFile "C:\WinCDS\Matt's Work\Help\WinCDS.chm", Target, True
    if Err.Number <> 0 Then 
      MsgBox "Failed to copy CHM to " & Target, vbInformation, "Failed"
      Err.Clear
    End If
  Next

  MsgBox "Help files distributed.", vbinformation, "Done"