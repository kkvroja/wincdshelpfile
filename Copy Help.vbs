Dim FSO, Target
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run "C:\Progra~1\HTMLHe~1\hhc.exe WinCDS.hhp" , , True

on error resume next
  Target = "C:\WinCDS\Matt's Work\":                                    FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear
  Target = "C:\Program Files\Microsoft Visual Studio\VB98\":            FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear
  Target = "C:\Program Files\WinCDS\":                                  FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear
  Target = "C:\Upgrades v9  Auto 03052005\":                         FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear
  Target = "\\Inventory\C\Program Files\Microsoft Visual Studio\VB98\": FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear
  Target = "\\Inventory\C\Program Files\WinCDS\":                       FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear
  Target = "\\Inventory\C\WinCDS\Matt's Work\":                         FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear
  Target = "\\Inventory\C\WinCDS\Current\Help\":                        FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear
  Target = "\\Inventory\C\Upgrades v9  Auto 03052005\":                 FSO.CopyFile "WinCds.chm", Target, True
    if err.number <> 0 then msgbox "Failed: " & target:err.clear

  MsgBox "Help files compiled and updated."