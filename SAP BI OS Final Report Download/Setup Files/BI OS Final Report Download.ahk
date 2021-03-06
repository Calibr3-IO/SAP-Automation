/*
Version 1
Created By:
Muhannad Kamal
+966531577001
calibr3.io@gmail.com
https://github.com/Calibr3-IO
*/

#SingleInstance Force
#NoEnv
#Warn
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
DetectHiddenWindows On
Process Priority,, High
#WinActivateForce
SetDefaultMouseSpeed 0
ListLines Off
CoordMode Mouse, Client
CoordMode Pixel, Client

JT := "BI OS Final Report Download - JTECO.txt"
JP := "BI OS Final Report Download - JAPCO.txt"
JH := "BI OS Final Report Download - JAHACO.txt"
JTB := "BI OS Final Report Download - JTECO Bahrain.txt"

Progress, B CBRed CWWhite X771 Y0 H47 ZX5 ZY5 ZH15 FS11 P0,- Starting BI OS Report Run -
SplashImage,J Logo.jpg, B X716 Y0 ZH47

InputBox, Date, Input Date, Date as YYYYMM`nExample for JUL 2020 - Enter as 202007`n`n, , , , , , , ,%A_YYYY%%A_MM%
  If ErrorLevel = 1
  ExitApp

InputBox, company, Input Company Code, Input Company Code:`n`n7100 - JTECO`n2500 - JAPCO`n3400 - JAHACO`n7600 - JTECO Bahrain, , ,230 , , , , , 7100
if ErrorLevel = 1
  ExitApp

InputBox, report, Report Type, Input Report Type:`n`n010 - Actual`n020 - Plan`n025 - BS Plan YTD`n060 - Forecasting`n065 - Forecasting YTD`n080 - Selection`n085 - Accrual Plan`n000 - Not Assigned, ,375 ,291 , , , , , 010
if ErrorLevel = 1
  ExitApp

FormatTime, Date1, %Date%, MMyyyy

If (company = 7100)
{
  file := % JT
}
If (company = 2500)
{
  file := % JP
}
If (company = 3400)
{
  file := % JH
}
If (company = 7600)
{
  file := % JTB
}
Run, "C:\Program Files\Internet Explorer\iexplore.exe" http://boprdas1:8080/BOE/BI?startFolder=FsajG1yZuAoA52YAAADH0vYRAFBWiGCp&isCat=false
Sleep, 500
WinWait, ahk_class IEFrame ahk_exe iexplore.exe
Sleep, 500
WinActivate, ahk_class IEFrame ahk_exe iexplore.exe
Sleep, 500
WinWaitActive, ahk_class IEFrame ahk_exe iexplore.exe
Sleep, 500
WinMaximize, ahk_class IEFrame ahk_exe iexplore.exe
Sleep, 500
;-----Company Selection-----
Loop ;Initial Load Check
{
  PixelGetColor, OutputVar, 950, 665, RGB
  If OutputVar != 0XE6E6E6
    Break
}
Progress, 1
Sleep, 1500
SendInput, ^0 ;Default Browser Zoom

Loop ;Main Window Selection Load Check
{
  PixelGetColor, OutputVar, 1049, 791, RGB
  If OutputVar = 0XE5EAF3
    Break
}

Progress, 5, - Company Selection -

Sleep, 1111
Click, 794, 507 ;Company Selection
Sleep, 1111
Click, 1173, 722 ;Company Search Key Drop-Down
Sleep, 1111
Click, 1204, 769 ;Company Search Key
Sleep, 1111
Progress, 7, - Clearing Existing Data -
Click, 1273, 512 ;Existing Data
Sleep, 1111
SendInput, {ShiftDown}{End}{ShiftUp}
Sleep, 1111
SendInput, {ShiftDown}{Tab 2}{ShiftUp}
Sleep, 1111
SendInput, {Space}
Sleep, 1111
Click, 1273, 512 ;Existing Data
Sleep, 1111
SendInput, {ShiftDown}{End}{ShiftUp}
Sleep, 1111
SendInput, {ShiftDown}{Tab}{ShiftUp}
Sleep, 1111
SendInput, {Space}
Sleep, 1111
Progress, 8, - Searching Company -
Click, 1096, 720 ;Company Search Entry Area
Sleep, 1111
SendInput, %Company%
Sleep, 1111
SendInput, {Enter}
Sleep, 1500
Loop ;Mouse Cursor Wait Check
{
  Sleep, 1111
  If A_Cursor != Wait
    Break
}
Sleep, 1500
Progress, 9, - Selecting Company -
Click, 1129, 583, 2 ;Select Company
Sleep, 1500
;-----Company Selection-----

Progress, 10, - Date Selection -

;-----Date Selection-----
Click, 794, 524 ;Month Selection
Sleep, 1500
Progress, 13, - Searching Date -
Click, 1096, 720 ;Month Search Entry Area
SendInput, %Date1%
Sleep, 1111
SendInput, {Enter}
Sleep, 1500
Loop ;Mouse Cursor Wait Check
{
  Sleep, 1111
  If A_Cursor != Wait
    Break
}
Sleep, 1500
Progress, 17, - Selecting Date -
Click, 1129, 583, 2 ;Select Month
Sleep, 1500
;-----Date Selection-----

Progress, 20, - Selecting Report Type -

;-----Report Selection-----
Click, 794, 534 ;Report Selection
Sleep, 1500
Progress, 25, - Searching Report Type -
Click, 1102, 720 ;Report Search Entry Area
SendInput, %report%
Sleep, 1111
SendInput, {Enter}
Sleep, 1500
Loop ;Mouse Cursor Wait Check
{
  Sleep, 1111
  If A_Cursor != Wait
    Break
}
Sleep, 1500
Progress, 27, - Selecting Report Type -
Click, 1129, 583, 2 ;Select Report
Sleep, 1500
;-----Report Selection-----

Progress, 30, - Selecting Profit Centre -

;-----Profit Centre Selection-----
FormatTime, Date2, %Date%, MMyy

Loop, Read, %file%
{
  prog_index := A_Index
}

calc_prog_index := (65 / prog_index)

lines := 30 + calc_prog_index

Loop, read, %file%
{
  Sleep, 1500
  Click, 794, 560 ;Profit Centre Selection
  Sleep, 2500
  Click, 1270, 510 ;Existing Data
  Sleep, 1111
  SendInput, {ShiftDown}{End}{ShiftUp}
  Sleep, 1111
  SendInput, {ShiftDown}{Tab}{ShiftUp}
  Sleep, 1111
  SendInput, {Space}
  Sleep, 1111
  Click, 1097, 701 ;Profit Centre Search Entry Area
  Sleep, 1111

  fileline := A_LoopReadLine

  stringsplit, fileParse, fileline, "#"

  Loop, parse, fileParse2, CSV, %A_Space%%A_Tab%
  {
    PC := A_LoopField
    SendInput, *%PC%
    Sleep, 1111
    SendInput, {Enter}
    Loop ;Mouse Cursor Wait Check
    {
      Sleep, 1111
      If A_Cursor != Wait
        Break
    }
    Sleep, 1111
    Click, 1129, 561, 2 ;Select Profit Center
    Sleep, 1500
    SendInput, {Tab}
    Sleep, 3100
  }
  Click, 1228, 800 ;Select Ok
  Sleep, 2500
  Progress, %lines%, - Running OS Report for %fileParse1% -
    Loop ;Report Load Check
  {
    PixelGetColor, OutputVar, 1020, 688, RGB
    If OutputVar != 0XE5EAF3
      Break
  }
  Sleep, 1500
  SendInput, {Enter}

  ;-----Report Rename-----
  Sleep, 1000
  Click, 1847, 197, 2 ;Select Design Button
  Sleep, 1500
  Loop
  {
    Sleep, 1111
    If A_Cursor != Wait
      Break
  }
  Sleep, 2100
  Click, 881, 366, 2 ;Select Header Cell
  Sleep, 2100
  SendInput, %fileParse1%
  Sleep, 1500
  SendInput, {Enter}
  Loop
  {
    Sleep, 1111
    If A_Cursor != Wait
      Break
  }
  Sleep, 1500
  SendInput, {Enter}
  Sleep, 1500
  Click, 1773, 197, 2 ;Select Reading Button
  Sleep, 1500
  Loop
  {
    Sleep, 1111
    If A_Cursor != Wait
      Break
  }
  ;-----Report Rename-----

  Sleep, 2500
  Click, 322, 191 ;Export
  Sleep, 1500
  Click, 397, 242 ;Export Current Document
  Sleep, 1500
  Click, 515, 242 ;Export Current Document As PDF
  Loop ;Mouse Cursor Wait Check
  {
    Sleep, 1111
    If A_Cursor != Wait
      Break
  }
  Sleep, 1500
  Sleep, 500
  Click, 1318, 1025 ;Click Save Drop-Down
  Sleep, 1500
  Click, 1398, 1003 ;Select Save as
  Sleep, 1500
  WinWaitActive, Save As ahk_exe iexplore.exe ahk_class #32770
  Sleep, 1111
  WinActivate, Save As ahk_exe iexplore.exe ahk_class #32770
  Sleep, 1500
  If (company = 7100)
  {
    SendInput, %A_Index%. JT Operating Statements %fileParse1%%Date2%
  }
  If (company = 2500)
  {
    SendInput, %A_Index%. JP Operating Statements %fileParse1%%Date2%
  }
  If (company = 3400)
  {
    SendInput, %A_Index%. JH Operating Statements %fileParse1%%Date2%
  }
  If (company = 7600)
  {
    SendInput, %A_Index%. JTB Operating Statements %fileParse1%%Date2%
  }
  Sleep, 1500
  Click, 785, 476 ;Save File
  Sleep, 1500
  WinWaitClose, Save As ahk_exe iexplore.exe ahk_class #32770
  Sleep, 300
  WinActivate, BI launch pad - Internet Explorer ahk_exe iexplore.exe ahk_class IEFrame
  Sleep, 300
  WinWaitActive, BI launch pad - Internet Explorer ahk_exe iexplore.exe ahk_class IEFrame
  Sleep, 1500
  Click, 453, 196 ;Refresh Report
  Sleep, 500
  lines := lines + calc_prog_index
  Sleep, 500
  Loop ;Report Load Check
  {
    PixelGetColor, OutputVar, 1020, 688, RGB
    If OutputVar != 0XE5EAF3
      Break
  }
  Sleep, 1500
  Loop ;Main Window Selection Load Check
  {
    PixelGetColor, OutputVar, 1049, 791, RGB
    If OutputVar = 0XE5EAF3
      Break
  }
  Sleep, 2500
  Click, 794, 560 ;Profit Centre Selection
  Sleep, 2500
  Click, 1176, 704 ;Profit Centre Search Key Drop-Down
  Sleep, 1111
  Click, 1196, 749 ;Profit Centre Search Key
  Sleep, 500
  Progress, %lines%, - Running OS Report for %fileParse1% -
  }
Progress, 100, - BI OS Report Run Successfully Completed -
;-----Profit Centre Selection-----

MsgBox, , Success, BI Operating Statements Run Complete :)`n`n`nCreated By:`n`nMuhannad Kamal`n+966531577001`ncalibr3.io@gmail.com`nhttps://github.com/Calibr3-IO, 11

ExitApp

;---------------------
Pause::Pause
!^r::Reload

  /*
Created By:
Muhannad Kamal
+966531577001
calibr3.io@gmail.com
https://github.com/Calibr3-IO
*/