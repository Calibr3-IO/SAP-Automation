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

JT := "BI EX Final Report Download - JTECO.txt"
JP := "BI EX Final Report Download - JAPCO.txt"
JH := "BI EX Final Report Download - JAHACO.txt"
JTB := "BI EX Final Report Download - JTECO Bahrain.txt"

Progress, B CBRed CWWhite X771 Y0 H47 ZX5 ZY5 ZH15 FS11 P0,- Starting BI EX Report Run -
SplashImage,J Logo.jpg, B X716 Y0 ZH47

;Run "C:\Program Files\Internet Explorer\iexplore.exe" http://boprdas1:8080/BOE/BI?startFolder=FsajG1yZuAoA52YAAADH0vYRAFBWiGCp&isCat=false
InputBox, Date, Input Date, Date as YYYYMM`nExample for JUL 2020 - Enter as 202007, , , , , , , , %A_YYYY%%A_MM%

InputBox, Company, Company Code, Input Company Code:`n`n7100 - JTECO`n2500 - JAPCO`n3400 - JAHACO, , , , , , , , 7100

FormatTime, Date1, %Date%, YYYYMM

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

;-----Company Selection-----
Loop ;Main Prompts Window Color Check - Initial
{
  PixelGetColor, OutputVar, 1044, 865, RGB or
  If OutputVar = 0XE5EAF3
    Break
  PixelGetColor, OutputVar, 1008, 909 , RGB
  If OutputVar = 0XE5EAF3
    Break
}
Progress, 1
Sleep, 1500
SendInput, ^0 ;Default Browser Zoom

Progress, 5, - Company Selection -
Sleep, 1500
Loop,
{
  PixelGetColor, OutputVar, 1008, 909 , RGB
  If OutputVar = 0XE5EAF3
  {
    Click, 1141, 402
    Break
  }
  Else
  {
    Click, 1143, 442 ;Company Picker
    Break
  }
}

Sleep, 1111
Loop ;Company Selection Window Color Check
{
  PixelGetColor, OutputVar, 955, 740, RGB
  If OutputVar = 0XE5EAF3
    Break
}
Sleep, 1500
Click, 785, 497 ;Clear Existing Values
Sleep, 1500
Click, 907, 526, 2 ;Select Company Search Field
Sleep, 1111
SendInput, %Company%
Sleep, 500
SendInput, {Enter}
Sleep, 1500
Progress, 7, - Selecting Company -
Click, 907, 551 ;Select Company
Sleep, 1111
Click, 1042, 743 ;Select Ok
;-----Company Selection-----

;-----Date Selection-----
Progress, 10, - Date Selection -
Loop ;Main Prompts Window Color Check - Initial
{
  PixelGetColor, OutputVar, 1044, 865, RGB or
  If OutputVar = 0XE5EAF3
  {
    Sleep, 1500
    Click, 1219, 544 ;Select Date Selection Drop-Down
    Sleep, 1500
    Loop
    {
      PixelGetColor, OutputVar, 1065, 565 , RGB
      If OutputVar = 0XE4E9F0
        Break
    }
    Click, 1003, 595, 2 ;Select Date Selection Field
    Sleep, 1500
    SendInput, %Date%{Enter}
    Sleep, 1111
    SendInput, {Enter}
    Sleep, 1111
    Progress, 15, - Selecting Date -
    Click, 1040, 620 ;Select Date
    Sleep, 1111
    Break
  }
  Else
  {
    PixelGetColor, OutputVar, 1008, 909 , RGB
    If OutputVar = 0XE5EAF3
    {
      Sleep, 1500
      Click, 1220, 503 ;Select Date Selection Drop-Down
      Sleep, 1500
      Loop
      {
        PixelGetColor, OutputVar, 1085, 527 , RGB
        If OutputVar = 0XE4E9F0
          Break
      }
      Click, 1003, 554, 2 ;Select Date Selection Field
      Sleep, 1500
      SendInput, %Date%{Enter}
      Sleep, 1111
      SendInput, {Enter}
      Sleep, 1111
      Progress, 15, - Selecting Date -
      Click, 1020, 580 ;Select Date
      Sleep, 1111
      Break
    }
  }
}

;-----Date Selection-----

;-----Cost Centre Selection-----
FormatTime, Date2, %Date%, MMyy

Progress, 30, - Cost Centre Selection -
Loop ;Main Prompts Window Color Check - Initial
{
  PixelGetColor, OutputVar, 1044, 865, RGB or
  If OutputVar = 0XE5EAF3
  {
    Sleep, 1500
    Click, 1144, 600 ;Cost Centre Picker
    Sleep, 1500
    Break
  }
  Else
  {
    PixelGetColor, OutputVar, 1008, 909, RGB or
    If OutputVar = 0XE5EAF3
    {
      Sleep, 1500
      Click, 1141, 560 ;Cost Centre Picker
      Sleep, 1500
      Break
    }
  }
}
Loop ;Cost Centre Selection Window Color Check
{
  PixelGetColor, OutputVar, 963, 743, RGB
  If OutputVar = 0XE5EAF3
    Break
}
Sleep, 500
Loop ;Cost Centre Window Load Check
{
  PixelGetColor, OutputVar, 994, 629, RGB
  If OutputVar != 0XFBFCFD
    Break
}
Sleep, 1500
Click, 763, 497 ;Clear Existing Values
Sleep, 1500
Click, 907, 526, 2 ;Select Cost Centre Search Field
Sleep, 1111
Loop, Read, %file%
{
  prog_index := A_Index
}

calc_prog_index := (65 / prog_index)

lines := 30 + calc_prog_index

Loop, read, %file%
{
  fileline := A_LoopReadLine

  stringsplit, fileParse, fileline, "#"

  Loop, parse, fileParse2, CSV, %A_Space%%A_Tab%
  {
    CC := A_LoopField
    SendInput, %CC%
    Sleep, 1111
    SendInput, {Enter}
    Sleep, 1111
    Click, 907, 575 ;Select Cost Centre
    Sleep, 1111
    Loop ;Cost Centre Selection Window Color Check
    {
      PixelGetColor, OutputVar, 963, 743, RGB
      If OutputVar = 0XE5EAF3
        Break
    }
    Sleep, 1500
    Click, 907, 526, 2 ;Select Cost Centre Search Field
    Sleep, 1111
  }
  Sleep, 1111
  Click, 1046, 745 ;Select Ok
  Sleep, 1500
  Progress, %lines%, - Running EX Report for %fileParse1% -
    Loop ;Main Prompts Window
  {
    PixelGetColor, OutputVar, 1044, 865, RGB
    If OutputVar = 0XE5EAF3
    {
      Sleep, 1111
      Click, 1203, 863 ;Run Report
      Break
    }
    Else
    {
      PixelGetColor, OutputVar, 1008, 909, RGB
      If OutputVar = 0XE5EAF3
      {
        Sleep, 1111
        Click, 1201, 905 ;Run Report
        Break
      }
    }
  }
  Sleep, 2000
  Loop ;Report Load Check
  {
    PixelGetColor, OutputVar, 959, 672, RGB
    If OutputVar != 0XE5EBF4
      Break
  }
  Sleep, 1500
  Loop ;Print/Export Window Load Check
  {
    PixelGetColor, OutputVar, 951, 745, RGB
    If OutputVar != 0XE5EBF4
      Break
  }
  Sleep, 1500
  Click, 181, 192 ;Export Report Type
  Sleep, 1500
  Click, 1065, 545 ;Export Report Type Selection Drop-Down
  Sleep, 1500
  Click, 962, 593 ;Select Export Report Type
  Sleep, 1500
  Click, 1070, 733 ;Export Report
  Sleep, 1500
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
    SendInput, %A_Index%. JT Summary of Expenses %fileParse1%%Date2%
  }
  If (company = 2500)
  {
    SendInput, %A_Index%. JP Summary of Expenses %fileParse1%%Date2%
  }
  If (company = 3400)
  {
    SendInput, %A_Index%. JH Summary of Expenses %fileParse1%%Date2%
  }
  If (company = 7600)
  {
    SendInput, %A_Index%. JTB Summary of Expenses %fileParse1%%Date2%
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
  Click, 151, 193 ;Refresh Report
  Sleep, 1111
  Loop ;Report Load Check
  {
    PixelGetColor, OutputVar, 959, 672, RGB
    If OutputVar != 0XE5EBF4
      Break
  }
  Sleep, 1500
  Loop ;Main Prompts Window - Refreshed
  {
    PixelGetColor, OutputVar, 1008, 909 , RGB
    If OutputVar = 0XE5EAF3
      Break
  }

  Sleep, 1111
  Click, 1143, 562 ;Cost Centre Picker
  Loop ;Cost Centre Selection Window Color Check
  {
    PixelGetColor, OutputVar, 963, 743, RGB
    If OutputVar = 0XE5EAF3
      Break
  }
  Sleep, 500
  Loop ;Cost Centre Window Load Check
  {
    PixelGetColor, OutputVar, 994, 629, RGB
    If OutputVar != 0XFBFCFD
      Break
  }
  Sleep, 1500
  Click, 763, 497 ;Clear Existing Values
  Sleep, 1500
  Click, 907, 526, 2 ;Select Cost Centre Search Field
  Sleep, 1111
  Progress, %lines%, - Running EX Report for %fileParse1% -
  }
Progress, 100, - BI EX Report Run Successfully Completed -
;-----Cost Centre Selection-----

MsgBox, , Success, BI Summary of Expenses Run Complete :)`n`n`nCreated By:`n`nMuhannad Kamal`n+966531577001`ncalibr3.io@gmail.com`nhttps://github.com/Calibr3-IO, 11

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