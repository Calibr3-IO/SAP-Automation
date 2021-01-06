;Created By Muhannad Kamal V3

#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetTitleMatchMode, RegEx
CoordMode, Mouse, Client

/*
Created By:
Muhannad Kamal
+966531577001
calibr3.io@gmail.com
https://github.com/Calibr3-IO
*/

setKeyDelay, 0
Clipboard := ; Erase clipboard

XL := ComObjActive("Excel.Application")
sap_wb := XL.Workbooks("SAP CDN Document Auto Print.xlsb").Activate
sap_ws := XL.Worksheets(1)

;Document Auto Print

row := 2

If (sap_ws.Range("B" row).Value) = ""
{
    MsgBox, 64, SAP CDN Document Auto Print, No CDN Documents to Print, 5
    ExitApp
}

IfWinExist, Credit*
{
    WinActivate
}
Else
{
    Run, %A_ScriptDir%\PE1 CDN Memo Print.sap
    WinActivate, Credit*
    WinWaitActive, Credit*
    If (sap_ws.Range("E2").Value) = "Debit Note"
    {
        Sleep, 311
        SendInput, {ShiftDown}{Tab}{ShiftUp}
        Sleep, 311
        SendInput, {Down}
        Sleep, 311
        SendInput, {Tab}
        Sleep, 311
    }
}

Loop
{
    If (sap_ws.Range("B" row).Value) = ""
    {
        MsgBox, 64, SAP CDN Document Auto Print, Done Printing All Documents :)`nCreated By Muhannad Kamal`n+966531577001`ncalibr3.io@gmail.com`n https://github.com/Calibr3-IO, 11
        ExitApp
    }
    Else
    {
        WinActivate, ahk_class SAP_FRONTEND_SESSION
        WinWaitActive, ahk_class SAP_FRONTEND_SESSION
        Sleep, 311
        sap_ws.Range("B" row).Copy
        Sleep, 311
        SendInput, ^v{Tab}
        Sleep, 311
        sap_ws.Range("C2").Copy
        Sleep, 311
        SendInput, ^v{Tab}
        Sleep, 311
        sap_ws.Range("D2").Copy
        Sleep, 311
        SendInput, ^v
        Sleep, 511
        SendInput, {F8}
        Loop,2 ;Mouse Cursor Wait Check
        {
            Sleep, 1111
            If A_Cursor != Wait
                Break
        }
        ;WinActivate, .*Print ahk_class SAP_FRONTEND_SESSION
        ;WinWaitActive, .*Print ahk_class SAP_FRONTEND_SESSION
        Sleep, 1500
        Click, 337, 17
        Sleep, 1500
        Loop,2 ;Mouse Cursor Wait Check
        {
            Sleep, 1111
            If A_Cursor != Wait
                Break
        }
        WinActivate, Print
        WinWaitActive, Print
        SendInput, {Enter}
        Sleep, 1111
        Click, 260, 17
    }
    Sleep, 1111
    row++
}

Clipboard := ; Erase clipboard

ExitApp
;---------------------
Pause::Pause
^Home::Reload

/*
Created By:
Muhannad Kamal
+966531577001
calibr3.io@gmail.com
https://github.com/Calibr3-IO
*/