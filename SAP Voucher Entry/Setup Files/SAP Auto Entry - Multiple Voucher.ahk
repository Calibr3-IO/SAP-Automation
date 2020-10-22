#SingleInstance Force
#NoEnv
#Persistent
SetWorkingDir %A_ScriptDir%
SendMode Input
SetBatchLines -1
DetectHiddenWindows On
Process Priority,, Realtime
#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.

setKeyDelay, 0
;Clipboard := ; Erase clipboard
XL := ComObjActive("Excel.Application")
sap_wb := XL.Workbooks("SAP Auto-Entry.xlsb").Activate
sap_ws := XL.Worksheets(2)

MsgBox,1 , SAP F-02 Auto Entry, The SAP F-02 Auto-Entry is Active.`nPress "Pause" Button at anytime to pause the process.`n`n`n`t`t`t- Muhannad Kamal`n`t`t`t (calibr3.io@gmail.com),5

IfMsgBox Cancel
ExitApp

cell = 65
num = 5
Loop
{
    Transform, OutputChar, Chr, %cell%
    if (sap_ws.Range(OutputChar num).Value) = "NA"
    {
        MsgBox, , Done, Done
        Break
    }
    Transform, OutputChar, Chr, %cell%
    if (sap_ws.Range(OutputChar num).Value) = "D"
    {
        WinActivate, ahk_class SAP_FRONTEND_SESSION
        WinWaitActive, ahk_class SAP_FRONTEND_SESSION
        Sleep, 711
        cell++
        Transform, OutputChar, Chr, %cell%
        sap_ws.Range(OutputChar num).Copy
        Sleep, 711
        SendInput, ^v
        Sleep, 711
        SendInput, {Tab}
        Sleep, 711
        SendInput, SA
        Sleep, 711
        SendInput, {Tab}
        Sleep, 711
        sap_ws.Range("C1").Copy
        SendInput, ^v
        Sleep, 711
        SendInput, {Tab}
        Sleep, 711
        Transform, OutputChar, Chr, %cell%
        sap_ws.Range(OutputChar num).Copy
        SendInput, ^v
        Sleep, 711
        SendInput, {Tab}
        sap_ws.Range("E1").Copy
        Sleep, 711
        SendInput, ^v
        SendInput, {Tab}
        Sleep, 711
        SendInput, SAR
        Sleep, 711
        SendInput, {Tab 4}
        Sleep, 711
        cell += 5
        Transform, OutputChar, Chr, %cell%
        sap_ws.Range(OutputChar num).Copy
        SendInput, ^v
        Sleep, 711
        SendInput, {Tab 2}
        Sleep, 711
        cell++
        Transform, OutputChar, Chr, %cell%
        sap_ws.Range(OutputChar num).Copy
        SendInput, ^v
        Sleep, 711
        SendInput, +{F8}
        Sleep, 3001
        ;Check Calculate Tax
        SendInput, {ShiftDown}{Tab}{ShiftUp}{Space}{Tab}
        Sleep, 1001
        cell = 65
        Transform, OutputChar, Chr, %cell%
        num++
    }
    sap_ws.Range(OutputChar num).Copy
    if (sap_ws.Range(OutputChar num).Value) = "N"
    {
        ;Simulate & Save Document
        SendInput, !ds
        ;SplashTextOn,,,`Press Enter When Ready!...
        ;KeyWait, Enter, D
        Sleep, 3000
        ;SplashTextOff
        SendInput, ^s
        Sleep, 500
        ;Retrieve Document Number and Save/Append in Text File
        CoordMode, Mouse, Client
        WinActivate, ahk_class SAP_FRONTEND_SESSION
        Click, 50, 988
        Sleep, 3000
        WinActivate, ahk_class DialogBox Container Class
        Sleep, 2000
        Click, 120, 50, 2
        SendInput, ^c
        Sleep, 711
        FileAppend,%clipboard% %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% `n, Vouchers.txt
        SplashTextOn,,,Voucher Saved...
        Sleep, 2000
        SplashTextOff
        SendInput, !{F4}
        Sleep, 501
        num++
        WinActivate, ahk_class SAP_FRONTEND_SESSION
        WinWaitActive, ahk_class SAP_FRONTEND_SESSION
        Sleep, 711
    }
    Else
    {
        WinActivate, ahk_class SAP_FRONTEND_SESSION
        WinWaitActive, ahk_class SAP_FRONTEND_SESSION
        Sleep, 311
        SendInput, ^v{Tab}
        Sleep, 311
        if (cell)=72
        {
            cell = 65
            num++
        }
        Else
        {
            cell++
        }
    }
}

;---------------------
Pause::Pause
!^r::Reload