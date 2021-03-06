﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

setKeyDelay, 0
Clipboard := ; Erase clipboard
XL := ComObjActive("Excel.Application")
sap_auto_entrywb := XL.Workbooks("SAP Auto-Entry.xlsb").Activate
sap_auto_entryws := XL.Worksheets(2)

;Document Entry;

sap_auto_entryws.Range("A1").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v{Tab}
Sleep, 51
sap_auto_entryws.Range("B1").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v{Tab}
Sleep, 51
sap_auto_entryws.Range("C1").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v{Tab}
Sleep, 51
sap_auto_entryws.Range("D1").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v{Tab}
Sleep, 51
sap_auto_entryws.Range("E1").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v{Tab}
Sleep, 300
sap_auto_entryws.Range("F1").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, {CtrlDown}v{CtrlUp}{Tab}{Tab}{Tab}{Tab}
Sleep, 51
sap_auto_entryws.Range("G1").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, {CtrlDown}v{CtrlUp}{Tab}{Tab}
Sleep, 51
sap_auto_entryws.Range("H1").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, {CtrlDown}v{CtrlUp}{ShiftDown}{F8}{ShiftUp}
Sleep, 3000

;Fast Entry;


sap_auto_entryws.Range("A4:D36").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
Sleep, 1000
SendInput, {ShiftDown}{Tab}{ShiftUp}{Space}{Tab}
Sleep, 1000
SendInput, ^v{Tab}{Tab}{Tab}{Tab}
Sleep, 100
sap_auto_entryws.Range("E4:G36").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v{Tab}{Tab}{Tab}
Sleep, 100
sap_auto_entryws.Range("H4:H36").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, {CtrlDown}v{CtrlUp}{AltDown}{d}{AltUp}s         ;Simulate Voucher
Sleep, 3000
SendInput, ^s
Sleep, 500
CoordMode, Mouse, Client
WinActivate, ahk_class SAP_FRONTEND_SESSION
Click, 50, 988
Sleep, 3000
WinActivate, ahk_class DialogBox Container Class
Sleep, 2000
Click, 120, 50, 2
SendInput, ^c
Sleep, 1111
FileAppend,%clipboard% %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% `n, Vouchers.txt
SplashTextOn,,,Voucher Saved...
Sleep, 2000
SplashTextOff
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;





Pause::Pause
^!r::Reload