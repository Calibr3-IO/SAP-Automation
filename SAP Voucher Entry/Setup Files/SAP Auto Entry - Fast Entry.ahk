#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

setKeyDelay, 0
Clipboard := ; Erase clipboard
XL := ComObjActive("Excel.Application")
sap_auto_entrywb := XL.Workbooks("SAP Auto-Entry.xlsx")
sap_auto_entryws := XL.Worksheets(2)

sap_auto_entryws.Range("A4:D36").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v{Tab}{Tab}{Tab}{Tab}
Sleep, 1
sap_auto_entryws.Range("E4:G36").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v{Tab}{Tab}{Tab}
Sleep, 1
sap_auto_entryws.Range("H4:H36").Copy
WinActivate, ahk_class SAP_FRONTEND_SESSION
WinWaitActive, ahk_class SAP_FRONTEND_SESSION
SendInput, ^v

Pause::Pause
^!r::Reload