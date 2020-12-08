#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory
;MODIFIED=20140605
;- Listview   >>>   Add New  / MODIFY / DELETE / SEARCH

CoordMode Mouse, Client
CoordMode Pixel, Client

MainWindowTitle=ListView_Test1
transform,T,chr,09
delim = -
first1:=0

F1 = JTECO PC.txt
F2 = JAPCO PC.txt
F3 = JAHACO PC.txt
F4 = JTECO-B PC.txt

P1 = JTECO PC
P2 = JAPCO PC
P3 = JAHACO PC
P4 = JTECO-B PC

; If (company = 7100)
; {
;   file := % JT
; }
; If (company = 2500)
; {
;   file := % JP
; }
; If (company = 3400)
; {
;   file := % JH
; }
; If (company = 7600)
; {
;   file := % JTB
; }

;gosub,test1 ;-- create 2 text-files for 2 ListViews  ( test )

Gui,2:default
Gui,2: Font,CDefault,Fixedsys.
Gui,2: Margin, 10, 10

Tabnumber:=1
gui,2:add, Tab2, x10 y10 w540 h250 gtabchange vTabnumber AltSubmit,JTECO - 7100|JAPCO - 2500|JAHACO - 3400|JTECO-B - 7600

gui,2:tab,1
gui,2:add, listview,x10 y40 w520 h400 grid cWhite backgroundteal hwndLV1 vLV1 gListViewEvents +altsubmit -multi, A|B
gosub,fill1
gosub,width1
Gui,2:add,button, x10 y450 w70 gPrintLV1,Print
Gui,2:Add, Edit, x100 y450 w250 gFind vSrch1,
Gui,2:add,button, x450 y450 w70 gAddNew1,Add New

gui,2:tab,2
gui,2:add, listview,x10 y40 w520 h400 grid cWhite backgroundteal hwndLV2 vLV2 gListViewEvents +altsubmit -multi, A|B
gosub,fill2
gosub,width1
Gui,2:add,button, x10 y450 w70 gPrintLV1,Print
Gui,2:Add, Edit, x100 y450 w250 gFind vSrch2,
Gui,2:add,button, x450 y450 w70 gAddNew1,Add New

gui,2:tab,3
gui,2:add, listview,x10 y40 w520 h400 grid cWhite backgroundteal hwndLV3 vLV3 gListViewEvents +altsubmit -multi, A|B
gosub,fill3
gosub,width1
Gui,2:add,button, x10 y450 w70 gPrintLV1,Print
Gui,2:Add, Edit, x100 y450 w250 gFind vSrch3,
Gui,2:add,button, x450 y450 w70 gAddNew1,Add New

gui,2:tab,4
gui,2:add, listview,x10 y40 w520 h400 grid cWhite backgroundteal hwndLV4 vLV4 gListViewEvents +altsubmit -multi, A|B
gosub,fill4
gosub,width1
Gui,2:add,button, x10 y450 w70 gPrintLV1,Print
Gui,2:Add, Edit, x100 y450 w250 gFind vSrch4,
Gui,2:add,button, x450 y450 w70 gAddNew1,Add New

gui,2: show,x10 y1 w600 h500,%MainWindowTitle%
gosub,tabchange
RETURN

2Guiclose:
2Guiescape:
exitapp

width1:
  T1=300
  T2=200
  LV_ModifyCol(1,T1)
  LV_ModifyCol(2,T2)
  ;LV_ModifyCol(2,"Integer")
return

;-------------------------------------------------------------------------------------
tabchange:
  GuiControlGet, Tabnumber
  GuiControl,2:Focus,srch%tabnumber%
Return
;-------------------------------------------------------------------------------------

;---------------- SEARCH -------------------
Find:
  Gui,2: Submit, Nohide
  Gui,2:listview, LV%Tabnumber%
  Fx=% F%Tabnumber%
  src:= % srch%Tabnumber%
  if (SRC="")
  {
    goto,Fill%Tabnumber%
    ; return
  }
  LV_Delete()
  loop,read,%fx%
  {
    LR=%A_loopReadLine%
    if SRC<>
    {
      if LR contains %src%
      {
        stringsplit,C,A_LoopReadLine,%delim%
        LV_Add("",C1,C2)
      }
    }
    else
      continue
  }
  LV_Modify(LV_GetCount(), "Vis")
  if (SRC="")
    goto,Fill%Tabnumber%
return

;------------------- LISTVIEW --------------
ListViewEvents:
  Gui,2:default
  Gui,2:listview, LV%Tabnumber%

  ; if(A_GuiEvent == "Normal")
  ; {
  ;   LV_GetText(C1, A_EventInfo, 1)
  ;   LV_GetText(C2, A_EventInfo, 2)
  ; }

  if(A_GuiEvent == "DoubleClick")
  {
    LV_GetText(C1, A_EventInfo, 1)
    LV_GetText(C2, A_EventInfo, 2)

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
    Run, "C:\Program Files\Internet Explorer\iexplore.exe" http://boprdas1:8080/BOE/BI?startFolder=FsajG1yZuAoA52YAAADH0vYRAFBWiGCp&isCat=false
    Sleep, 500
    WinWait, ahk_class IEFrame ahk_exe iexplore.exe
    Sleep, 500
    WinActivate, ahk_class IEFrame ahk_exe iexplore.exe
    Sleep, 500
    WinWaitActive, ahk_class IEFrame ahk_exe iexplore.exe
    Sleep, 500
    WinMaximize, ahk_class IEFrame ahk_exe iexplore.exe
    Loop ;Initial Load Check
    {
      PixelGetColor, OutputVar, 929, 662, RGB
      If OutputVar != 0XFFFFFF
        Break
    }
    Sleep, 2100
    SendInput, ^0 ;Default Browser Zoom
    Progress, 1
    Sleep, 500
    Click, 1726, 221 ;Sort Created on Column - To adjust the Operating Statement Row Position
    Sleep, 2100
    Click, 350, 222 ;Sort Title Column - To adjust the Operating Statement Row Position
    Sleep, 2100
    Click, Right, 413, 427
    Sleep, 1500
    Click, 461, 432
    ;-----Company Selection-----
    Loop ;Initial Load Check
    {
      PixelGetColor, OutputVar, 950, 665, RGB
      If OutputVar != 0XE6E6E6
        Break
    }
    Sleep, 1500
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
    Loop, parse, C2, CSV, %A_Space%%A_Tab%
    {
      prog_index := A_Index
    }

    calc_prog_index := (65 / prog_index)

    lines := 30 + calc_prog_index

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
    Loop, parse, C2, CSV, %A_Space%%A_Tab%
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
      Sleep, 1500
    }
    Click, 1228, 800 ;Select Ok
    Sleep, 2500
    Progress, %lines%, - Running OS Report for %C1% -
      Loop ;Report Load Check
    {
      PixelGetColor, OutputVar, 1020, 688, RGB
      If OutputVar != 0XE5EAF3
        Break
    }
    Sleep, 1500
    Progress, 100, - BI OS Report Run Successfully Completed -

    ; fileline := A_LoopReadLine

    ; stringsplit, fileParse, fileline, "-"

    ; Loop, parse, fileParse2, CSV, %A_Space%%A_Tab%
    ; {
    ;   PC := A_LoopField
    ;   SendInput, *%PC%
    ;   Sleep, 1111
    ;   SendInput, {Enter}
    ;   Loop ;Mouse Cursor Wait Check
    ;   {
    ;     Sleep, 1111
    ;     If A_Cursor != Wait
    ;       Break
    ;   }
    ;   Sleep, 1111
    ;   Click, 1129, 561, 2 ;Select Profit Center
    ;   Sleep, 1500
    ;   SendInput, {Tab}
    ;   Sleep, 1500
    ; }
    ; Click, 1228, 800 ;Select Ok
    ; Sleep, 2500
    ExitApp
  }

  if(A_GuiEvent == "RightClick")
  {
    LV_GetText(C1, A_EventInfo, 1)
    LV_GetText(C2, A_EventInfo, 2)
    gosub,Modify1
    return
  }

  if A_GuiEvent=K
  {
    GetKeyState,state,DEL ;- << DELETE
    if state=D
    {
      MsgBox, 4, Confirm Delete, Are you Sure you want to Delete?
      IfMsgBox Yes
      {
        RowNumber:=LV_GetNext()
        LV_Delete(RowNumber)
        gosub,Modify2
        Return
      }
      Else
      {
        Return
      }
    }
    return
  }
RETURN
;--------------------------------------------------------------

;----------------   MODIFY  ----------------
Modify1:
  Gui,3: +AlwaysonTop
  Gui,3: Font, s10, Verdana
  gui,3:listview, LV%Tabnumber%
  Gui,3:add, edit, w300 h30 vC1, %C1%
  Gui,3:add, edit, w300 h30 vC2, %C2%

  Gui,3: Add,Button, x12 gACCEPT1 default, Accept
  Gui,3: Add,Button, x+4 gCANCEL1, Cancel
  Gui,3:show,center, LV_Modify
return

accept1:
  Gui,2:default
  Gui,3:submit,nohide
  gui 3:listview, LV%Tabnumber%
  RowNumber := LV_GetNext()
  c1:= % c1
  c2:= % c2
  lv_modify(rownumber, "col1" , C1 )
  lv_modify(rownumber, "col2" , C2 )
  gosub,modify2
  Gui,3:destroy
return

cancel1:
3Guiclose:
  Gui,3:destroy
return
;-----------------------------------------------------------------

;----------------   ADD NEW  ----------------
ADDNEW1:
  Gui,4: +AlwaysonTop
  Gui,4: Font, s10, Verdana
  Gui,4:listview, LV%Tabnumber%
  Gui,4:add, edit, w300 h30 vC1,
  Gui,4:add, edit, w300 h30 vC2,

  Gui,4: Add,Button, x12 gACCEPT4 default, Accept
  Gui,4: Add,Button, x+4 gCANCEL4, Cancel
  Gui,4:show,center,Add New
return

accept4:
  Gui,2:default
  Gui,4:submit,nohide
  Gui,4:listview, LV%Tabnumber%
  Fx=% F%Tabnumber%
  Fileappend,`n%c1%%A_Space%%delim%%A_Space%%c2%,%fx%
  gosub,fill%tabnumber%
  Gui,4:destroy
return

cancel4:
4Guiclose:
  Gui,4:destroy
return
;-----------------------------------------------------------------

;------------------ FILL ----------------------------
Fill1:
  gui,2:listview, listview%Tabnumber%
  LV_Delete()
  loop,read,%F1%
  {
    LR=%A_loopReadLine%
    if LR=
      continue
    C1 =
    C2 =
    stringsplit,C,LR,%delim%
    LV_Add("", c1,c2)
  }
  LV_ModifyCol(1, "Sort CaseLocale") ; or "Sort CaseLocale"
  LV_Modify(LV_GetCount(), "Vis") ;scrolls down
return

Fill2:
  gui, 2:listview, listview%Tabnumber%
  LV_Delete()
  loop,read,%F2%
  {
    LR=%A_loopReadLine%
    if LR=
      continue
    C1 =
    C2 =
    stringsplit,C,LR,%delim%,
    LV_Add("", c1,c2)
  }
  LV_ModifyCol(1, "Sort CaseLocale") ; or "Sort CaseLocale"
  LV_Modify(LV_GetCount(), "Vis") ;scrolls down
return

Fill3:
  gui, 2:listview, listview%Tabnumber%
  LV_Delete()
  loop,read,%F3%
  {
    LR=%A_loopReadLine%
    if LR=
      continue
    C1 =
    C2 =
    stringsplit,C,LR,%delim%,
    LV_Add("", c1,c2)
  }
  LV_ModifyCol(1, "Sort CaseLocale") ; or "Sort CaseLocale"
  LV_Modify(LV_GetCount(), "Vis") ;scrolls down
return

Fill4:
  gui, 2:listview, listview%Tabnumber%
  LV_Delete()
  loop,read,%F4%
  {
    LR=%A_loopReadLine%
    if LR=
      continue
    C1 =
    C2 =
    stringsplit,C,LR,%delim%,
    LV_Add("", c1,c2)
  }
  LV_ModifyCol(1, "Sort CaseLocale") ; or "Sort CaseLocale"
  LV_Modify(LV_GetCount(), "Vis") ;scrolls down
return

;---------------------------------------------------

;------------------- Modify Text -------------------------
Modify2:
  Fx=% F%Tabnumber%
  ifexist,%fx%
    filedelete,%fx%
  ControlGet,AA,List,,SysListView32%tabnumber%,%MainWindowTitle% ;<< the correct name of listview
  if aa<>
  {
    stringreplace,AA,AA,%t%,%delim%,all ;<< replaces TAB with Delimiter
    stringreplace,AA,AA,`n,`r`n,all
    ;msgbox, 262208, ,%aa%
    fileappend,%AA%,%fx%
    aa=
    return
  }
return
;------------------------------------------------------------------

;------------------- PRINT-Listview -------------------------
PrintLv1:
  FileTest=PrintList%Tabnumber%.txt
  ifexist,%filetest%
    filedelete,%filetest%
  ControlGet,AA,List,,SysListView32%tabnumber%,%MainWindowTitle% ;<< the correct name of listview
  if aa<>
  {
    stringreplace,AA,AA,%t%,%delim%,all ;<< replaces TAB with Delimiter
    stringreplace,AA,AA,`n,`r`n,all
    ;msgbox, 262208, ,%aa%
    fileappend,%AA%,%filetest%
    aa=
    run,%filetest%
    return
  }
return
;------------------------------------------------------------------

;--- create a testfile ---------------
; test1:
;   ;=delim = `,
;   F1 = JTECO PC.txt
;   F2 = JAPCO PC.txt
;   F3 = JAHACO PC.txt
;   F4 = JTECO-B PC.txt

;   P1 = JTECO PC
;   P2 = JAPCO PC
;   P3 = JAHACO PC
;   P4 = JTECO-B PC
; ifnotexist,%f1%
; {
;   e1=
;   Fileappend,%e1%`r`n,%f1%
; }
; e1 =
; return
;--- end create a testfile -----------

;================== END script ==========================================================
;---------------------

Pause::Pause
!^r::Reload