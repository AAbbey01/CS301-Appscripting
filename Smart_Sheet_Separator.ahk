#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

Esc::ExitApp ; Exit script with Escape key

; Runs the Macro to duplicate sheets
; Must add authorization before running
;enter, 2tab enter, 7tab enter
^1::
;Authorizing the SmartSheetMaker2 Macro in Sheets
Send !+n
Sleep 100
Send {Down}
Sleep 100
Send {Down}
Sleep 100
Send {Enter}
Sleep 100
Send {Down}
Sleep 100
Send {Down}
Sleep 100
Send {Enter}
Sleep 1000
Send {Enter}
Sleep 500
Send {Tab}
Sleep 200
Send {Tab}
Sleep 200
Send {Enter}
Sleep, 1000
Loop, 7
{
    Send {Tab}
    Sleep 100
}
Send {Enter}
Sleep 200

Send !+n
Sleep 100
Send {Down}
Sleep 100
Send {Down}
Sleep 100
Send {Enter}
Sleep 100
Send {Down}
Sleep 100
Send {Down}
Sleep 100
Send {Enter}
Sleep 100
Loop, 3
{
    Send {Down}
    Sleep 100
}
Send {Enter}
Sleep 45000
return


; Seperates the Sheets
^2::
CoordMode, Mouse, Screen
MouseMove, A_ScreenHeight, A_ScreenWidth
Send ^{Home}
Sleep 150
Send {Right}
Sleep 250   
Send {Right}
Sleep 200
Send, ^c
Sleep 200
myVar := Clipboard
Sleep, 200
Loop, %myVar%
{
Sleep 200
; Open list of spreadsheets
Send !+k
Sleep, 200
; Go to second one (first in list of smartsheets)
Send {Down}
Sleep, 200
Send {Enter}
Sleep, 200
;Alt+shift+s = current tab
Send !+s
Sleep, 200
;Make a copy
Loop, 3
{
    Send {Down}
    Sleep 200
}
Send {Enter}
Sleep, 200
Send {Enter}
Sleep, 3000
;Open the Copied Sheet
Send {Tab}
Sleep, 200
Send {Tab}
Sleep, 200
Send {Enter}
;Allow Perms on the New Sheet
Sleep, 3759
Send ^{Home}
Sleep, 300
Send {Tab}
Sleep, 300
Send {Enter}
Sleep, 300
;Get the Student's Name
Send {Down}
Sleep 1300
Send {Right}
Sleep 400
Send {Right}
Sleep 400
Send {Right}
Sleep 400
Send ^c
Sleep 800
; Name the spreadsheet
Send !1
Sleep 1500
Send %Clipboard%
Sleep, 400
Send {Enter}
Sleep, 400
;Return to get the email and Hide the row // Not The Tag version
Send ^{Home}
Sleep, 400
Send {Down}
Sleep, 400
Send ^c
ClipWait, [ SecondsToWait, 1]
Sleep, 400
Send ^!0
Sleep, 400
CoordMode, Mouse, Screen
MouseMove, A_ScreenWidth/2, A_ScreenHeight/2
 Click
Sleep, 400
Send ^{Home}

Sleep, 800
;Share Spreadsheet
Send !/     
Sleep, 200
Send Share
Sleep, 200
Send {Enter}
Sleep, 3000
Send %Clipboard%
Sleep, 4000
Send, {Enter}
Sleep, 400
;Tab Enter enter
Send {Tab}
Sleep, 400
Send {Enter}
Sleep, 400
Send {Enter}
Sleep, 400
Loop, 6
{
    Send {Tab}
    Sleep, 200
}
Send {Enter}
Sleep, 2500
;Return to original sheet
Send ^+{Tab}
Sleep, 200
Send {Tab}
Sleep, 200
Send {Enter}
Sleep, 200
;Now Delete this sheet from original book
Send !+s
Sleep, 200
Send {Down}
Sleep, 200
Send {Enter}
Sleep, 200
Send {Enter}
Sleep 2000
Sleep 2000
}
return
