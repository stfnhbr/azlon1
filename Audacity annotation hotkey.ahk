#Requires AutoHotkey v2.0
#SingleInstance Force
; Ctrl+Shift+Alt+K inside Audacity: dump the project's labels to an .xlsx.
;
; Double-click this file to arm the hotkey. To have it armed at every login,
; press Win+R, run "shell:startup", and put a shortcut to this file there.
;
; To use a different key, edit the ^+!k line below:
;   ^ = Ctrl,  + = Shift,  ! = Alt.  e.g. ^+q is Ctrl+Shift+Q.

TraySetIcon("shell32.dll", 25)
A_IconTip := "Audacity annotations - Ctrl+Shift+Alt+K"

#HotIf WinActive("ahk_exe Audacity.exe")
^+!k:: {
    script := A_ScriptDir "\ExportAnnotations.ps1"
    if !FileExist(script) {
        MsgBox("Cannot find:`n" script, "Audacity annotations", "Iconx")
        return
    }
    Run 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' script '"',
        A_ScriptDir, "Hide"
}
#HotIf
