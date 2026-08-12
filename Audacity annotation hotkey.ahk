#Requires AutoHotkey v2.0
#SingleInstance Force
; Three hotkeys, all only while Audacity is focused:
;
;   Ctrl+Shift+Alt+K   dump the project's labels to an .xlsx
;   Ctrl+Shift+Alt+I   import a sound-nouns .txt as one label track per category
;   Ctrl+Shift+Alt+N   make that .txt from a workbook and import it, in one go
;
; Double-click this file to arm them. To have them armed at every login,
; press Win+R, run "shell:startup", and put a shortcut to this file there.
;
; To use different keys, edit the ^+!k, ^+!i and ^+!n lines below:
;   ^ = Ctrl,  + = Shift,  ! = Alt.  e.g. ^+q is Ctrl+Shift+Q.

TraySetIcon("shell32.dll", 25)
A_IconTip := "Audacity annotations - Ctrl+Shift+Alt+K export, I import, N nouns"

RunAnnotationScript(name) {
    script := A_ScriptDir "\" name
    if !FileExist(script) {
        MsgBox("Cannot find:`n" script, "Audacity annotations", "Iconx")
        return
    }
    Run 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' script '"',
        A_ScriptDir, "Hide"
}

#HotIf WinActive("ahk_exe Audacity.exe")
^+!k:: RunAnnotationScript("ExportAnnotations.ps1")
^+!i:: RunAnnotationScript("ImportSoundNouns.ps1")
^+!n:: RunAnnotationScript("MakeSoundNouns.ps1")
#HotIf
