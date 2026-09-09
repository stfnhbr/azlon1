#Requires AutoHotkey v2.0
#SingleInstance Force
; Four hotkeys. Three of them fire only while Audacity is focused:
;
;   Ctrl+Shift+Alt+K   dump the project's labels to an .xlsx
;   Ctrl+Shift+Alt+I   import a sound-nouns .txt as one label track per category
;   Ctrl+Shift+Alt+N   make that .txt from a workbook and import it, in one go
;
; "Audacity" means either of the two installed here - 3.7.9 (Audacity.exe) and
; 4.0 (Audacity4.exe). The same three keys work in both. Which one you pressed
; them in is passed to the script as -Version, because 4.0 ships no scripting
; interface and has to be reached through its project file instead; the scripts
; explain the difference. 3.7.9 still goes over mod-script-pipe exactly as it
; always did, so nothing about the old way of working has moved.
;
; The one thing that differs in 4.0: Ctrl+Shift+Alt+I and Ctrl+Shift+Alt+N
; write into the project file, which Audacity must not be holding open. Save
; and close the project in 4.0 first.
;
; and one fires anywhere, because the window it opens is meant to stand beside
; the browser you are typing into, not beside Audacity:
;
;   Ctrl+Alt+/         open Caption Filler III on an open workbook
;
; Double-click this file to arm them. To have them armed at every login,
; press Win+R, run "shell:startup", and put a shortcut to this file there.
;
; To use different keys, edit the ^+!k, ^+!i, ^+!n and ^!/ lines below:
;   ^ = Ctrl,  + = Shift,  ! = Alt.  e.g. ^+q is Ctrl+Shift+Q.

TraySetIcon("shell32.dll", 25)
A_IconTip := "Audacity annotations - Ctrl+Shift+Alt+K export, I import, N nouns"
           . "`nAudacity 3.7.9 and 4.0"
           . "`nCtrl+Alt+/ Caption Filler III"

/*  WinExist below looks for a window whose title STARTS WITH the filler's, which
    is the whole of it up to the workbook and sheet it bound to. Said outright
    rather than left to the default, which matches the text anywhere in the
    title and would just as happily find a browser tab or an explorer window
    that happened to mention the name.  */
SetTitleMatchMode 1

RunAnnotationScript(name, arguments := "") {
    script := A_ScriptDir "\" name
    if !FileExist(script) {
        MsgBox("Cannot find:`n" script, "Audacity annotations", "Iconx")
        return
    }
    Run 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' script '" ' arguments,
        A_ScriptDir, "Hide"
}

/*  Which Audacity the key was pressed in.

    Passed on as -Version so the script never has to guess. It can work this
    out by itself from whichever window holds the foreground, but by the time
    PowerShell has started the foreground may be something else entirely - and
    with both versions open at once, guessing wrong sends the labels to the
    wrong program. The window that was active when the key went down is the
    only reliable answer, and only AutoHotkey is in a position to know it.  */
AudacityVersion() {
    return WinActive("ahk_exe Audacity4.exe") ? "4" : "3"
}

/*  Opens the caption guide, or raises the one already open.

    Raises rather than restarts, because CaptionFiller3.ahk is #SingleInstance
    Force: starting it a second time would tear the open window down, ask again
    which sheet to bind, and lose the check marks against the fields already
    typed into the form you are in the middle of.

    Raises rather than activates. The guide carries WS_EX_NOACTIVATE so that
    clicking it never takes focus off the website field it is about to type
    into, and WinActivate on such a window does nothing useful; WinMoveTop puts
    it in front of the other always-on-top windows without touching focus, which
    is the only thing worth doing to a window that is already visible.

    An .ahk, so it goes through the file association rather than through
    powershell.exe like the three above - the same thing double-clicking it in
    Explorer does.  */
OpenCaptionFiller3() {
    if WinExist("AutoNyx Caption Filler III") {
        WinMoveTop()
        return
    }
    script := A_ScriptDir "\CaptionFiller3.ahk"
    if !FileExist(script) {
        MsgBox("Cannot find:`n" script, "Audacity annotations", "Iconx")
        return
    }
    try
        Run '"' script '"', A_ScriptDir
    catch
        MsgBox("Could not start:`n" script "`n`nAutoHotkey v2 has to be the "
             . "program that opens .ahk files.", "Audacity annotations", "Iconx")
}

#HotIf WinActive("ahk_exe Audacity.exe") || WinActive("ahk_exe Audacity4.exe")
^+!k:: RunAnnotationScript("ExportAnnotations2.ps1",  "-Version " AudacityVersion())
^+!i:: RunAnnotationScript("MakeSoundNouns2.ps1",     "-Version " AudacityVersion() " -ImportOnly")
^+!n:: RunAnnotationScript("MakeSoundNouns2.ps1",     "-Version " AudacityVersion())
#HotIf

/*  No #HotIf around this one: the guide is opened while a browser is in front,
    and a key that only worked inside Audacity could never open it.

    "/" is the character, so this follows the keyboard layout: on the US layout
    in use here it is the unshifted key left of the right Shift, which is what
    Ctrl+Alt+/ means. On the Danish, Norwegian and Swedish layouts also
    installed here "/" is Shift+7, so the hotkey would want Shift as well - to
    have the same physical key under all four, write ^!SC035 instead of ^!/.  */
^!/:: OpenCaptionFiller3()
