#NoEnv
; iSpring Publish Automator
; @author Matt Gleeson <mgleeson@studygroup.com.
; last updated: 15/04/2014
; note: there is minimal error checking in this thing and it will force overwrite 
; existing presentations and possibly do unspeakable things to your cat. 
; Ensure you have iSpring installed and you know what you're doing before 
; firing this off. 

Gui, Add, Button, x6 y10 w140 h30, &Load PowerPoint Files
;Gui, Add, Checkbox, vOverwrite, &Overwrite existing
Gui, Add, Button, x428 y10 w140 h30, &Process
Gui, 1:Add, ListBox, X6 Y70 R28 W564 +VScroll +HScroll vTCLIST
Gui, Show
return

ButtonLoadPowerPointFiles:
FileSelectFile, files, M3, , Select PowerPoint files to process:, PowerPoint (*.pptx)
If ( files="" OR ErrorLevel ) {
     MsgBox, The user pressed cancel.
     Return
}
FileList := ""
Loop, Parse, files, `n
{
    If ( A_Index = 1 ) 
    {
       Folder := A_LoopField
       Continue
    }         
    FileList .= ((FileList<>"") ? "|" : "" ) Folder "\" A_LoopField
}
GuiControl,,TCLIST, %FileList%
Return

ButtonProcess:
IfWinNotExist, ahk_class PPTFrameClass
{
    Run, POWERPNT
    sleep 1000 
    SendInput, ^w
}

Loop, Parse, FileList, |
{ 
    if A_LoopField =  ; Ignore the blank item at the end of the list.
       continue

    sleep 1000
    Runwait, %A_LoopField%
    SendInput, {Alt}{y}{y}{2}
     
    SetTitleMatchMode 1
     WinWaitActive, Publish Presentation,,5 ; wait for the iSpring publish window
     if Errorlevel
          {
               MsgBox, WinWait timed out waiting for Publish Presentation.
               Break
          }
    SendInput, {Enter} ; kick off the publshing
        sleep 2000

; OVERWRITE ALL THE FILESES!!  Couldn't be bothered with an overwrite all checkbox option, hang it all , overwrite the lot of them for now
     OrigSavedClip3 := ClipboardAll ; Saves Original Clipboard For After Script Restore.
    Clipboard :=             ; Erases Old Clipboard
     SendInput,  ^c ; capture to clipboard the contents of the confirm overwrite dialog - this is a hack as usual methods of window ident don't seem to work
     Sleep , 30
     if Clipboard contains Confirm Overwrite
     {
     SendInput, y
     }
     Clipboard := OrigSavedClip3    ; Restores Original Clipboard From Before Script Was Run.
          
     IfWinExist, Confirm Overwrite ; overwrite existing - this doesn't seem to work
    {
         WinActivate
        SendInput, {Enter}
    }
    
     SetTitleMatchMode 1
     WinWaitActive, Generating content,,5 ; look to make sure the publishing process has started pop a window and bomb if it hasn't because something has gone wrong...
     if Errorlevel
          {
               MsgBox, WinWait timed out waiting for Generating content.
               Break
          }
     
    WinWaitClose, Generating content ; wait for the progress window to close
    
    WinWaitActive, Presentation Preview,,5 ; confirm finished
     if Errorlevel
          {
               MsgBox, WinWait timed out waiting for Presentation Preview.
               Break
          }
     WinActivate, Presentation Preview
     WinClose, Presentation Preview
     
     WinActivate, ahk_class PPTFrameClass
    sleep 2000
    SendInput, ^w
     
     sleep 500
     IfWinExist, ahk_class NUIDialog ;  choose don't save on closing 
    {
         WinActivate
        SendInput, n
    }
     SendInput, {Alt}{h}
     
}

Return

MsgBox, Processing complete!
GuiClose:
ExitApp
