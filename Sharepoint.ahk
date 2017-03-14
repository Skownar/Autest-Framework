#Include Browser.ahk
#Include functions.ahk
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, RegEx

class Sharepoint extends Browser{
	
	__New(Page){
		base.__New(Page)
	}
	
	;; TODO : maybe remove ?
	openExplorerSPFrame(){
		found := false 
		base.WaitWB()
		Frames := this.Pwb.document.parentWindow.frames
		outer:
		Loop, % Frames.Length
		{
			Inputs := Frames[A_Index-1].document.getElementsByTagName("input") 
			If(Inputs.Length <> 0)
			{
				Loop, % Inputs.Length
				{
					if(Inputs[A_Index-1].type = "file"){
						Inputs[A_Index-1].click()
						found := true
						break outer
					}
				}
			}
		}
		break_outer:
		
		return found
	}
	
	uploadFile(dir = "", filename = ""){
		searchedImage := A_ScriptDir . "\browse_input.png"
		found := false
		while !found {
			found := clickSearchImg(searchedImage)
		}
		WinWait, Choose File to Upload
		WinActivate, Choose File to Upload 
		clickCtrlByClass("Choose File to Upload","ToolbarWindow322",0.99,0.5)
		Send %dir% {Enter}
		clickCtrlByClass("Choose File to Upload","Edit1",0.9,0.5)
		Send %filename%
		Sleep 500
		clickCtrlByClass("Choose File to Upload","Button1",0.5,0.5)
		searchedImage := A_ScriptDir . "\sharepoint_upload_okButton.png"		
		found := false
		while ! found {
			found := clickSearchImg(searchedImage)
		}
	}
}	