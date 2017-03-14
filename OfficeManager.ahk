#Include Properties.ahk
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#Include Properties.ahk
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

class OfficeManager{
	Path := ""
	p := ""
	__New(){
		this.p := new Properties()
		this.Path := this.p.getProperty("officepath")
	}
}