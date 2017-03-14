#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, RegEx
#Include Properties.ahk

class Browser{
	
	Pwb := ""
	Path := ""
	ExeName := ""
	__New(Page){
		p := new Properties()
		this.Path := this.p.getProperty("iepath")
		this.ExeName := this.p.getProperty("iexe")
		this.createIEObject(Page)
	}
	
	createIEObject(URL := ""){
		Process, Exist, iexplorer.exe
		If(!ErrorLevel){
			this.Pwb := ComObjCreate("InternetExplorer.Application")
		}
		this.Pwb := this.IEGet()
		this.Pwb.Visible := true
		this.Pwb.Navigate(URL)
		this.Pwb.WaitWB()
	}

IEGet(Name="")        ;Retrieve pointer to existing IE window/tab
{
    IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
        Name := ( Name="New Tab - Windows Internet Explorer" ) ? "about:Tabs"
        : RegExReplace( Name, " - (Windows|Microsoft) Internet Explorer" )
    For wb in ComObjCreate( "Shell.Application" ).Windows
        If ( wb.LocationName = Name ) && InStr( wb.FullName, "iexplore.exe" )
            Return wb
}

clickLink(text := ""){
		found := False
		this.WaitWB()
		Links := this  .Pwb.Document.Links ; collection of hyperlinks on the page
		Loop % Links.Length{ ; Loop through links  // TO FUNC
			If ((Link := Links[A_Index-1].InnerText) != "") { 
				(OuterHTML := Links[A_Index-1].OuterHTML) 
				Link:=StrReplace(Link,"`n")
				Link:=StrReplace(Link,"`r")
				Link:=StrReplace(Link,"`t")
				OuterHTML:=StrReplace(OuterHTML,"`n")
				OuterHTML:=StrReplace(OuterHTML,"`r")
				OuterHTML:=StrReplace(OuterHTML,"`t")
				Msg .= A_Index-1 A_Space A_Space A_Space Link A_Tab OuterHTML "`r`n" ; add items to the msg list
			}
			If (Links[A_Index-1].InnerText = text) {
				found := True
				Links[A_Index-1].click()				
				Break
			}
		}
		return found
	}
	
	WaitWB(){
		while this.Pwb.busy OR this.Pwb.ReadyState != 4 OR this.Pwb.Document.Readystate <> "Complete" {
			;Wait for page to load
			Sleep, 100
		}
	}
}