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
		this.Path := p.getProperty("iepath")
		this.ExeName := p.getProperty("iexe")
		this.createIEObject(Page)
	}
	
	createIEObject(URL := ""){
		MsgBox % this.Path this.ExeName
		Process, Exist, % this.ExeName
		If(!ErrorLevel){
			Run % this.Path . "" . this.ExeName
		}
		else{
			WinShow, Internet Explorer
		}
		this.Pwb := this.WBGet()
		this.Pwb.Visible := true
		this.Pwb.Navigate(URL)
		this.Pwb.WaitWB()
		
}



;************Pointer to Open IE Window******************
WBGet(WinTitle="ahk_class IEFrame", Svr#=1) {               ;// based on ComObjQuery docs
   static msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
        , IID := "{0002DF05-0000-0000-C000-000000000046}"   ;// IID_IWebBrowserApp
;//     , IID := "{332C4427-26CB-11D0-B483-00C04FD90119}"   ;// IID_IHTMLWindow2
   SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%

   if (ErrorLevel != "FAIL") {
      lResult:=ErrorLevel, VarSetCapacity(GUID,16,0)
      if DllCall("ole32\CLSIDFromString", "wstr","{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr",&GUID) >= 0 {
         DllCall("oleacc\ObjectFromLresult", "ptr",lResult, "ptr",&GUID, "ptr",0, "ptr*",pdoc)
         return ComObj(9,ComObjQuery(pdoc,IID,IID),1), ObjRelease(pdoc)
      }
   }
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