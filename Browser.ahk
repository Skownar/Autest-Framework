DetectHiddenWindows, On
class Browser{
	
	Pwb := ""
	Path := ""
	ExeName := ""
	Flag := ""
	
	__New(Page){
		FileWrite("|INTERNET EXPLORER| initialization ")
		p := new Properties()
		this.Path := p.getProperty("iepath")
		this.ExeName := p.getProperty("iexe")
		this.createIEObject(Page)
	}

	createIEObject(URL := ""){
		FileWrite("|INTERNET EXPLORER| running D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E (medium) mode")
		this.Pwb := ComObjCreate("{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")   ;// Create an IE object
		this.Pwb.Visible := true                                  ;// Make the IE object visible              
		loading := true            ;// Set the variable "loading" as TRUE
		this.Pwb.Navigate(URL)
		this.WaitWB()
		WinMaximize, ahk_class IEFrame
	}

	Navigate(page)
	{
		If(page = "")
			FileWrite("|INTERNET EXPLORER| page values {empty}")
		else
			FileWrite("|INTERNET EXPLORER| page values {" . page . "}")
		Return This.Pwb.Navigate(page)
	}

	;************Pointer to Open IE Window******************
	WBGet(WinTitle="ahk_class IEFrame", Svr#=1) {               ;// based on ComObjQuery docs
   		static msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
        , IID := "{0002DF05-0000-0000-C000-000000000046}"   ;// IID_IWebBrowserApp
	    ;, IID := "{332C4427-26CB-11D0-B483-00C04FD90119}"   ;// IID_IHTMLWindow2
   		SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%
   		if (ErrorLevel != "FAIL") {
      		lResult:=ErrorLevel, VarSetCapacity(GUID,16,0)
      		if DllCall("ole32\CLSIDFromString", "wstr","{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr",&GUID) >= 0 {
         		DllCall("oleacc\ObjectFromLresult", "ptr",lResult, "ptr",&GUID, "ptr",0, "ptr*",pdoc)
         		return ComObj(9,ComObjQuery(pdoc,IID,IID),1), ObjRelease(pdoc)
      		}
   		}
	}

	WaitWB(){
		while this.Pwb.Busy or this.Pwb.ReadyState != 4 ;Wait for page to load
			Sleep, 100
	}

	clickLink(text := ""){
		found := False
		this.WaitWB()
		Links := this.Pwb.Document.Links ; collection of hyperlinks on the page
		FileWrite("|INTERNET EXPLORER| looking for link with text {" . text . "}")
		while !found {
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
				FileWrite("|INTERNET EXPLORER| link  {" . text . "} found")
				Links[A_Index-1].click()				
				Break
			}
			else{
				found := False
				FileWrite("|INTERNET EXPLORER| link with text {" . text . "} not found")
			}
		}
		}
		return found
	}

	getInputById(id := "")
	{
		input := This.Pwb.Document.getElementById(id)
		if(!input.tagName := "input")
			return False
		else
			return new this.Input(input)
	}

	getSpanById(id := "")
	{
		span := this.Pwb.Document.getElementById(id)
		if(!input.tagName := "input")
			return False
		else
			return new this.Span(span)
	}

	getSpanByText(text := "",classToExclude := "")
	{
		this.WaitWB()
		Sleep 100
		excludedFound := false
		spanToReturn := ""
		spans := this.Pwb.Document.getElementsByTagName("span")
		;MsgBox % spans.Length

		Loop % spans.Length
		{
			excludedFound := false
			;Msgbox % spans[A_Index-1].InnerHtml
			if(spans[A_Index-1].InnerHtml = text)
			{
				;Msgbox % spans[A_Index-1].InnerHtml
				if(classToExclude != "")
				{
					cll := spans[A_Index-1].classList
					Loop % cll.Length
					{
						if(cll.item[A_Index-1] = classToExclude)
						{
							excludedFound := true
							Break
						}
					}
					if(!excludedFound)
					{
						cll := spans[A_Index-1].parentElement.classList
						Loop % cll.Length
						{
							if(cll.item[A_Index-1] = classToExclude)
							{
								excludedFound := true
								Break
							}
						}
					}
					
				}
				if(classToExclude = "" OR excludedFound = false)
				{
					excludedFound := false
					span := spans[A_Index-1]
					Break
				}
			}
			else{
				Continue
			}
		}
		return new this.Span(span)

	}	
		

	getListById(id := "")
	{
		list := this.Pwb.getElementById(id)
		if(list.tagName := "ol" OR list.tagName := "ul")
			return new this.List()
		else
			return false
	}

	class Input{
		input := ""
		type := ""

		__New(input)
		{
			this.input := input
			this.type := input.type
		}

		setValue(values := "")
		{
			if(this.type != "checkbox" AND this.type != "radio")
			{
				this.input.value := values
			}
			else
			{
				MsgBox cannot applied on checkbox or radiobutton
				return false
			}
		}

		setChecked(chk := false)
		{
			if(this.type = "checkbox" OR this.type = "radio")
			{
				if(chk = true OR chk = false)
				{
					this.input.checked := chk
					return true
				}
				else
				{
					Msgbox value {%chk%} not allowed, must be true or false
					return false
				}
			}
		}
	}

	class Span
	{
		span := ""
		text := ""
		
		__New(val)
		{
			this.span := val
			;MsgBox % this.span.InnerHtml
		}

		getText()
		{
			return this.span.InnerText
		}

		click()
		{
			return this.span.click()
		}
	}

	; Tag <ul> or <ol>
	class List
	{
		list := ""
		children := ""
	}
}
