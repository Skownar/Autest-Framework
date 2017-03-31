class Sharepoint extends Browser{
	pathOkButton := ""
	pathBrowseButton := ""

	__New(page){
		base.__New(page)
		location = %A_ScriptDir%\..\..\Lib\Properties\images.properties
		p := new Properties(location)
		this.pathBrowseButton := p.getProperty("SpBrowseButton")
		this.pathOkButton  := p.getProperty("SpOkButton")
		this.waitingimage := p.getProperty("SpWaiting")
		FileWrite("|SHAREPOINT| connected to " . page)

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
	    FileWrite("|SHAREPOINT| try upload {" . dir . "\" . filename . "} ")

		searchedImage := A_ScriptDir . "\" . this.pathBrowseButton
		found := false
		while !found {
			Sleep 20
			found := findSearchImg(searchedImage)
			
		}
		if found 
		{
			FileWrite("|SHAREPOINT| click on Browse... button")
			clickSearchImg(searchedImage)
		}
		FileWrite("|SHAREPOINT| waiting for Explorer appears")	
		WinWait, Choose File to Upload
		WinActivate, Choose File to Upload 
		FileWrite("|SHAREPOINT| write {" . dir . "} in top bar")	
		clickCtrlByClass("Choose File to Upload","ToolbarWindow322",0.99,0.5)
		Send %dir% {Enter}
		FileWrite("|SHAREPOINT| write {" . filename . "} in down bar")
		clickCtrlByClass("Choose File to Upload","Edit1",0.9,0.5)
		Send %filename%
		Sleep 500
		FileWrite("|SHAREPOINT| click on OkButton")
		clickCtrlByClass("Choose File to Upload","Button1",0.5,0.5)
		Sleep 500
		Loop, 5 {
			Sleep 200
			Send {TAB}
		}	
		Sleep 200
		Send {Enter}
	}	
	WaitBetweenUploads(){
		searchedImage := A_ScriptDir . "\" . this.waitingimage
		found := false
		while !found {
			Sleep 200
			found := findSearchImg(searchedImage)
		}
		if(found){
			FileWrite("|SHAREPOINT| ready for other upload")
		}
	}
}