SetTitleMatchMode , RegEx

#Include Properties.ahk
#Include functions.ahk
#Include Explorer.Ahk
class IQP {
	currentClassification := ""
	pathPublicCurrent := ""
	pathInternalCurrent := ""
	pathConfidentialCurrent := ""
	pathSecretCurrent := ""
	pathSetPublic := ""
	pathSetInternal := ""
	pathSetConfidential := ""
	pathSetSecret := ""
	okbutton := ""
	cancelbutton := ""
	closedClassificationPanel := ""
	setJustification := ""
	WaitReady := ""
	p := ""

	__New(){
		location = %A_ScriptDir%\Properties\images.properties
		this.p := new Properties(location)
		this.pathPublicCurrent := this.p.getProperty("PublicCurrent")
		this.pathInternalCurrent := this.p.getProperty("InternalCurrent")
		this.pathConfidentialCurrent := this.p.getProperty("ConfidentialCurrent")
		this.pathSecretCurrent := this.p.getProperty("SecretCurrent")
		this.pathSetPublic := this.p.getProperty("SetPublic")
		this.pathSetInternal := this.p.getProperty("SetInternal")
		this.pathSetConfidential := this.p.getProperty("SetConfidential")
		this.pathSetSecret := this.p.getProperty("SetSecret")
		this.okbutton := this.p.getProperty("OkButton")
		this.cancelbutton := this.p.getProperty("CancelButton")
		this.closedClassificationPanel := this.p.getProperty("ClosedClassificationPanel")
		this.justificationNeeded := this.p.getProperty("JustificationNeeded")
		this.setJustification := this.p.getProperty("SetJustification")
		this.WaitReady := this.p.getProperty("WaitReady")
	}
	OpenClassificationWindow(file){
		oe := new FileExplorer()
		oe.OpenContext2(file,"Classification")
		MouseMove, 0,0
		this.WaitIsReady()
		
	}
	OpenProtectionWindows(file){
		;TO IMPLEMENTS
	}
	GetCurrentClassification(){
	
		if findSearchImg(A_ScriptDir . "\" . this.pathPublicCurrent)
			this.currentClassification := "public"
		else if findSearchImg(A_ScriptDir . "\" . this.pathInternalCurrent)	
			this.currentClassification := "internal"
		else if findSearchImg(A_ScriptDir . "\" . this.pathConfidentialCurrent)
			this.currentClassification := "confidential"
		else if findSearchImg(A_ScriptDir . "\" . this.pathSecretCurrent)		
			this.currentClassification := "secret"
		else
			this.currentClassification := "no classification"
		return this.currentClassification
		
	}

	GetCurrentClassificationSP(){
		found := false
		searchedImage = %A_ScriptDir%\classification_title.png
		while !found {
			found := findSearchImg(searchedImage)
		}
		return this.GetCurrentClassification()
	}
	
	;TODO : Verify if Classification panel is closed 
	SetClassification(classification){
		this.GetCurrentClassification()

		if (classification = this.currentClassification)
			MsgBox,,, % "Classification already :" this.currentClassification

		else{
			;click on classification panel if it's closed
			if (findSearchImg(A_ScriptDir . "\" . this.closedClassificationPanel)) {
				clickSearchImg(A_ScriptDir . "\" . this.closedClassificationPanel)
			}
			;click on specified radio button
			if (classification = "public")
				searchedImage := A_ScriptDir . "\" . this.pathSetPublic
			else if (classification = "internal")
				searchedImage := A_ScriptDir . "\" . this.pathSetInternal
			else if (classification = "confidential")
				searchedImage := A_ScriptDir . "\" . this.pathSetConfidential
			else if (classification = "secret")
				searchedImage := A_ScriptDir . "\" . this.pathSetSecret
			else
				MsgBox Classification missing
			clickSearchImg(searchedImage)
		}
	}
	
	CommitWithJustification(){
		;click ok button
		clickSearchImg(A_ScriptDir . "\" . this.okbutton)
		;Verify if justification asked
		cpt := 1000
		found := false
		while !found {
			found := findSearchImg(A_ScriptDir . "\" . this.JustificationNeeded)
			IfWinNotExist, Classification
				break
		}
		;set fictive justification
		if(found){
			while !(clickSearchImg(A_ScriptDir . "\" . this.setJustification)){
				Sleep 20
			}
		}
		clickSearchImg(A_ScriptDir . "\" . this.okbutton)
	}

	Commit(){
		clickSearchImg(A_ScriptDir . "\" . this.okbutton)
		WinWaitClose, Classification
	}

	CommitSharePoint(){	
		Sleep 200
		Send {Enter}
		WinWaitClose, Classification
	}

	WaitIsReady(){
		WinWait Classification
		WinActivate Classification
		found := false
		searchedImage := this.WaitReady
		while !found {
			found := findSearchImg(searchedImage)
			Sleep 100
		}
	}
}