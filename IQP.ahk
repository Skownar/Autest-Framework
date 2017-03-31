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
		;TODO : verif if iqp exist
	__New(){
		
		location = %A_ScriptDir%\..\..\Lib\Properties\images.properties
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
	;TODO : verif if file exist
	OpenClassificationWindow(file){
		FileWrite("|IQP| try opening IQP classification on {" . file . "}")
		oe := new FileExplorer()
		oe.OpenContext2(file,"Classification")
		MouseMove, 0,0
		this.WaitIsReady()

	}
	OpenProtectionWindows(file){
		;TO IMPLEMENTS
	}
	GetCurrentClassification(){
		FileWrite("|IQP| looking for the current classification")
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
		FileWrite("|IQP| current classification is {" . this.currentClassification . "}")
		return this.currentClassification
		
	}

	
	;TODO : Verify if Classification panel is closed 
	SetClassification(classification){
		this.GetCurrentClassification()

		if (classification = this.currentClassification)
		{
			FileWrite("|IQP| classification was already {" . this.currentClassification . "}")
		}
		else{
			;click on classification panel if it's closed
			if (findSearchImg(A_ScriptDir . "\" . this.closedClassificationPanel)) {
				FileWrite("|IQP| opening classification panel")
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
				classification := "not found"
			FileWrite("|IQP| setting classification {" . classification . "}")
			clickSearchImg(searchedImage)
		}
	}
	
	Commit(){
		FileWrite("|IQP| commit classification ")
		;click ok button
		clickSearchImg(A_ScriptDir . "\" . this.okbutton)
		;Verify if justification asked
		found := false
		FileWrite("|IQP| looking if justification needed")
		while !found {
			found := findSearchImg(A_ScriptDir . "\" . this.JustificationNeeded)
			IfWinNotExist, Classification
			{
				FileWrite("|IQP| no justification required")
				break
			}
		}
		;set fictive justification
		if(found){
			while !(clickSearchImg(A_ScriptDir . "\" . this.setJustification)){
				Sleep 20
			}
			FileWrite("|IQP| set fictive justification")
		}
		clickSearchImg(A_ScriptDir . "\" . this.okbutton)
		WinWaitClose, Classification
		FileWrite("|IQP| Close Application")
	}

	CommitSharePoint(){	
		Sleep 200
		Send {Enter}
		WinWaitClose, Classification
	}
	Close(){
		FileWrite("|IQP| Close Application")
		Send {ESC}
	}
	WaitIsReady(){
		FileWrite("|IQP| Wait classification is ready")
		WinWait Classification
		WinActivate Classification
		found := false
		while !found {
			found := findSearchImg(A_ScriptDir . "\" . this.WaitReady)
			Sleep 100
		}
	}
}