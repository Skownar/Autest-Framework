#Include Properties.ahk
#Include OfficeManager.ahk
#Include functions.ahk
class Word extends OfficeManager
{
	ExeName := ""
	oWord := ""
	flag := false
	__New(){
		base.__New()
		this.ExeName := this.p.getProperty("wordexe") ; p is define within parent class
	}
	Run(){
		Process, Exist, % this.ExeName
		If(!ErrorLevel){
			Run, % this.ExeName, this.Path
			WinWait ahk_class OpusApp
			WinMaximize ahk_class OpusApp
			this.flag := false
		}
		Else
			this.flag := true
		this.oWord := ComObjActive("Word.Application")
		If(this.flag)
			WinShow, Microsoft Word

	}
	WriteText(text = ""){
		this.oWord.Selection.TypeText(text)
	}
	OpenClass(){
		this.oWord.CommandBars.ExecuteMso("Classification")
	}
	
	Save(location ="", filename = "", format = ""){
		savepoint := location . "\" . filename 
		if(format = "pdf" OR format = "PDF" OR format = "Pdf")
			format = 17
		if(format = "doc" OR format = "docx" OR format = "Doc" OR format = "Docx")
			format = 16
		this.oWord.ActiveDocument.SaveAs2(savepoint,format,,,,,false)
		
	}
	
	Close(){
		SetKeyDelay , 200 
		WinKill , ahk_class OpusApp
		WinWait ahk_class NUIDialog,,,1
		Send {Right}{Enter}
		WinWaitClose, ahk_class OpusApp
	}
	GetAddin(){
		Addins := this.oWord.COMAddIns
		for Addin in Addins{
			;MsgBox % Addin.Description
			if (InStr(Addin.Description,"Secure")){
				pid := Addin.ProgId
				Addin.Connect := True
			}
			
		}
		
	}
}
