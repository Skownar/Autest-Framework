class Word extends OfficeManager
{
	ExeName := ""
	oWord := ""
	flag := false
	__New(){
		base.__New()
		this.ExeName := this.p.getProperty("wordexe") ; p is define within parent class
	}

	;TODO if ! exist wordapplication -> return + warning
	Run(){
		Process, Exist, % this.ExeName
		If(!ErrorLevel)
		{
			this.oWord := ComObjCreate("Word.Application")
			this.oWord.Visible := true
			WinWait ahk_class OpusApp
			WinMaximize ahk_class OpusApp
			this.flag := false
			this.oWord.Documents.Add
			FileWrite("|WORD| word wasnt running")
		}
		Else
		{

			this.flag := true
			this.oWord := ComObjActive("Word.Application")
			this.oWord.Visible := true
			this.oWord.ActiveDocument.Content.Select
			WinActivate, ahk_class OpusApp
			WinShow, ahk_class OpusApp
			FileWrite("|WORD| word was running")
		}	
	}

	WriteText(text = ""){
		FileWrite("|WORD| Writing text in word :" . text)
		this.oWord.Selection.TypeText(text)
	}
	OpenClass(){
		this.oWord.CommandBars.ExecuteMso("Classification")
	}

	;TODO if exist file, return + writelog , else warning
	Save(location ="", filename = "", format = ""){
		FileWrite("|WORD| begin save file : " . location . "\" . filename . "." . format )
		savepoint := location . "\" . filename
		if(format = "pdf" OR format = "PDF" OR format = "Pdf")
			format = 17
		if(format = "doc" OR format = "docx" OR format = "Doc" OR format = "Docx")
			format = 16
		this.oWord.ActiveDocument.SaveAs2(savepoint,format,,,,,false)
		FileWrite("|WORD| file saved " . location . "\" . filename . "." . format)
	}
	
	Close(save := false){
		choice := 0 
		if(save)
			choice := -1
		this.oWord.Application.Quit(choice)
		FileWrite("|WORD| close Application")
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
