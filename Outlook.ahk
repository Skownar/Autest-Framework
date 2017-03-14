#Include OfficeManager.ahk
#Include IQP.ahk
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, RegEx

class Outlook extends OfficeManager{ 
	
	ExeName := ""
	Outlook := ""
	Mail := ""
	flag := false
	iqp := ""
	
	__New(){
		base.__New()
		this.ExeName := this.p.getProperty("outlookexe")
		this.iqp := new IQP() 
	}
	
	Run(){
		Process, Exist, % this.ExeName
		If(!ErrorLevel){
			Run, % this.ExeName, this.Path
			WinWait, ^.*@nbb.be.*$
			WinActivate, ^.*@nbb.be.*$
			this.flag := false
		}
		Else 
			this.flag := true
		this.Outlook := ComObjActive("Outlook.Application").GetNameSpace("MAPI").GetDefaultFolder(6)
		If(this.flag)
			WinShow, Microsoft Outlook
		
	}
	
	WriteMail(recepient = "", subject = "" , body = ""){
		this.Mail := ComObjActive("Outlook.Application").CreateItem(0)
		this.Mail.Recipients.Add(recepient)
		this.Mail.Subject := subject
		this.Mail.body := body
		this.Mail.display	
	}
	
	AddAttachment(location = "",filename = ""){
		path := location . "\" . filename
		this.Mail.Attachments.Add(path)
	}
	SendMail(){
		this.Mail.Send()
	}
	SendMailWithAttachment(){
		WinWait Message
		WinActivate Message
		Send ^{Enter}
		Send {Enter}
		WinWait Classification
		WinActivate Classification
		Send {Enter}
	}

	WaitForMailBySubject(box = "", subject = "")
	{
		Inbox := ComObjActive("Outlook.Application").GetNameSpace("MAPI").GetDefaultFolder(6)
		while Inbox.items(1).Subject <> subject{
			Sleep 10
			ToolTip wait or mail
		}
		this.Mail := Inbox.items(1)
	}

	BrowseMailBySubject(box = "" , subject = "")
	{
		this.flag := false
		if(box = "Inbox")
			Box := this.Outlook.folders(1)
		for Mail in Box.items {
			if (Mail.subject = subject){
				this.Mail := Mail
				this.flag := true
			}
		}
		if(!flag)
			MsgBox % subject . "not found"
	}
	
	SaveAttachement(newname = "", location = ""){
		path := location . "\" . newname
		MsgBox % path
		this.Mail.Attachments(1).SaveAsFile(path)
	}
}