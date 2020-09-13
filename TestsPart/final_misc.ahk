;SETUP
SetTitleMatchMode, RegEx
#Include Word.ahk
#Include Outlook.ahk
#Include Explorer.ahk
#Include IQP.ahk
#Include Properties.ahk
;FIELDS
ow := new Word()
oo := new Outlook()
oe := new FileExplorer()
oiqp := new IQP()
op := new Properties()
;SCRIPTS

ow.Run()
ow.Save(A_Desktop, "test.pdf", "pdf")
ow.Close()

file = %A_Desktop%\test.pdf
oiqp.OpenClassificationWindow(file)
oiqp.SetClassification("confidential")
oiqp.CommitWithJustification()

oo.Run()

address := op.getProperty("myaddress")
oo.WriteMail(address,"test mail", "hello")

oo.AddAttachment(A_Desktop,"test.pdf")
oo.SendMailWithAttachment()

oo.WaitForMailBySubject("Inbox","test mail")
oo.SaveAttachement("test_saved.pdf",A_Desktop)

file = %A_ScriptDir%\test_saved.pdf
oe.OpenContext2(file,"Open")
Sleep 2000
WinClose test_saved.pdf
WinWaitClose test_saved.pdf

oiqp.OpenClassificationWindow(file)
classif := oiqp.GetCurrentClassification()
MsgBox % classif
oiqp.Commit()

