#Include Sharepoint.ahk
#Include functions.ahk
#Include Explorer.ahk
#Include IQP.ahk
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
logfile := FileOpen("useBrowser.log", "a")
SetTitleMatchMode RegEx
SetKeyDelay, 150,200


sp := new Sharepoint("http://dc.uat.nbb/SitePages/Home.aspx")
oe := new FileExplorer()
oiqp := new IQP()

WinWait ^.*Internet Explorer.*$
WinActivate ^.*Internet Explorer.*$
WinMaximize ^.*Internet Explorer.*$
sp.clickLink("tests-2016")
sp.clickLink("nbb-public")
sp.clickLink("Add document")
sp.uploadFile(A_ScriptDir "\Documents","noclass.pdf")

found := false
while !found  {
	Sleep 100
;	searchedImage = %A_ScriptDir%\addocument_sharepoint_icon.png
	;found := findSearchImg(searchedImage)
	found := sp.clickLink("Add document")
}

;; Docx
sp.uploadFile(A_ScriptDir "\Documents","noclass.docx")
found := False
while found <> True {
	Sleep 100
	searchedImage = %A_ScriptDir%\addocument_sharepoint_icon.png
	found := findSearchImg(searchedImage)
}

/*
 DOCX
*/
oiqp.OpenClassificationWindow("\\dc.uat.nbb\DavWWWRoot\tests2016\nbb-public\noclass.docx")
classif := oiqp.GetCurrentClassification()
MsgBox % classif
if(classif = "public"){
	FileWriteSuccess(logfile,"noclass.docx classification is ok")
}
else{
	FileWriteError(logfile,"noclass.docx classification is wrong")
}
oiqp.CommitSharePoint()

/*
PDF
*/

oiqp.OpenClassificationWindow("\\dc.uat.nbb\DavWWWRoot\tests2016\nbb-public\noclass.pdf")
classif := oiqp.GetCurrentClassification()
MsgBox % classif
if(classif = "public"){
	FileWriteSuccess(logfile,"noclass.pdf classification is ok")
}
else{
	FileWriteError(logfile,"noclass.pdf classification is wrong")
}
oiqp.CommitSharePoint()


^µ::
ExitApp

