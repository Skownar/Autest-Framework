;REGION vars
folderlocation := ""


;REGION Gui controls
Gui, Font, s15, Calibri
Gui, Add, Text, x10 y15 , Select your script folder 
Gui, Font
Gui, Add, Button , w70 x230 y15 gPickFolder, Pick
Gui, Font, s10, Calibri
Gui, Add, Text, vPath w425 x10, (Path)
Gui, Font, s15, Calibri
Gui, Add, Text,x10 , Name of your script :
Gui, Font
Gui, Add, Edit, w200 x210 y80 vNameScript , Name 
Gui, Add, Button, Default w200 x115 gGenerate, Generate Script files
Gui, Show , w430 h160 ,FileCreator


;REGION Gui events
PickFolder:
	FileSelectFolder, folderlocation
	GuiControl,, Path, %folderlocation%
Return

Generate:
	GuiControlGet, NameScript
	if(NameScript = "Name"){
		MsgBox, Name input must be completed 
	}
	else{
		Try {
			FileCreateDir, %folderlocation%\%NameScript%
			FileCreateDir, %folderlocation%\%NameScript%\Img
			StringLower, nameScript, NameScript
			FileAppend,,%folderlocation%\%NameScript%\%nameScript%.ahk
			FileAppend,,%folderlocation%\%NameScript%\%nameScript%_test.ahk			
			FileAppend,,%folderlocation%\%NameScript%\%nameScript%.log
			location := folderlocation . "\" . NameScript . "\"  . nameScript . ".ahk"
			MsgBox % location
			copyfile(location,nameScript)

			MsgBox, Files created !
			Run, %folderlocation%\%NameScript%\
		}
		Catch, e{
			MsgBox, % "Error : " e		
		}
	}
Return
copyfile(file,name){
	Loop,Read, sample.txt, file
	{
			IfInString, A_LoopReadLine, logfile
			{
				MsgBox % A_LoopReadLine
				FileAppend, global logfile = "%name%.log" `n , %file%
				continue
			}
			FileAppend, %A_LoopReadLine% `n, %file%
	}
}
GuiClose:
	ExitApp
;REGION Functions
