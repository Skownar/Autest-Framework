class FileExplorer {
	
	Run(path = ""){
		Run , % "explorer.exe " , path
	}
	
	KillAll(){
		Run taskkill /f /im explorer.exe
	}
	
	Reload(){
		Run, explorer.exe
	}
	
	ResetExplorer(){
		this.KillAll()
		this.Reload()
	}
	
	OpenContextualOnSelectedFile(path = "" , file = ""){
		FileWrite("|EXPLORER| try opening file {" . path . "\" . file . "}")
		SetTitleMatchMode RegEx 
		selection = %path%\%file%
		MsgBox % A_Desktop
		Run, % "explorer.exe /select`," selection 
		WinWait , ahk_class Cabi
		WinMaximize, ahk_class Cabi
		WinActivate, ahk_class Cabi
		Send +{F10}
	}
	
	OpenContext2(path, menu, validate=True){
	FileWrite("|EXPLORER| try opening file {" . path . "}  menu " . menu)
	;by A_Samurai
	;v 1.0.1 http://sites.google.com/site/ahkref/custom-functions/invokeverb
    objShell := ComObjCreate("Shell.Application")
    if InStr(FileExist(path), "D") || InStr(path, "::{") {
        objFolder := objShell.NameSpace(path)   
        objFolderItem := objFolder.Self
    } else {
        SplitPath, path, name, dir
        objFolder := objShell.NameSpace(dir)
        objFolderItem := objFolder.ParseName(name)
    }
    if validate {
        colVerbs := objFolderItem.Verbs   
        loop % colVerbs.Count {
            verb := colVerbs.Item(A_Index - 1)
            retMenu := verb.name
            StringReplace, retMenu, retMenu, &       
            if (retMenu = menu) {
                verb.DoIt
				if(menu = "Classification"){
					SetTitleMatchMode, RegEx
					WinWait Classification
					WinActivate Classification
				}
				
                Return True
            }
        }
        Return False
    } else{
        objFolderItem.InvokeVerbEx(Menu)
		
	}
	}
}

