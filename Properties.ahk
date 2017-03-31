class Properties{
	
	propers := Object()
	propersFile := ""
	
	__New(pathPropertiesFile:=""){
		;MsgBox % pathPropertiesFile
		if(pathPropertiesFile = ""){
			this.propersFile := A_ScriptDir . "\..\..\Lib\Properties\application.properties"
		}
		else{
			this.propersFile := pathPropertiesFile
		}
		if(!FileExist(this.propersFile)){
			;TODO: write log 	
		}
		if(FileExist(this.propersFile))
		{
			this.loadProperty()
		}
	}
	
	loadProperty(){
		FileWrite("|PROPERTIES| loading ")
		Loop, read, % this.propersFile
		{
			Loop, parse, A_LoopReadLine, %A_Tab%
			{
				if ErrorLevel
					Break				
				tmpArray := StrSplit(A_LoopField,"=")
				this.propers[tmpArray[1]] := tmpArray[2]
			}
			
		}
	}
	getProperty(property){
		FileWrite("|PROPERTIES| getting property {" . property . "}" )
		For key,value in this.propers
		{
			if(key = property){
				FileWrite("|PROPERTIES| getting property {" . property . "} value {" . value . "}" )
				return value
				Break
			}
		}
	}
}