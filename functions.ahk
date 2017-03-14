global debug := true

FileWriteSuccess(file,message){
	FormatTime, CurrentTime,, Time 
	FormatTime, CurrentDate,, ShortDate	
	file.WriteLine("[Success]" . CurrentDate . " " . CurrentTime . " : " . message . "`r`n")
}

FileWriteError(file,message){
	FormatTime, CurrentTime,, Time 
	FormatTime, CurrentDate,, ShortDate	
	file.WriteLine("[Error]" . CurrentDate . " " . CurrentTime . " : " . message . "`r`n")	
}

clickCtrlByClass(win = "",controle = "",posX=0,posY=0){
	CoordMode, Mouse
	;Loop{
	WinGetPos, X, Y, Width, Height, %win%
	ControlGetPos, ctrlX, ctrlY, ctrlWidth, ctrlHeight, %controle%, %win%
	;MsgBox % "winX = " X "WinY = " Y "ctrlX = " ctrlX "ctrlY = " ctrlY 
		;MouseGetPos, OutputVarX, OutputVarY 
	 	;ToolTip, % "WinX" X "WinY" Y "CtrlX" ctrlX "CtrlY" ctrlY, OutputVarX, OutputVarY 
		;Sleep, 50
	MouseClick, Left, X+ctrlX+(ctrlWidth*posX), Y+ctrlY+(ctrlHeight*posY)
	Sleep 500
	;}
}


ClickWinClassPosPix(winclass,posPixX,posPixY){
	SetTitleMatchMode, RegEx
	CoordMode, Mouse
	WinGetPos, X, Y, Width, Height, %winclass%
	PosX :=  X + posPixX
	PosY :=  Y + posPixY
	;MsgBox, %  " Win X :" X " Width :" Width " PosX : " PosX " Win Y : " Y " Height : " Height " PosY : " PosY
	;MsgBox, % PosX PosY
	MouseClick, Left,PosX,PosY
	
}

ClickWinClassPosPercent(winclass,posPercentX,posPercentY){
	CoordMode Mouse
	WinGetPos, X, Y, Width, Height, %winclass%
	PosX :=  X + (Width * posPercentX)
	PosY :=  Y + (Height * posPercentY)
	MsgBox, %  " Win X :" X " Width :" Width " PosX : " PosX " Win Y : " Y " Height : " Height " PosY : " PosY
	Sleep, 500
	MouseClick,Left,PosX,PosY
}

findSearchImg(searchedImage,button = 0){
	Width := A_ScreenWidth * 2
	Height := A_ScreenHeight
	CoordMode, Pixel
	ImageSearch, OutputVarX, OutputVarY, 0, 0, Width, Height, %searchedImage%
	If ErrorLevel = 2
		found := False
	If ErrorLevel = 1
		found := False
	If ErrorLevel = 0
		found := True
	Return, found
}


ClickSearchImg(searchedImage,button = 0){
	Width := A_ScreenWidth * 2
	Height := A_ScreenHeight
	;MsgBox, % Width Height
	CoordMode, Pixel
	ImageSearch, OutputVarX, OutputVarY, 0, 0, Width, Height, %searchedImage%
	If ErrorLevel = 2
		found := False
	If ErrorLevel = 1
		found := False
	If ErrorLevel = 0
		found := True
	CoordMode, Mouse
	MouseClick, Left, OutputVarX, OutputVarY
	Sleep, 500
	Return, found
}

ForcedClickSearchImg(winPosX, WinPosY, posClickX,posClickY,searchedImage){
	Width := A_ScreenWidth * 2
	Height := A_ScreenHeight
	MouseClick, Left, winPosX+=posClickX+=8, winPosY+=posClickY
	CoordMode, Pixel
	ImageSearch, OutputVarX, OutputVarY, 0, 0, Width, Height, %searchedImage%
	If ErrorLevel = 2
		found := False
	If ErrorLevel = 1
		found := False
	If ErrorLevel = 0
		found := True
	Return found
}

ClickSearchImgR(searchedImage){
	Width := A_ScreenWidth * 2
	Height := A_ScreenHeight
	;MsgBox, % Width Height
	CoordMode, Pixel
	ImageSearch, OutputVarX, OutputVarY, 0, 0, Width, Height, %searchedImage%
	If ErrorLevel = 2
		found := False
	If ErrorLevel = 1
		found := False
	If ErrorLevel = 0
		found := True
	CoordMode, Mouse
	MouseClick, Right, OutputVarX, OutputVarY
	Sleep, 500
	Return found
}



