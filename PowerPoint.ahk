class PowerPoint{
    ExeName := ""
    PPO := ""

    __New(){
        base.__New()
        this.ExeName := this.p.getProperty("ppexe")
    }
    Run()
    {
        FileWrite("|POWERPOINT| run Application")
        this.PPO := ComObjCreate("PowerPoint.Application")
        this.PPO.Visible := true
        WinWait, ahk_class PPTFrameClass
        WinMaximize, ahk_class PPTFrameClass
        this.PPO.Presentations.Add()
        this.PPO.ActivePresentation.Slides.Add(1,12)
    }
    WriteInTextBox(str,posX,posY,width,height)
    {
        FileWrite("|POWERPOINT| write {" . str . "} in textbox")
        this.PPO.ActivePresentation.Slides(1).Shapes.AddTextbox(1,posX,posY,width,height).TextFrame.TextRange.Text := str
    }
    ;TODO : verif if file exist
    SaveAs(location,filename,formatstr)
    {

        savepoint := location . "\" . filename
		if(formatstr = "pdf" OR formatstr = "PDF" OR formatstr= "Pdf")
			format = 32
		if(formatstr = "pptx" OR formatstr = "PPTX" OR formatstr = "Pptx")
			format = 11
        
		this.PPO.ActivePresentation.SaveAs(savepoint, format)
        FileWrite("|POWERPOINT| saved file {" . savepoint . "." . formatstr . "}")

        Return savepoint . "." . formatstr
    }
    Close(save)
    {
        if(!save)
            this.PPO.ActivePresentation.Close()
        this.PPO.Quit()
       FileWrite("|POWERPOINT| close Application")
    }
}