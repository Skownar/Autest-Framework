class DotStat extends Browser{

    __New(page = "")
    {
        base.__New(page)
    }
    CheckLanguage()
    {
        this.WaitWB()
        lg := ""
        found := false
        spans := This.Pwb.Document.getElementsByTagName("span")
        Loop % spans.Length
        {
            cll := spans[A_Index-1].classList
            Loop % cll.Length
            {
                if(cll.item[A_Index-1] = "currentlanguagelink") 
                { 
                    found := true
                    Break
                }
            }
            if(found)
            {
                lg := spans[A_Index-1].InnerText
                Break
            }
        }
        FileWrite("|DOTSTAT| current language is {" . lg . "}")
        img := this.Pwb.Document.getElementById("_ctl0__ctl0_Logo2")
        src := img.src
        FileWrite("|DOTSTAT| current logo source is {" . src . "}")
        if (%lg% = Nederlands)
        {
            if(InStr(src , "NL"))
                match := true
        }
        if (%lg% = Français)
        {
            if(InStr(src , "FR"))
                match := true
        }
       if (%lg% = English)
       {
            if(InStr(src , "EN"))
                match := true
        }
        else
        {
            match := false
        }
        if(match = true)
               FileWrite("|DOTSTAT| [SUCCESS] the logo match the language set")
        else
               FileWrite("|DOTSTAT| [WARNING] : the logo may not match the language set")
    
    }


}