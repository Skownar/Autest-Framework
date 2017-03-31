class Excel extends OfficeManager{
    ExeName := ""
    XL := ""
    __New(){
        base.__New()
        this.ExeName := this.p.getProperty("excelexe")
    }
    Run()
    {
            FileWrite("|EXCEL| run Application")
            this.XL := ComObjCreate("Excel.Application")
            this.XL.Visible := True
            Winwait ahk_class XLMAIN
            WinMaximize ahk_class XLMAIN
            this.flag := false
            this.XL.Workbooks.Add
    }
    
    WriteInCell(cell = "", value = ""){
        FileWrite("|EXCEL| write in cell {" . cell . "} value {" . value . "}")
        this.XL.ActiveSheet.Range(cell).Value := value
    }
    ;TODO check if exist file saved , if ! -> return + warning
    SaveAsPdf(path,filename){
        FileWrite("|EXCEL| try save as pdf {" . path . "\" . filename . ".pdf }")
        wsa := this.XL.ActiveSheet
        ;xlTypePDF = 0
        loc := path . "\" . filename
        wsa.ExportAsFixedFormat(0,loc)
        FileWrite("|EXCEL| saved as pdf {" . path . "\" . filename . ".pdf }")

        return loc . ".pdf"
    }
    ;TODO check if exist file saved , if ! -> return + warning  
    Save(location,filename){
        FileWrite("|EXCEL| try save as xlsx {" . path . "\" . filename . ".xlsx }")
        savepoint := location . "\" . filename
        this.XL.ActiveSheet.SaveAs(savepoint)
        FileWrite("|EXCEL| saved as xlsx {" . path . "\" . filename . ".xlsx }")

        return savepoint . ".xlsx"
    }

    Close(save){
        if(!save)
            this.XL.Application.DisplayAlerts := false
        this.XL.Application.Quit()
        WinWaitClose, ahk_class XLMAIN
        FileWrite("|EXCEL| close Application")
    }

}