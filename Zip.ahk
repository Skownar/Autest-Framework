;TODO check if in/out dirs & files exist 

ZipFile(path,file,archiveLocation,archiveName){
    RunWait, % A_ScriptDir . "\..\..\Lib\7zip\7zip.exe a " . archiveLocation . "\" . archiveName . " " . path . "\" . file
}

UnzipFile(archiveLocation,archiveName,destpath){
    RunWait, % A_ScriptDir . "\..\..\Lib\7zip\7zip.exe e " . archiveLocation . "\" . archiveName . " -o" . destpath 
}
