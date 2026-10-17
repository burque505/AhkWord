#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Makes a PDF of every document in a folder (Word, OpenDocument, RTF, Markdown, HTML and text files), laid out as each
; prints, into a "PDF" folder inside it. The document you have open is opened again afterwards. Save it first.
if Rte.Modified {
    MsgBox("Save (or close) the document you are working on first: Folder to PDF opens each document in turn.", "Folder to PDF", "Icon!")
    ExitApp
}
Folder := DirSelect(, 0, "Folder to PDF: choose the folder whose documents to make PDFs of")
if Folder = ""
    ExitApp
Files := []
loop files Folder "\*.*"
    if RegExMatch(A_LoopFileExt, "i)^(docx|doc|odt|rtf|md|markdown|html|htm|txt)$")
        Files.Push(A_LoopFileFullPath)
if Files.Length = 0 {
    MsgBox("There are no documents in " Folder ".", "Folder to PDF", "Iconi")
    ExitApp
}
if MsgBox("Make " Files.Length " PDF" (Files.Length = 1 ? "" : "s") " in " Folder "\PDF?", "Folder to PDF", "OKCancel Iconi") != "OK"
    ExitApp
Open := Rte.FileName
DirCreate(Folder "\PDF")
Made := 0, Failed := []
for Document in Files {
    SplitPath(Document, &Name, , , &Stem)
    Rte.SetStatus("Folder to PDF: " A_Index " of " Files.Length ": " Name "...")
    try {
        Rte.Open(Document)
        Rte.ExportPdf(Folder "\PDF\" Stem ".pdf")
        Made++
    } catch as e
        Failed.Push(Name ": " e.Message)
}
if Open != "" {
    try Rte.Open(Open)
} else
    Rte.New()
Message := Made " PDF" (Made = 1 ? "" : "s") " made in " Folder "\PDF."
for Failure in Failed
    Message .= (A_Index = 1 ? "`n`nNot made:" : "") "`n" Failure
Rte.SetStatus("Folder to PDF: " Made " made" (Failed.Length ? ", " Failed.Length " not" : "") ".")
MsgBox(Message, "Folder to PDF", Failed.Length ? "Icon!" : "Iconi")
