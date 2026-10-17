#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Each time the document is saved, a PDF of it is made beside it (report.docx -> report.pdf), laid out as it prints.
; Runs until the editor closes, or until you exit it from its tray icon. The PDF goes through Microsoft Print to PDF;
; your default printer stays as it is.
Rte.On("Saved", MakePdf)
Rte.On("AsyncError", (*) => TrayTip(Rte.LastAsyncError, "PDF on save failed", "Iconx"))
Rte.On("Closed", (*) => ExitApp())
Rte.SetStatus("PDF on save is on: each save also makes a PDF beside the document.")

MakePdf(*) {
    try {
        SplitPath(Rte.FileName, , &Folder, , &Name)
        Pdf := Folder "\" Name ".pdf"
        Rte.Post("ExportPdf", Pdf)    ; (queued: the editor carries on while the PDF is made)
        Rte.Post("SetStatus", "Saved, and made " Name ".pdf beside it.")
    }
}
