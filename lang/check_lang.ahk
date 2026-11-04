#Requires AutoHotkey v2.0 64-bit
#Include "..\Csharp-in-AHK.ahk"
; Checks the languages' files (lang\<code>.json) against AhkWord's strings, with check_lang.cs: what each lacks (shown in
; English), what nothing uses any more, placeholders that don't fit; and writes lang\en.json, every key with its English.
; Run it after adding or changing strings in the code (T and TN calls) or in the ribbon, and after merging AhkWord.
;   check_lang.ahk [console] [code] [folder]
;     console: report on the console (exit code 1 on errors); code: one language (es); folder: another program folder
;     than this one's (with lang\ and ribbon\RteRibbon.strings.json: for tests)

Console := false, Code := "", Root := A_ScriptDir "\.."
for Arg in A_Args {
    if Arg = "console"
        Console := true
    else if InStr(Arg, "\")
        Root := Arg
    else
        Code := Arg
}
try {
    Text := RunCSharpFile(A_ScriptDir "\check_lang.cs", Root, Code)
    Failed := CSExitCode != 0
    if Text = ""    ; (an exception: its text is in the output)
        Text := CSOutput
} catch as e {
    Text := "check_lang.cs failed: " e.Message, Failed := true
}
if Console
    FileAppend(Text "`n", "*", "UTF-8")
else
    MsgBox(Text, "Check languages", Failed ? "Iconx" : "Iconi")
ExitApp(Failed ? 1 : 0)
