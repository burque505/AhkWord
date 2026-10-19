#Requires AutoHotkey v2.0 64-bit
#Include "..\Csharp-in-AHK.ahk"
; Builds RteRibbon.ribbon, the compiled ribbon that rte_writer35.cs loads, from RteRibbon.xml:
;   1. make_icons.cs draws the icons into build\icons (32-bit BMPs, 16 to 64 pixels for the DPI steps) from Microsoft's
;      Fluent UI System Icons (MIT license), whose font is downloaded once into build\fluent, at a fixed commit;
;   2. uicc.exe (Windows SDK) compiles the markup into a binary ribbon (.bml) and a resource script (.rc);
;   3. rc.exe (Windows SDK) compiles that into a .res file;
;   4. link.exe (Visual Studio C++ tools) wraps it in a resource-only DLL: RteRibbon.ribbon.
; Only needed after changing the markup or the icons; running RTE Writer needs only RteRibbon.ribbon.
; Pass "console" to report on the console instead of in a message box.

Console := A_Args.Length && A_Args[1] = "console"
Build := A_ScriptDir "\build"
; Fluent UI System Icons, https://github.com/microsoft/fluentui-system-icons (MIT license), at this commit
FluentCommit := "a347af95478f989038cccbf085978865264c048a"
try {
    FluentDir := Build "\fluent"
    DirCreate(FluentDir)
    for Item in ["fonts/FluentSystemIcons-Regular.ttf", "fonts/FluentSystemIcons-Regular.json", "LICENSE"] {
        Target := FluentDir "\" StrSplit(Item, "/")[-1]
        if !FileExist(Target) {
            Report("Downloading " Item)
            Download("https://raw.githubusercontent.com/microsoft/fluentui-system-icons/" FluentCommit "/" Item, Target)
        }
    }
    if DirExist(Build "\icons")
        DirDelete(Build "\icons", true)    ; no icons left over from an earlier build
    DirCreate(Build "\icons")
    Report(RunCSharpFile(A_ScriptDir "\make_icons.cs", Build "\icons", FluentDir))
    if CSExitCode
        throw Error("make_icons.cs failed:`n" CSOutput)
    SdkBin := SdkBinDir()
    Step("uicc", SdkBin "\uicc.exe", '"' A_ScriptDir '\RteRibbon.xml" RteRibbon.bml /header:RteRibbon.h /res:RteRibbon.rc /name:APPLICATION')
    Step("rc", SdkBin "\rc.exe", "/nologo RteRibbon.rc")
    Step("link", LinkExe(), '/NOLOGO /NOENTRY /DLL /MACHINE:X64 /OUT:"' A_ScriptDir '\RteRibbon.ribbon" RteRibbon.res')
    Report("Built " A_ScriptDir "\RteRibbon.ribbon (" FileGetSize(A_ScriptDir "\RteRibbon.ribbon") " bytes)", true)
} catch as e {
    Report("Build failed: " e.Message, true, true)
    ExitApp(1)
}
ExitApp(0)

; Runs a build tool in the build folder; its output goes to build\<name>.log and, if it fails, into the error.
Step(Name, Exe, Args) {
    Log := Build "\" Name ".log"
    Code := RunWait(A_ComSpec ' /c ""' Exe '" ' Args ' > "' Log '" 2>&1"', Build, "Hide")
    Output := Trim(FileRead(Log), " `r`n")
    if Code
        throw Error(Name " failed (exit code " Code "):`n" Output)
    Report(Name ": ok" (Output != "" ? "`n" Output : ""))
}

; The newest Windows 10/11 SDK bin folder that has uicc.exe and rc.exe (x64 preferred).
SdkBinDir() {
    Root := EnvGet("ProgramFiles(x86)") "\Windows Kits\10\bin"
    Best := "", BestVersion := ""
    Loop Files Root "\10.*", "D" {
        for Arch in ["x64", "x86"]
            if FileExist(A_LoopFileFullPath "\" Arch "\uicc.exe") && FileExist(A_LoopFileFullPath "\" Arch "\rc.exe") {
                if BestVersion = "" || VerCompare(A_LoopFileName, BestVersion) > 0
                    Best := A_LoopFileFullPath "\" Arch, BestVersion := A_LoopFileName
                break
            }
    }
    if Best = ""
        throw Error("uicc.exe and rc.exe not found under " Root ". Install the Windows SDK.")
    return Best
}

; link.exe of the newest Visual Studio with the C++ tools, found with vswhere.
LinkExe() {
    VsWhere := EnvGet("ProgramFiles(x86)") "\Microsoft Visual Studio\Installer\vswhere.exe"
    if !FileExist(VsWhere)
        throw Error("vswhere.exe not found: install Visual Studio (or Build Tools) with the C++ desktop tools.")
    Out := Build "\vswhere.txt"
    RunWait(A_ComSpec ' /c ""' VsWhere '" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath > "' Out '""', , "Hide")
    Install := Trim(FileRead(Out), " `r`n")
    Best := "", BestVersion := ""
    if Install != ""
        Loop Files Install "\VC\Tools\MSVC\*", "D"
            if FileExist(A_LoopFileFullPath "\bin\Hostx64\x64\link.exe") && (BestVersion = "" || VerCompare(A_LoopFileName, BestVersion) > 0)
                Best := A_LoopFileFullPath "\bin\Hostx64\x64\link.exe", BestVersion := A_LoopFileName
    if Best = ""
        throw Error("link.exe not found: install the C++ desktop tools of Visual Studio (or Build Tools).")
    return Best
}

Report(Text, Final := false, Failed := false) {
    global Console
    static All := ""
    All .= Text "`n"
    if Console
        FileAppend(Text "`n", "*", "UTF-8")
    else if Final
        MsgBox(All, "Build ribbon", Failed ? "Iconx" : "Iconi")
}
