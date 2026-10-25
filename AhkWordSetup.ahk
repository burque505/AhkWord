#Requires AutoHotkey v2.0 64-bit
; AhkWord's first start (iteration 54): the .NET it runs on. #Included by AhkWord.ahk after Csharp-in-AHK.ahk; needs no
; .NET itself.
;
; AhkWord runs on Microsoft's .NET 10 Desktop Runtime. If the computer has it (installed for all of Windows), or
; AhkWord's folder has its own copy (a folder named dotnet beside AhkWord.exe: the "complete" download has it, or it
; was downloaded here before), nothing is asked. Else the user chooses:
;   - Download it for AhkWord only (as recommended): Microsoft's own zip of the runtime (about 55 MB), its address and
;     SHA-512 checksum from Microsoft's list of .NET 10 releases, checked, and unpacked into AhkWord\dotnet. Nothing
;     is installed in Windows, no administrator is needed, and deleting AhkWord's folder takes it away with the rest.
;   - Install it for all of Windows: Microsoft's download page opens (AhkWord closes; start it again after).
;   - Not now: AhkWord closes, and asks again next time. Nothing is ever downloaded without being asked for.

; Makes sure there is a .NET for AhkWord; returns false if there isn't one (and AhkWord should close).
AhkWordDotnet() {
    global CSDotnetRoot
    if CSDotnetRoot != "" || EnvGet("CSHARP_IN_AHK_DOTNET") != ""
        return true    ; (chosen already: by a script, or the environment)
    if _AhkWordHasDesktop(_CSharpDetectDotnet())
        return true    ; (the computer's own: Csharp-in-AHK finds it)
    Own := A_ScriptDir "\dotnet"
    if _AhkWordHasDesktop(Own) {
        CSDotnetRoot := Own
        return true
    }
    Choice := _AhkWordAskDotnet()
    if Choice = "windows" {
        Run("https://dotnet.microsoft.com/download/dotnet/10.0")
        MsgBox("When .NET 10 is installed (the .NET Desktop Runtime, x64, on the page that opened), start AhkWord again.", "AhkWord", "Iconi")
        return false
    }
    if Choice != "here"
        return false
    try {
        _AhkWordFetchDotnet(Own, true)
    } catch as e {
        MsgBox("AhkWord could not download .NET:`n`n" e.Message "`n`nTry again later, or download the complete AhkWord (with .NET inside) from`nhttps://github.com/burque505/AhkWord/releases", "AhkWord", "Iconx")
        return false
    }
    CSDotnetRoot := Own
    return true
}

; Whether a folder is a .NET with its Windows Desktop part, 10 or later.
_AhkWordHasDesktop(Dir) {
    if Dir = "" || !FileExist(Dir "\host\fxr")
        return false
    Desktop := _CSharpLatestDir(Dir "\shared\Microsoft.WindowsDesktop.App")
    if Desktop = ""
        return false
    SplitPath(Desktop, &Version)
    return VerCompare(Version, "10.0.0") >= 0
}

; The question: "here" (download it for AhkWord), "windows" (install it for all of Windows), or "" (not now).
_AhkWordAskDotnet() {
    Choice := ""
    Ask := Gui("-MinimizeBox", "AhkWord needs .NET")
    Ask.SetFont("s10", "Segoe UI")
    Ask.AddText("w440", "AhkWord runs on Microsoft's .NET 10 Desktop Runtime, which this computer hasn't got yet.")
    Ask.AddText("w440 y+8", "AhkWord can download it from Microsoft for itself only (about 55 MB): it goes into AhkWord's own folder, nothing is installed in Windows, and deleting the folder removes it. Or you can install it for all of Windows from Microsoft's page.")
    Here := Ask.AddButton("w440 y+14 Default", "Download it for AhkWord only (recommended)")
    Windows := Ask.AddButton("w440", "Install it for all of Windows (opens Microsoft's page)")
    Later := Ask.AddButton("w440", "Not now")
    Here.OnEvent("Click", (*) => (Choice := "here", Ask.Destroy()))
    Windows.OnEvent("Click", (*) => (Choice := "windows", Ask.Destroy()))
    Later.OnEvent("Click", (*) => Ask.Destroy())
    Ask.OnEvent("Close", (*) => Ask.Destroy())
    Ask.Show()
    WinWaitClose(Ask.Hwnd)
    return Choice
}

; Downloads the latest .NET 10 Desktop Runtime (win-x64 zip) into Into, checked against Microsoft's SHA-512 for it.
; Show: a window with the download's progress. Throws if it can't (offline, the check fails ...).
_AhkWordFetchDotnet(Into, Show := false) {
    Temp := A_Temp "\AhkWord-dotnet-" DllCall("GetCurrentProcessId")
    try DirDelete(Temp, true)
    DirCreate(Temp)
    Window := 0
    try {
        ; Microsoft's list of .NET 10 releases, newest first: the first win-x64 zips in it of the .NET runtime (the host
        ; and Microsoft.NETCore.App) and of the Windows Desktop runtime (Microsoft.WindowsDesktop.App only), both needed
        List := Temp "\releases.json"
        Download("https://builds.dotnet.microsoft.com/dotnet/release-metadata/10.0/releases.json", List)
        Json := FileRead(List, "UTF-8")
        Files := []
        for Kind in ["dotnet-runtime", "windowsdesktop-runtime"] {
            if !RegExMatch(Json, 's)"rid"\s*:\s*"win-x64"\s*,\s*"url"\s*:\s*"(https://[^"]*/' Kind '-([\d.]+)-win-x64\.zip)"\s*,\s*"hash"\s*:\s*"([0-9a-fA-F]{128})"', &Found)
                throw Error("Microsoft's list of .NET 10 releases has no " Kind " zip for x64.")
            Files.Push({Url: Found[1], Version: Found[2], Hash: Found[3], Zip: Temp "\" Kind ".zip"})
        }
        Version := Files[2].Version
        ; downloaded by Windows' curl in the background, the size shown as it comes
        if Show {
            Window := Gui("-MinimizeBox -SysMenu", "AhkWord: downloading .NET " Version)
            Window.SetFont("s10", "Segoe UI")
            Label := Window.AddText("w380", "Downloading the .NET " Version " Desktop Runtime from Microsoft...")
            Bar := Window.AddProgress("w380 h16 -Smooth +0x8")    ; (PBS_MARQUEE)
            SendMessage(0x40A, 1, 30, Bar)    ; PBM_SETMARQUEE
            Window.Show()
        }
        Got := 0
        for File in Files {
            Pid := 0
            Run('"' A_WinDir '\System32\curl.exe" -sSfL -o "' File.Zip '" "' File.Url '"', , "Hide", &Pid)
            while ProcessExist(Pid) {
                if Show && FileExist(File.Zip)
                    Label.Text := Format("Downloading the .NET {} Desktop Runtime from Microsoft... {:.0f} MB", Version, (Got + FileGetSize(File.Zip)) / 1048576)
                Sleep(250)
            }
            if !FileExist(File.Zip) || FileGetSize(File.Zip) < 1000000
                throw Error("The download didn't finish (is the computer online?).")
            Got += FileGetSize(File.Zip)
        }
        if Show
            Label.Text := "Checking and unpacking .NET " Version "..."
        Unpacked := Temp "\dotnet"
        DirCreate(Unpacked)
        for File in Files {
            if StrLower(_AhkWordSha512(File.Zip)) != StrLower(File.Hash)
                throw Error("A file downloaded isn't Microsoft's (its SHA-512 checksum doesn't match): it was not used.")
            RunWait('"' A_WinDir '\System32\tar.exe" -xf "' File.Zip '" -C "' Unpacked '"', , "Hide")
        }
        if !_AhkWordHasDesktop(Unpacked) || _CSharpLatestDir(Unpacked "\shared\Microsoft.NETCore.App") = ""
            throw Error("The runtime could not be unpacked.")
        try DirDelete(Into, true)
        DirMove(Unpacked, Into)
    } finally {
        if Window
            Window.Destroy()
        try DirDelete(Temp, true)
    }
}

; A file's SHA-512, in hexadecimal (Windows' certutil).
_AhkWordSha512(File) {
    Out := A_Temp "\AhkWord-sha512-" DllCall("GetCurrentProcessId") ".txt"
    RunWait(A_ComSpec ' /c ""' A_WinDir '\System32\certutil.exe" -hashfile "' File '" SHA512 > "' Out '""', , "Hide")
    Text := FileExist(Out) ? FileRead(Out) : ""
    try FileDelete(Out)
    return RegExMatch(Text, "m)^\s*([0-9a-fA-F ]{128,})\s*$", &M) ? StrReplace(M[1], " ") : ""
}
