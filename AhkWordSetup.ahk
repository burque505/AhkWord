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
; (Iteration 59) AhkWord's own .NET kept up to date: Microsoft fixes .NET about once a month, often for security, and
; Windows Update doesn't know a .NET in AhkWord's folder. At most once a week AhkWord fetches Microsoft's list of .NET
; 10 releases in the background (curl, not waited for: the start never waits for the internet), into its .NET folder;
; at the next start, if the list has a newer .NET 10 than the folder's, it asks, before .NET is loaded: Update now
; (downloaded and checked as the first time), Later (asks again next time), or Skip this version. The new .NET goes
; in in one move: the old folder renamed out of the way first (which fails while another AhkWord window runs on it:
; then nothing changes), the new one put in its place, the old one deleted.

; Makes sure there is a .NET for AhkWord; returns false if there isn't one (and AhkWord should close).
AhkWordDotnet() {
    global CSDotnetRoot
    if CSDotnetRoot != "" || EnvGet("CSHARP_IN_AHK_DOTNET") != ""
        return true    ; (chosen already: by a script, or the environment)
    if _AhkWordHasDesktop(_CSharpDetectDotnet())
        return true    ; (the computer's own: Csharp-in-AHK finds it)
    Own := A_ScriptDir "\dotnet"
    try DirDelete(Own ".old", true)    ; (left by an update while another window still ran on it)
    if _AhkWordHasDesktop(Own) {
        _AhkWordDotnetUpdate(Own)    ; (iteration 59)
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
        _AhkWordSwapDotnet(Unpacked, Into)
    } finally {
        if Window
            Window.Destroy()
        try DirDelete(Temp, true)
    }
}

; (Iteration 59) The new .NET (unpacked in New) in Into's place, in one move: Into renamed out of the way first (Into.old;
; that fails while a program runs on it, and then nothing has changed), New moved in, the old one deleted (left, if a
; file of it is still in use, for the next start). Throws if Into couldn't be moved.
_AhkWordSwapDotnet(New, Into) {
    if !DirExist(Into) {
        DirMove(New, Into)
        return
    }
    Old := Into ".old"
    try DirDelete(Old, true)
    try DirMove(Into, Old)
    catch
        throw Error("AhkWord's .NET is in use (another AhkWord window is open?). Close every AhkWord window, then start AhkWord again to update it.")
    try DirMove(New, Into)
    catch as e {
        DirMove(Old, Into)    ; (back as it was)
        throw e
    }
    try DirDelete(Old, true)
}

; (Iteration 59) The .NET in a folder: its Windows Desktop runtime's version ("10.0.13"), or "".
_AhkWordDotnetVersion(Dir) {
    Desktop := _CSharpLatestDir(Dir "\shared\Microsoft.WindowsDesktop.App")
    if Desktop = ""
        return ""
    SplitPath(Desktop, &Version)
    return Version
}

; (Iteration 59) What Microsoft's list of .NET 10 releases says (the file List): {Latest, Security, Date}, or 0.
_AhkWordReleaseInfo(List) {
    try Json := FileRead(List, "UTF-8")
    catch
        return 0
    if !RegExMatch(Json, '"latest-runtime"\s*:\s*"(10\.\d+\.\d+)"', &Latest)
        return 0
    Security := RegExMatch(Json, '"latest-release-security"\s*:\s*true') > 0
    Date := RegExMatch(Json, '"latest-release-date"\s*:\s*"(\d{4}-\d\d-\d\d)"', &D) ? D[1] : ""
    return {Latest: Latest[1], Security: Security, Date: Date}
}

; (Iteration 59) AhkWord's own .NET (in Own): a newer one offered if Microsoft's list, fetched before, has one; the
; list fetched again in the background if it is a week old or more (used at the next start).
_AhkWordDotnetUpdate(Own) {
    List := Own "\ahkword-releases.json", Skip := Own "\ahkword-skip.txt"
    ; (a list fetched last time: in use now if it came whole)
    if FileExist(List ".new") {
        if _AhkWordReleaseInfo(List ".new") {
            try FileMove(List ".new", List, true)
        } else {
            try FileDelete(List ".new")
        }
    }
    Have := _AhkWordDotnetVersion(Own)
    Info := FileExist(List) ? _AhkWordReleaseInfo(List) : 0
    if Info && Have != "" && VerCompare(Info.Latest, Have) > 0 && !(FileExist(Skip) && Trim(FileRead(Skip)) = Info.Latest) {
        Choice := _AhkWordAskUpdate(Have, Info)
        if Choice = "skip" {
            try FileOpen(Skip, "w").Write(Info.Latest)
        } else if Choice = "update" {
            try {
                _AhkWordFetchDotnet(Own, true)
                try FileDelete(Skip)
                MsgBox("AhkWord now runs on .NET " _AhkWordDotnetVersion(Own) ".", "AhkWord", "Iconi T4")
            } catch as e
                MsgBox("AhkWord could not update .NET, and goes on with " Have ":`n`n" e.Message, "AhkWord", "Iconx")
        }
    }
    ; (Microsoft's list again, at most once a week, by curl in the background: not waited for; never with a file named
    ; no-update-check beside AhkWord)
    if FileExist(A_ScriptDir "\no-update-check") || FileExist(A_ScriptDir "\no-update-check.txt")
        return
    Age := FileExist(List) ? DateDiff(A_Now, FileGetTime(List, "M"), "Days") : 999
    if Age >= 7 && !FileExist(List ".new")
        try Run('"' A_WinDir '\System32\curl.exe" -sSfL --max-time 60 -o "' List '.new" "https://builds.dotnet.microsoft.com/dotnet/release-metadata/10.0/releases.json"', , "Hide")
}

; (Iteration 59) The question: "update", "skip" or "" (later).
_AhkWordAskUpdate(Have, Info) {
    Choice := ""
    When := Info.Date != "" ? " (" FormatTime(StrReplace(Info.Date, "-") "000000", "d MMMM yyyy") ")" : ""
    Ask := Gui("-MinimizeBox", "AhkWord: a .NET update")
    Ask.SetFont("s10", "Segoe UI")
    Ask.AddText("w460", "Microsoft has released .NET " Info.Latest When (Info.Security ? ", a security update" : "") ". AhkWord has its own copy of .NET, " Have ", which Windows Update doesn't update.")
    Ask.AddText("w460 y+8", "AhkWord can download the new one from Microsoft (about 70 MB, checked against Microsoft's checksums, as the first time) and put it in place of the old one.")
    Update := Ask.AddButton("w460 y+14 Default", "Update now (recommended)")
    Later := Ask.AddButton("w460", "Later (ask again next time)")
    SkipIt := Ask.AddButton("w460", "Skip " Info.Latest)
    Update.OnEvent("Click", (*) => (Choice := "update", Ask.Destroy()))
    Later.OnEvent("Click", (*) => Ask.Destroy())
    SkipIt.OnEvent("Click", (*) => (Choice := "skip", Ask.Destroy()))
    Ask.OnEvent("Close", (*) => Ask.Destroy())
    Ask.Show()
    WinWaitClose(Ask.Hwnd)
    return Choice
}

; A file's SHA-512, in hexadecimal (Windows' certutil).
_AhkWordSha512(File) {
    Out := A_Temp "\AhkWord-sha512-" DllCall("GetCurrentProcessId") ".txt"
    RunWait(A_ComSpec ' /c ""' A_WinDir '\System32\certutil.exe" -hashfile "' File '" SHA512 > "' Out '""', , "Hide")
    Text := FileExist(Out) ? FileRead(Out) : ""
    try FileDelete(Out)
    return RegExMatch(Text, "m)^\s*([0-9a-fA-F ]{128,})\s*$", &M) ? StrReplace(M[1], " ") : ""
}
