#Requires AutoHotkey v2.0

; RteClient2.ahk - control a running RTE Writer (rte_writer2.ahk) from any AutoHotkey v2 script, and react to its events.
; Same as RteClient.ahk, plus (iteration 2) Rte.On / Rte.Off for events and Rte.Post for calls that do not wait.
;
; Usage:
;   #Include "RteClient2.ahk"
;   Rte.InsertText("Hello")             ; call any RTE Writer / RichEditBox method by name
;   Rte.ToggleBold()
;   Rte.SetBlockStyle("Heading1")       ; enum values by name
;   Rte.SetFontColor("#C00000")         ; colors as #RRGGBB or names ("Red")
;   MsgBox Rte.Html                     ; read a property ...
;   Rte.SelectionStart := 0             ; ... or set one
;   MsgBox Rte.Help()                   ; everything that can be called (also Help > Automation API in the editor)
;
;   Rte.On("TextChanged", (e) => ToolTip(e.TextLength " characters"))    ; react to an event
;   Rte.Off("TextChanged")                                              ; stop (or Rte.Off() for all events)
;   Rte.Post("Save")                    ; queue a call and continue at once
;
; Arguments are sent as text and converted to the parameter types (numbers, 1/0 for booleans, enum names, colors).
; Results come back as text (booleans as 1/0, lists one item per line). Errors in the editor are thrown as AHK errors.
;
; Calls are synchronous: each runs on the editor's UI thread and returns when it has finished, waiting at most
; Rte.Timeout milliseconds. The editor stays responsive between calls.
;   - From another script: the request goes to the editor's hidden window "RTE Writer automation server <pid>" as
;     WM_COPYDATA, and the reply comes back the same way. Scripts started from the editor's Tools > Macros menu talk to
;     that editor (RTE_WRITER_HWND); otherwise the RTE Writer window used last (the frontmost; several can be open).
;   - From rte_writer2.ahk itself (hotkeys defined there): the call goes straight to the editor through RunCSharp.
;
; Events: the editor posts them to this script's window (A_ScriptHwnd) and never waits for the script. Callbacks run
; one at a time, in the order the events happened, in a separate thread; they may call Rte. The callback gets an
; object e with e.Name and, depending on the event:
;   TextChanged       e.TextLength      sent when typing pauses (about 0.2 s)
;   SelectionChanged  e.Start, e.Length sent when the caret stops moving (about 0.1 s)
;   ModifiedChanged   e.Modified        1 = there are unsaved changes, 0 = none
;   Saved, Opened (also New), Closed (the editor has closed), AsyncError (see Rte.LastAsyncError)
; Rte.On makes the script persistent; exit with ExitApp (for example on the Closed event).

class Rte {
    static Timeout := 10000

    ; True if an RTE Writer is running.
    static IsRunning => this._Server(false) != 0

    ; Rte.Call("Name", args*) is the same as Rte.Name(args*).
    static Call(Name, Args*) => this._Request("call", Name, Args*)

    ; Queues a call and returns at once, without its result. Queued calls run in order, before any later call.
    ; If one fails, the error is in Rte.LastAsyncError and the AsyncError event happens.
    static Post(Name, Args*) {
        this._Request("post", Name, Args*)
    }

    ; Calls Callback(e) whenever the event happens (see the list above).
    static On(Event, Callback) {
        Name := this._EventName(Event)
        if !this._Handlers.Has(Name) {
            this._Request("call", "Subscribe", Name, A_ScriptHwnd)
            this._Handlers[Name] := []
        }
        this._Handlers[Name].Push(Callback)
        Persistent()
    }

    ; Removes one callback, all callbacks of an event, or (no parameters) all callbacks.
    static Off(Event := "*", Callback?) {
        Names := []
        if Event = "*" {
            for Name in this._Handlers
                Names.Push(Name)
        } else
            Names.Push(this._EventName(Event))
        for Name in Names {
            if !this._Handlers.Has(Name)
                continue
            Callbacks := this._Handlers[Name]
            if IsSet(Callback) {
                Loop Callbacks.Length
                    if Callbacks[i := Callbacks.Length - A_Index + 1] = Callback
                        Callbacks.RemoveAt(i)
                if Callbacks.Length
                    continue
            }
            this._Handlers.Delete(Name)
            try this._Request("call", "Unsubscribe", Name, A_ScriptHwnd)    ; the editor may have closed already
        }
    }

    static __Call(Name, Args) => this._Request("call", Name, Args*)
    static __Get(Name, Params) => this._Request("get", Name, Params*)
    static __Set(Name, Params, Value) {
        ; The static initializers below (Timeout, _Reply, ...) also arrive here: keep those in this class.
        if Name = "Timeout" || SubStr(Name, 1, 1) = "_"
            return this.DefineProp(Name, {Value: Value})
        this._Request("set", Name, Value)
    }

    static _Reply := ""
    static _InProcess := '
    (
        var dispatch = (Func<string[], string>?)AppContext.GetData("rte.dispatch") ?? throw new InvalidOperationException("RTE Writer is not running.");
        AHK.Return(dispatch(args[0].Split((char)31)));
    )'
    ; Event numbers (wParam) as sent by the editor, starting at 1.
    static _EventNames := ["TextChanged", "SelectionChanged", "ModifiedChanged", "Saved", "Opened", "Closed", "AsyncError"]
    static _EventMessage := DllCall("RegisterWindowMessage", "str", "RTE_WRITER_EVENT", "uint")
    static _Handlers := Map()
    static _Queue := []

    static __New() {
        if this != Rte
            return
        OnMessage(0x4A, ObjBindMethod(this, "_OnCopyData"))                 ; WM_COPYDATA replies
        OnMessage(this._EventMessage, ObjBindMethod(this, "_OnEvent"))      ; events
        this._RunHandlersTimer := ObjBindMethod(this, "_RunHandlers")
    }

    static _Request(Verb, Name, Args*) {
        Server := this._Server()
        Text := Verb Chr(31) Name
        for Arg in Args
            Text .= Chr(31) String(Arg)
        DllCall("GetWindowThreadProcessId", "ptr", Server, "uint*", &Pid := 0)
        if Pid = DllCall("GetCurrentProcessId") {
            ; The editor runs in this process (rte_writer2.ahk), so call it directly. Dynamic call: only rte_writer2.ahk
            ; needs Csharp-in-AHK.
            Reply := %"RunCSharp"%(this._InProcess, Text)
        } else {
            Data := Buffer(3 * A_PtrSize, 0)    ; COPYDATASTRUCT
            NumPut("ptr", 0x52544531, "uint", (StrLen(Text) + 1) * 2, Data)
            NumPut("ptr", StrPtr(Text), Data, 2 * A_PtrSize)
            this._Reply := ""
            Hidden := A_DetectHiddenWindows
            A_DetectHiddenWindows := true
            try SendMessage(0x4A, A_ScriptHwnd, Data.Ptr, , "ahk_id " Server, , , , this.Timeout)
            catch TimeoutError
                throw TimeoutError("RTE Writer did not answer within " this.Timeout " ms (" Name ").", -2)
            finally
                A_DetectHiddenWindows := Hidden
            Reply := this._Reply
        }
        if Reply = ""
            throw Error("No reply from RTE Writer (" Name ").", -2)
        if SubStr(Reply, 1, 1) = "1"
            throw Error(SubStr(Reply, 2), -2, Name)
        return SubStr(Reply, 2)
    }

    static _OnCopyData(wParam, lParam, Msg, Hwnd) {
        if NumGet(lParam, "ptr") != 0x52544532
            return
        this._Reply := StrGet(NumGet(lParam, 2 * A_PtrSize, "ptr"))
        return 1
    }

    ; Queues the event; the callbacks run in a timer thread, so this message handler returns at once.
    static _OnEvent(wParam, lParam, Msg, Hwnd) {
        if wParam < 1 || wParam > this._EventNames.Length
            return
        e := {Name: this._EventNames[wParam]}
        switch e.Name {
            case "TextChanged": e.TextLength := lParam
            case "SelectionChanged": e.Start := lParam & 0xFFFFFFFF, e.Length := lParam >> 32
            case "ModifiedChanged": e.Modified := lParam
        }
        this._Queue.Push(e)
        SetTimer(this._RunHandlersTimer, -1)
        return 0
    }

    static _RunHandlers() {
        while this._Queue.Length {
            e := this._Queue.RemoveAt(1)
            for Callback in this._Handlers.Get(e.Name, []).Clone()
                Callback(e)
        }
    }

    static _EventName(Event) {
        for Name in this._EventNames
            if Name = Event
                return Name
        throw ValueError("Unknown event: " Event ". Events: TextChanged, SelectionChanged, ModifiedChanged, Saved, Opened, Closed, AsyncError.", -2)
    }

    static _Server(Required := true) {
        Hidden := A_DetectHiddenWindows, Mode := A_TitleMatchMode
        A_DetectHiddenWindows := true, A_TitleMatchMode := "RegEx"
        try {
            if (Hwnd := EnvGet("RTE_WRITER_HWND")) != "" && WinExist("ahk_id " Hwnd)
                return Hwnd + 0
            ; the RTE Writer used last: the frontmost of its windows (WinGetList lists them front to back)
            for Window in WinGetList(" - RTE Writer$")
                if Hwnd := WinExist("^RTE Writer automation server " WinGetPID(Window) "$")
                    return Hwnd
            if Hwnd := WinExist("^RTE Writer automation server \d+$")
                return Hwnd
        } finally {
            A_DetectHiddenWindows := Hidden, A_TitleMatchMode := Mode
        }
        if Required
            throw TargetError("RTE Writer is not running (start rte_writer2.ahk).", -3)
        return 0
    }
}
