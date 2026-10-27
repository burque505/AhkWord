// Event hooks, for AhkWord (AhkWord.cs, through #load). Iteration 56.
//
// A hook is a macro (an AutoHotkey script, in AhkWord's macros folder or anywhere) that AhkWord runs by itself when
// something happens: AhkWord starts, a document is opened, made, saved, printed or closed, or a style is changed. No
// script of your own has to keep running (as one with Rte.On must). The macro is started as Run AHK Script starts one,
// talking to this AhkWord window through RteClient2.ahk, and is told what happened:
//   A_Args[1]                      the event's name (Saved)
//   A_Args[2]                      what it was about: the file (Opened, BeforeSave, Saved, Closing), the style
//                                  (StyleChanged), the printer or the PDF (Printed); may be empty
//   EnvGet("AHKWORD_EVENT"), EnvGet("AHKWORD_EVENT_DETAIL")   the same
// BeforeSave and Closing wait for their macro to finish (at most WaitSeconds), so that it can change the document
// first; the others don't. While an event's macro runs, that event doesn't start it again (a Saved hook that saves).
// A hook can be switched off (its macro kept, not run: a PDF on each save only some of the time).
//   - HookEvent, EventHooks: the events, what each means, the hooks kept in the settings (event name -> macro).
//   - EventHooksDialog (Tools > Event Hooks): a macro for each event; New makes one from a starting script, Edit opens
//     it, Run Now tries it.
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

/// The events a hook can run on.
public enum HookEvent { Started, Opened, NewDocument, BeforeSave, Saved, StyleChanged, Printed, Closing }

public static class EventHooks
{
	/// How long BeforeSave and Closing wait for their macro.
	public const int WaitSeconds = 15;

	public static bool Waits(HookEvent e) => e is HookEvent.BeforeSave or HookEvent.Closing;

	public static string Describe(HookEvent e) => e switch {
		HookEvent.Started => "AhkWord has started (its window is open, with its first document)",
		HookEvent.Opened => "A document was opened (A_Args[2]: its file)",
		HookEvent.NewDocument => "A new document was made (File > New)",
		HookEvent.BeforeSave => "A document is about to be saved: AhkWord waits for the macro (A_Args[2]: the file)",
		HookEvent.Saved => "A document was saved (A_Args[2]: its file)",
		HookEvent.StyleChanged => "A style was changed, made, deleted or copied, or the style set (A_Args[2]: the style)",
		HookEvent.Printed => "A document was printed or exported as PDF (A_Args[2]: the printer, or the PDF)",
		HookEvent.Closing => "The window is about to close: AhkWord waits for the macro (A_Args[2]: the file)",
		_ => ""
	};

	public static HookEvent Parse(string name) =>
		Enum.TryParse<HookEvent>(name.Trim(), true, out var e) && Enum.IsDefined(e) ? e
			: throw new ArgumentException($"Unknown event \"{name}\". Hook events: {string.Join(", ", Enum.GetNames<HookEvent>())}.");

	/// The macro a hook runs (null: none), from the settings' list (event name -> macro, ignoring case).
	public static string? MacroFor(Dictionary<string, string>? hooks, HookEvent e) =>
		hooks?.FirstOrDefault(h => string.Equals(h.Key, e.ToString(), StringComparison.OrdinalIgnoreCase)).Value is { Length: > 0 } macro ? macro : null;

	/// Whether a hook is on (the settings' list of those switched off: event names).
	public static bool IsOn(List<string>? off, HookEvent e) => off == null || !off.Contains(e.ToString(), StringComparer.OrdinalIgnoreCase);

	/// A macro's file: a name in the macros folder, or a full path.
	public static string PathOf(string macro, string macrosFolder) => Path.IsPathRooted(macro) ? macro : Path.Combine(macrosFolder, macro);

	/// A new hook macro's starting script.
	public static string Starter(HookEvent e) {
		var sb = new StringBuilder();
		sb.AppendLine("#Requires AutoHotkey v2.0");
		sb.AppendLine("#Include \"..\\RteClient2.ahk\"");
		sb.AppendLine($"; A hook for AhkWord's {e} event (Tools > Event Hooks): {Describe(e)}.");
		sb.AppendLine("; AhkWord runs it by itself each time; Rte is the AhkWord window it happened in (Tools > Automation API lists");
		sb.AppendLine("; what it can do).");
		sb.AppendLine("Event := A_Args.Length >= 1 ? A_Args[1] : \"\"     ; \"" + e + "\"");
		sb.AppendLine("Detail := A_Args.Length >= 2 ? A_Args[2] : \"\"    ; " + (e switch {
			HookEvent.StyleChanged => "the style's name",
			HookEvent.Printed => "the printer, or the PDF's path",
			HookEvent.Started or HookEvent.NewDocument => "(empty)",
			_ => "the document's file (empty: not saved yet)"
		}));
		sb.AppendLine();
		sb.AppendLine(e switch {
			HookEvent.Saved => "; for example, a copy of the saved file in a Backups folder beside it, with the date and time in its name:\n; if Detail != \"\" {\n;     SplitPath(Detail, &Name, &Dir, &Ext, &Stem)\n;     DirCreate(Dir \"\\Backups\")\n;     FileCopy(Detail, Dir \"\\Backups\\\" Stem \" \" FormatTime(, \"yyyy-MM-dd HHmm\") \".\" Ext, true)\n; }",
			HookEvent.BeforeSave => "; for example, the date the document was saved, in its {{Saved}} field:\n; Rte.FillFields(\"Saved=\" FormatTime(, \"LongDate\"))",
			HookEvent.Opened => "; for example, a document opened in Print Layout, its page fitted to the window:\n; Rte.PrintLayout := true\n; Rte.ZoomMode := \"PageWidth\"",
			HookEvent.NewDocument => "; for example, every new document with today's date at the top:\n; Rte.InsertText(FormatTime(, \"LongDate\") \"`r`r\")",
			_ => "; what to do:\n; TrayTip(Event \" \" Detail, \"AhkWord\")"
		});
		return sb.ToString();
	}
}

/// Tools > Event Hooks: a macro for each event (none: nothing runs), and whether it is on.
sealed class EventHooksDialog : Form
{
	readonly Dictionary<HookEvent, ComboBox> boxes = new();
	readonly Dictionary<HookEvent, CheckBox> ons = new();
	readonly string macrosFolder;
	readonly Action<HookEvent, string> tryHook;

	/// The hooks chosen (event name -> macro).
	public Dictionary<string, string> Result => boxes.Where(b => b.Value.SelectedIndex > 0).ToDictionary(b => b.Key.ToString(), b => (string)b.Value.SelectedItem!);
	/// The hooks switched off (their macros kept): event names.
	public List<string> Off => boxes.Where(b => b.Value.SelectedIndex > 0 && !ons[b.Key].Checked).Select(b => b.Key.ToString()).ToList();

	/// tryHook: runs a macro as its event would (Run Now).
	public EventHooksDialog(Dictionary<string, string>? hooks, List<string>? off, string macrosFolder, Action<HookEvent, string> tryHook) {
		this.macrosFolder = macrosFolder;
		this.tryHook = tryHook;
		Text = "Event Hooks";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = MaximizeBox = ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Padding = new Padding(12);
		var layout = new TableLayoutPanel { AutoSize = true, ColumnCount = 6, Dock = DockStyle.Fill };
		var intro = new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(760), 0), Margin = new Padding(3, 0, 3, 10),
			Text = "A macro for an event runs by itself when it happens: no script of yours has to keep running. It is told the event in A_Args[1] and the document or style in A_Args[2], and works on this window through Rte (RteClient2.ahk). Macros are AutoHotkey scripts in the macros folder. Untick On to keep a macro chosen without it running."
		};
		layout.Controls.Add(intro);
		layout.SetColumnSpan(intro, 6);
		var tip = new ToolTip();
		foreach (var e in Enum.GetValues<HookEvent>()) {
			var name = new Label { Text = e.ToString(), AutoSize = true, Anchor = AnchorStyles.Left, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 8, 8, 0) };
			var box = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = LogicalToDeviceUnits(260), AccessibleName = e + " macro", Margin = new Padding(3, 4, 3, 0) };
			boxes[e] = box;
			var on = new CheckBox { Text = "On", AutoSize = true, Anchor = AnchorStyles.Left, Checked = EventHooks.IsOn(off, e), AccessibleName = e + " on", Margin = new Padding(8, 6, 3, 0) };
			ons[e] = on;
			Fill(box, EventHooks.MacroFor(hooks, e));
			var make = Small("New...", () => NewMacro(e), e + " new macro");
			var edit = Small("Edit", () => { if (box.SelectedIndex > 0) EditMacro((string)box.SelectedItem!); }, e + " edit");
			var run = Small("Run Now", () => { if (box.SelectedIndex > 0) tryHook(e, (string)box.SelectedItem!); }, e + " run now");
			void Enable() => edit.Enabled = run.Enabled = on.Enabled = box.SelectedIndex > 0;
			box.SelectedIndexChanged += (s, a) => Enable();
			Enable();
			tip.SetToolTip(name, EventHooks.Describe(e));
			tip.SetToolTip(on, "Ticked: the macro runs on the event. Unticked: it stays chosen, and doesn't run (Run Now still tries it).");
			layout.Controls.AddRange(new Control[] { name, box, on, make, edit, run });
			var what = new Label { Text = EventHooks.Describe(e), AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 0, 3, 2) };
			layout.Controls.Add(new Label { AutoSize = true });
			layout.Controls.Add(what);
			layout.SetColumnSpan(what, 5);
		}
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var folder = new Button { Text = "Macros Folder", AutoSize = true };
		folder.Click += (s, a) => { Directory.CreateDirectory(macrosFolder); Process.Start(new ProcessStartInfo(macrosFolder) { UseShellExecute = true }); };
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 12, 0, 0) };
		buttons.Controls.AddRange(new Control[] { cancel, ok, folder });
		layout.Controls.Add(buttons);
		layout.SetColumnSpan(buttons, 6);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
	}

	Button Small(string text, Action action, string name) {
		var button = new Button { Text = text, AutoSize = true, AccessibleName = name, Margin = new Padding(3, 3, 0, 0) };
		button.Click += (s, e) => action();
		return button;
	}

	// The macros folder's scripts (and a macro chosen elsewhere), "(none)" first.
	void Fill(ComboBox box, string? chosen) {
		var macros = Directory.Exists(macrosFolder) ? Directory.GetFiles(macrosFolder, "*.ahk").Select(Path.GetFileName).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList() : new();
		if (chosen != null && !macros.Contains(chosen, StringComparer.OrdinalIgnoreCase)) macros.Insert(0, chosen);
		box.Items.Clear();
		box.Items.Add("(none)");
		foreach (var m in macros) box.Items.Add(m!);
		int at = chosen == null ? 0 : box.Items.Cast<string>().ToList().FindIndex(i => string.Equals(i, chosen, StringComparison.OrdinalIgnoreCase));
		box.SelectedIndex = Math.Max(0, at);
	}

	// New...: a macro for the event from a starting script ("On Saved.ahk", or the next free name), chosen, and opened
	void NewMacro(HookEvent e) {
		Directory.CreateDirectory(macrosFolder);
		string name = $"On {e}.ahk";
		for (int n = 2; File.Exists(Path.Combine(macrosFolder, name)); n++) name = $"On {e} {n}.ahk";
		File.WriteAllText(Path.Combine(macrosFolder, name), EventHooks.Starter(e), new UTF8Encoding(true));
		foreach (var (other, box) in boxes) {
			string? kept = box.SelectedIndex > 0 ? (string)box.SelectedItem! : null;
			Fill(box, other == e ? name : kept);
		}
		EditMacro(name);
	}

	// Edit: the macro in the editor Windows has for scripts (AutoHotkey's Edit, or Notepad)
	void EditMacro(string macro) {
		string path = EventHooks.PathOf(macro, macrosFolder);
		try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true, Verb = "edit" }); }
		catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException) {
			Process.Start(new ProcessStartInfo("notepad.exe") { ArgumentList = { path }, UseShellExecute = false });
		}
	}
}
