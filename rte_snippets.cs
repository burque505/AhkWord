// Snippets, for AhkWord (AhkWord.cs, through #load). Iteration 55.
//
// A snippet is a part (a template in a Parts folder: a signature block, a caption, a certificate of service) with an
// abbreviation: typed in AhkWord's document (sig, then Space or Enter, or at once), the abbreviation becomes the part,
// its formatting and all, and its {{fields}} are asked for as Insert from Template asks for them. The abbreviations are
// AutoHotkey hotstrings, set up by AhkWord.ahk (Hotstring, HotIf: this AhkWord's own window, the document in focus),
// so they work in AhkWord only, and in each AhkWord window by itself; a snippet's text goes in through the editor, not
// as keystrokes. Export as Script writes them as a script of plain-text hotstrings, to use in any program.
//   - Snippets: the list, kept in snippets.json beside the user's own templates (portable mode: in AhkWord's folder),
//     the abbreviations' checks, the hotstrings' definitions for AhkWord.ahk, and the export.
//   - SnippetsDialog (Insert > Snippets) and SnippetDialog (one snippet: its abbreviation, part and options).
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;

/// A snippet: its abbreviation, the part it inserts (by name), and how it fires: on an ending character (Space, Enter,
/// Tab, punctuation) or at once (Instant); whether the ending character typed stays after it (KeepEnding); whether the
/// abbreviation's case must match (CaseSensitive).
public sealed record Snippet(string Abbreviation, string Part, bool Instant = false, bool KeepEnding = false, bool CaseSensitive = false);

public static class Snippets
{
	/// Iteration 55: RTF without the paragraph mark its text ends with (the last \par, with only formatting and the closing
	/// braces after it): RichEdit writes one for the last paragraph, and inserted in the middle of a line it starts a new
	/// paragraph. An Enter typed at the end (an empty paragraph after) stays.
	public static string WithoutFinalParagraph(string rtf) {
		var last = Regex.Match(rtf, @"\\par(?![a-z])(?=(?:\s|\}|\\[a-z]+-?\d* ?)*$)", RegexOptions.RightToLeft);
		return last.Success ? rtf.Remove(last.Index, 4) : rtf;
	}

	/// The list's file, in a templates folder (the user's own).
	public static string FileIn(string ownTemplates) => Path.Combine(ownTemplates, "snippets.json");

	static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

	public static List<Snippet> Load(string file) {
		try {
			if (File.Exists(file) && JsonSerializer.Deserialize<List<Snippet>>(File.ReadAllText(file)) is { } list)
				return list.Where(s => s.Abbreviation != "" && s.Part != "").ToList();
		} catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { }
		return new List<Snippet>();
	}

	public static void Save(string file, IEnumerable<Snippet> snippets) {
		Directory.CreateDirectory(Path.GetDirectoryName(file)!);
		string temp = file + ".tmp";
		File.WriteAllText(temp, JsonSerializer.Serialize(snippets.OrderBy(s => s.Abbreviation, StringComparer.OrdinalIgnoreCase).ToList(), Json));
		File.Move(temp, file, true);
	}

	/// What is wrong with an abbreviation ("": nothing): 1 to 40 letters, digits and marks that aren't spaces, no colon
	/// or backtick (AutoHotkey's), and not another snippet's (other: that snippet's own, when it is being changed).
	public static string Problem(string abbreviation, IEnumerable<Snippet> all, string? other = null) {
		if (abbreviation.Length is 0 or > 40) return "An abbreviation is 1 to 40 characters.";
		if (abbreviation.Any(char.IsWhiteSpace)) return "An abbreviation has no spaces.";
		if (abbreviation.IndexOfAny(new[] { ':', '`' }) >= 0) return "An abbreviation can't have a colon or a backtick in it.";
		if (all.Any(s => string.Equals(s.Abbreviation, abbreviation, StringComparison.OrdinalIgnoreCase) && !string.Equals(s.Abbreviation, other, StringComparison.OrdinalIgnoreCase)))
			return $"Another snippet has the abbreviation \"{abbreviation}\".";
		return "";
	}

	// Common words, and short ones in legal writing, that would fire when typed for their own sake.
	static readonly HashSet<string> Common = new(StringComparer.OrdinalIgnoreCase) {
		"a", "an", "and", "are", "as", "at", "be", "but", "by", "can", "do", "for", "from", "had", "has", "have", "he", "her",
		"his", "i", "if", "in", "is", "it", "its", "may", "me", "my", "no", "not", "of", "on", "or", "our", "she", "so", "than",
		"that", "the", "their", "them", "then", "there", "they", "this", "to", "up", "us", "was", "we", "were", "what", "when",
		"which", "who", "will", "with", "you", "your", "all", "any", "one", "two", "out", "see", "id", "re", "cos", "sig", "co",
		"inc", "llc", "pa", "pc", "st", "dr", "mr", "ms", "mrs", "vs", "v", "et", "al", "ex", "art", "sec", "para", "case", "date",
	};

	/// A warning for an abbreviation that is also a word ("": none): it fires whenever the word is typed.
	public static string Warning(string abbreviation, bool instant) =>
		Common.Contains(abbreviation) ? $"\"{abbreviation}\" is also a word: the snippet will come up whenever it is typed as one. A letter or symbol in front of it (;{abbreviation}, x{abbreviation}) avoids that."
		: instant && abbreviation.Length < 3 ? $"\"{abbreviation}\" is short to fire at once: it would fire inside other words that have it."
		: "";

	/// AutoHotkey's options for a snippet's hotstring: B0 (the editor takes the abbreviation out, not keystrokes), * at
	/// once, C its case; an abbreviation that ends inside a word ("?") isn't one.
	public static string Options(Snippet s) => "B0" + (s.Instant ? "*" : "") + (s.CaseSensitive ? "C" : "");

	/// The hotstrings AhkWord.ahk sets up, one a line: their options, a tab, the abbreviation.
	public static string Definitions(IEnumerable<Snippet> snippets) => string.Join("\n", snippets.Select(s => Options(s) + "\t" + s.Abbreviation));

	/// The text before a position that the snippet's hotstring leaves when it fires: the abbreviation, and the ending
	/// character typed after it (Enter: a paragraph mark). Its length there, or -1 if the text before it isn't that.
	public static int Typed(string story, int caret, Snippet s, string ending) {
		string end = s.Instant ? "" : ending switch { "\n" or "\r" or "\r\n" => "\r", _ => ending };
		string wanted = s.Abbreviation + end;
		if (caret < wanted.Length || caret > story.Length) return -1;
		string before = story.Substring(caret - wanted.Length, wanted.Length);
		return string.Equals(before, wanted, s.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) ? wanted.Length : -1;
	}

	/// The snippets as an AutoHotkey v2 script of plain-text hotstrings, for any program (text: a part's plain text, its
	/// fields as their defaults or [their names]).
	public static string Script(IEnumerable<Snippet> snippets, Func<Snippet, string?> text) {
		var sb = new StringBuilder();
		sb.AppendLine("#Requires AutoHotkey v2.0");
		sb.AppendLine("; AhkWord's snippets as plain-text hotstrings, for any program (exported " + DateTime.Now.ToString("yyyy-MM-dd") + ").");
		sb.AppendLine("; Run this script with AutoHotkey v2 beside your own. AhkWord itself inserts them formatted, and asks for their fields.");
		sb.AppendLine("; To use them in some programs only, put #HotIf WinActive(\"ahk_exe notepad.exe\") (and the like) above them.");
		sb.AppendLine();
		foreach (var s in snippets.OrderBy(s => s.Abbreviation, StringComparer.OrdinalIgnoreCase)) {
			if (text(s) is not { } plain) continue;
			plain = Regex.Replace(plain.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n').TrimEnd('\n'), @"\{\{\s*([^{}|]+?)\s*(?:\|\s*([^{}]*?)\s*)?\}\}", m => m.Groups[2].Success && m.Groups[2].Value != "" ? m.Groups[2].Value : "[" + m.Groups[1].Value + "]");
			string options = "T" + (s.Instant ? "*" : "") + (s.KeepEnding ? "" : "O") + (s.CaseSensitive ? "C" : "");
			sb.AppendLine($"; {s.Part}");
			sb.AppendLine($":{options}:{s.Abbreviation}::");
			sb.AppendLine("(");
			foreach (string line in plain.Split('\n')) {
				// (in a continuation section a backtick escapes, and a line starting with ")" would end it)
				string escaped = line.Replace("`", "``");
				sb.AppendLine(escaped.StartsWith(')') ? "`" + escaped : escaped);
			}
			sb.AppendLine(")");
			sb.AppendLine();
		}
		return sb.ToString();
	}
}

/// Iteration 55: bold, italic and underline in a RichTextBox's selection, read and set as RichEdit has them
/// (EM_GETCHARFORMAT / EM_SETCHARFORMAT on the selection: only that effect, each character's font kept), and B, I, U
/// buttons that show them (pressed: all of the selection has it, or the caret's typing does).
static class RichFormat
{
	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	struct CHARFORMAT { public int cbSize; public uint dwMask, dwEffects; public int yHeight, yOffset, crTextColor; public byte bCharSet, bPitchAndFamily; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string szFaceName; }
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr Send(IntPtr hWnd, int msg, IntPtr wParam, ref CHARFORMAT format);
	const int EM_GETCHARFORMAT = 0x043A, EM_SETCHARFORMAT = 0x0444, SCF_SELECTION = 1;
	static uint Bit(FontStyle flag) => flag switch { FontStyle.Bold => 1u, FontStyle.Italic => 2u, _ => 4u };    // CFM_ / CFE_ BOLD, ITALIC, UNDERLINE

	/// Whether all of the selection has the effect (no selection: the caret's typing).
	public static bool Has(RichTextBox box, FontStyle flag) {
		if (!box.IsHandleCreated) return (box.SelectionFont ?? box.Font).Style.HasFlag(flag);
		var format = new CHARFORMAT { cbSize = Marshal.SizeOf<CHARFORMAT>(), szFaceName = "" };
		Send(box.Handle, EM_GETCHARFORMAT, (IntPtr)SCF_SELECTION, ref format);
		uint bit = Bit(flag);
		return (format.dwMask & bit) != 0 && (format.dwEffects & bit) != 0;
	}

	/// The effect on or off for the selection (no selection: for what is typed next).
	public static void Set(RichTextBox box, FontStyle flag, bool on) {
		uint bit = Bit(flag);
		var format = new CHARFORMAT { cbSize = Marshal.SizeOf<CHARFORMAT>(), dwMask = bit, dwEffects = on ? bit : 0, szFaceName = "" };
		Send(box.Handle, EM_SETCHARFORMAT, (IntPtr)SCF_SELECTION, ref format);
	}

	/// B, I and U as buttons that stay pressed while the effect is on; set(flag, on) when one is pressed or released.
	/// Show(box) presses them as box's selection has it.
	public sealed class Buttons
	{
		readonly Dictionary<FontStyle, CheckBox> buttons = new();
		bool showing;
		public Buttons(Control into, Font font, Func<int, int> scale, Action<FontStyle, bool> set) {
			var tip = new ToolTip();
			foreach (var (label, flag, name) in new[] { ("B", FontStyle.Bold, "Bold (Ctrl+B)"), ("I", FontStyle.Italic, "Italic (Ctrl+I)"), ("U", FontStyle.Underline, "Underline (Ctrl+U)") }) {
				var button = new CheckBox {
					Appearance = Appearance.Button, Text = label, Font = new Font(font, flag), TextAlign = ContentAlignment.MiddleCenter,
					Width = scale(30), Height = scale(26), TabStop = false, AccessibleName = name.Split(' ')[0]
				};
				button.CheckedChanged += (s, e) => { if (!showing) set(flag, button.Checked); };
				tip.SetToolTip(button, name);
				buttons[flag] = button;
				into.Controls.Add(button);
			}
		}
		public void Show(RichTextBox? box) {
			showing = true;
			try { foreach (var (flag, button) in buttons) button.Checked = box != null && Has(box, flag); }
			finally { showing = false; }
		}
	}
}

/// Insert > Snippets: the snippets, with their options; to add, change, remove, insert and export them. (Iteration 55)
/// A snippet's text is made and changed in its own dialog's box (SnippetDialog), typed or pasted there.
sealed class SnippetsDialog : Form
{
	readonly ListView list = new() { View = View.Details, FullRowSelect = true, HideSelection = false, MultiSelect = false, Dock = DockStyle.Fill };
	readonly Label note = new() { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 6, 3, 3) };
	List<Snippet> snippets;
	readonly Func<List<TemplateInfo>> parts;
	readonly string? selectionRtf;
	readonly Func<string, string?> loadPart;
	readonly Action<string, string> savePart;
	readonly Font textFont;

	/// What was asked for when the dialog closed: a snippet to insert now (Insert Now).
	public Snippet? ToInsert;
	/// The list as it is now (changes are saved as they are made: Changed).
	public event Action<List<Snippet>>? Changed;
	/// Export as Script: the path chosen.
	public event Action<string>? Export;

	/// selectionRtf: the document's selection (null: none), for New from Selection; loadPart, savePart: a part's RTF by
	/// its name (null: none), and saving it (a new part in your own templates' Parts folder); textFont: the text's font.
	public SnippetsDialog(List<Snippet> snippets, Func<List<TemplateInfo>> parts, string? selectionRtf, Func<string, string?> loadPart, Action<string, string> savePart, Font textFont) {
		this.snippets = snippets.ToList();
		(this.parts, this.selectionRtf, this.loadPart, this.savePart, this.textFont) = (parts, selectionRtf, loadPart, savePart, textFont);
		Text = "Snippets";
		FormBorderStyle = FormBorderStyle.Sizable;
		MinimizeBox = MaximizeBox = ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		// (room for the columns and the buttons at the start: the list's columns, its scroll bar, the buttons' column)
		Size = new Size(LogicalToDeviceUnits(900), LogicalToDeviceUnits(480));
		MinimumSize = new Size(LogicalToDeviceUnits(760), LogicalToDeviceUnits(440));
		Padding = new Padding(10);
		list.Columns.Add("Abbreviation", LogicalToDeviceUnits(110));
		list.Columns.Add("Inserts", LogicalToDeviceUnits(230));
		list.Columns.Add("Fires", LogicalToDeviceUnits(170));
		list.Columns.Add("Case", LogicalToDeviceUnits(70));
		var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3 };
		layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
		layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.Controls.Add(new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(660), 0), Margin = new Padding(3, 0, 3, 8),
			Text = "Type a snippet's abbreviation in the document (then Space or Enter, unless it fires at once) and it becomes its text, formatted, its fields asked for. Snippets work in AhkWord only; Export as Script makes them plain-text hotstrings for any program."
		}, 0, 0);
		layout.SetColumnSpan(layout.GetControlFromPosition(0, 0)!, 2);
		layout.Controls.Add(list, 0, 1);
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
		Button Add(string text, Action action, bool enabled = true) {
			var button = new Button { Text = text, AutoSize = true, MinimumSize = new Size(LogicalToDeviceUnits(170), 0), Enabled = enabled, Margin = new Padding(6, 0, 0, 6) };
			button.Click += (s, e) => action();
			buttons.Controls.Add(button);
			return button;
		}
		Add("&New...", () => NewSnippet(null));
		Add("New from &Selection...", () => NewSnippet(selectionRtf), selectionRtf != null);
		var change = Add("&Change...", ChangeSelected);
		var remove = Add("&Remove", RemoveSelected);
		var insert = Add("&Insert Now", () => { if (Selected is { } s) { ToInsert = s; DialogResult = DialogResult.OK; } });
		Add("E&xport as Script...", ExportCommand);
		var close = Add("Close", () => DialogResult = DialogResult.Cancel);
		layout.Controls.Add(buttons, 1, 1);
		layout.Controls.Add(note, 0, 2);
		layout.SetColumnSpan(note, 2);
		Controls.Add(layout);
		CancelButton = close;
		void Enable() { bool one = Selected != null; change.Enabled = remove.Enabled = insert.Enabled = one; }
		list.SelectedIndexChanged += (s, e) => { Enable(); ShowNote(); };
		list.DoubleClick += (s, e) => ChangeSelected();
		Fill();
		Enable();
		// (as big as what it holds, measured at the screen's scaling when it opens: the text above, the buttons' column
		// whole, the list's columns, two lines of the note below; a size fixed beforehand cut the last button off)
		Load += (s, e) => {
			var header = layout.GetControlFromPosition(0, 0)!;
			int columns = list.Columns.Cast<ColumnHeader>().Sum(c => c.Width) + SystemInformation.VerticalScrollBarWidth + LogicalToDeviceUnits(8);
			int width = Padding.Horizontal + Math.Max(columns + buttons.PreferredSize.Width + buttons.Margin.Horizontal + list.Margin.Horizontal, header.PreferredSize.Width + header.Margin.Horizontal);
			int noteLines = 2 * note.Font.Height + note.Margin.Vertical;
			int height = Padding.Vertical + header.PreferredSize.Height + header.Margin.Vertical + buttons.PreferredSize.Height + buttons.Margin.Vertical + noteLines + LogicalToDeviceUnits(6);
			ClientSize = new Size(Math.Max(ClientSize.Width, width), Math.Max(ClientSize.Height, height));
			MinimumSize = Size;
			note.MaximumSize = new Size(ClientSize.Width - Padding.Horizontal - LogicalToDeviceUnits(10), 0);
		};
	}

	Snippet? Selected => list.SelectedItems.Count == 1 ? list.SelectedItems[0].Tag as Snippet : null;
	TemplateInfo? PartOf(Snippet s) => parts().FirstOrDefault(p => string.Equals(p.Name, s.Part, StringComparison.OrdinalIgnoreCase));
	List<string> PartNames() => parts().Where(p => p.Part && p.Path.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase)).Select(p => p.Name).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

	static string Fires(Snippet s) => s.Instant ? "at once" : "on Space, Enter, Tab ..." + (s.KeepEnding ? " (kept)" : "");

	void Fill(string? select = null) {
		list.BeginUpdate();
		list.Items.Clear();
		foreach (var s in snippets.OrderBy(s => s.Abbreviation, StringComparer.OrdinalIgnoreCase)) {
			var item = new ListViewItem(new[] { s.Abbreviation, s.Part + (PartOf(s) == null ? "  (missing)" : ""), Fires(s), s.CaseSensitive ? "exact" : "any" }) { Tag = s };
			list.Items.Add(item);
			if (string.Equals(s.Abbreviation, select, StringComparison.OrdinalIgnoreCase)) item.Selected = true;
		}
		list.EndUpdate();
		ShowNote();
	}

	void ShowNote() {
		note.Text = snippets.Count == 0 ? "There are no snippets yet: choose New and type or paste its text, or select a signature block (or any text) in the document and choose New from Selection."
			: Selected is { } s ? (PartOf(s) == null ? $"The part \"{s.Part}\" isn't in the templates' Parts folders any more: Change it to give it its text again." : Snippets.Warning(s.Abbreviation, s.Instant))
			: "";
	}

	void Save(string? select) {
		Changed?.Invoke(snippets);
		Fill(select);
	}

	// New: the text typed or pasted (or the selection's); its part saved, and the snippet added
	void NewSnippet(string? rtf) {
		using var dlg = new SnippetDialog(rtf != null ? "New Snippet from the Selection" : "New Snippet", null, snippets, PartNames(), loadPart, textFont, rtf);
		if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Result is not { } made) return;
		if (!SavePart(made.Part, dlg)) return;
		snippets.Add(made);
		Save(made.Abbreviation);
	}

	void ChangeSelected() {
		if (Selected is not { } s) return;
		using var dlg = new SnippetDialog("Change Snippet", s, snippets, PartNames(), loadPart, textFont, null);
		if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Result is not { } changed) return;
		if (!SavePart(changed.Part, dlg)) return;
		snippets.Remove(s);
		snippets.Add(changed);
		Save(changed.Abbreviation);
	}

	// the snippet's text, if it is new or changed, saved as its part
	bool SavePart(string name, SnippetDialog dlg) {
		if (!dlg.TextEdited) return true;
		try { savePart(name, dlg.Rtf); return true; }
		catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
			MessageBox.Show(this, $"The text couldn't be saved as the part \"{name}\": {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return false;
		}
	}

	void RemoveSelected() {
		if (Selected is not { } s) return;
		if (MessageBox.Show(this, $"Remove the snippet \"{s.Abbreviation}\"? (Its part, {s.Part}, stays a part.)", Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
		snippets.Remove(s);
		Save(null);
	}

	void ExportCommand() {
		if (snippets.Count == 0) return;
		using var save = new SaveFileDialog { Title = "Export the Snippets as a Hotstring Script", Filter = "AutoHotkey script (*.ahk)|*.ahk", FileName = "AhkWord snippets.ahk", OverwritePrompt = true };
		if (save.ShowDialog(this) == DialogResult.OK) Export?.Invoke(save.FileName);
	}
}

/// One snippet: its abbreviation, its text (typed or pasted in the box, formatted with B, I, U and Font...; kept as a
/// part of that name, a new one or one there is), and how it fires.
sealed class SnippetDialog : Form
{
	readonly TextBox abbreviation = new() { Width = 160, AccessibleName = "Abbreviation" };
	readonly ComboBox part = new() { DropDownStyle = ComboBoxStyle.DropDown, Width = 300, AccessibleName = "Part" };
	readonly RichTextBox text = new() { Multiline = true, WordWrap = true, ScrollBars = RichTextBoxScrollBars.Vertical, DetectUrls = false, HideSelection = false, AccessibleName = "Snippet text" };
	readonly Label partNote = new() { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 0, 3, 4) };
	readonly RadioButton onEnding = new() { Text = "On an ending character: Space, Enter, Tab or punctuation (AutoHotkey's usual)", AutoSize = true, Checked = true };
	readonly RadioButton instant = new() { Text = "At once, as soon as the abbreviation is typed (AutoHotkey's *)", AutoSize = true };
	readonly CheckBox keep = new() { Text = "Keep the ending character after the snippet (Enter: a new paragraph after it)", AutoSize = true };
	readonly CheckBox exact = new() { Text = "Case-sensitive: the abbreviation fires only typed as it is here (AutoHotkey's C)", AutoSize = true };
	readonly Label problem = new() { AutoSize = true, ForeColor = Color.Firebrick, Margin = new Padding(3, 8, 3, 3) };
	readonly List<Snippet> all;
	readonly Snippet? old;
	readonly List<string> parts;
	readonly Func<string, string?> loadPart;
	string loadedFor = "";    // the part whose text the box shows
	readonly RichFormat.Buttons styleButtons;
	bool edited;

	public Snippet? Result { get; private set; }
	/// The snippet's text as RTF, and whether it is new or was changed (then saved as the part).
	public string Rtf => Snippets.WithoutFinalParagraph(text.Rtf ?? "");    // (iteration 55: the box's last paragraph mark isn't typed text)
	public bool TextEdited => edited || !parts.Contains(PartName, StringComparer.OrdinalIgnoreCase);
	string PartName => part.Text.Trim();

	public SnippetDialog(string title, Snippet? old, List<Snippet> all, List<string> parts, Func<string, string?> loadPart, Font textFont, string? rtf) {
		(this.all, this.old, this.parts, this.loadPart) = (all, old, parts, loadPart);
		Text = title;
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = MaximizeBox = ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Padding = new Padding(12);
		problem.MaximumSize = partNote.MaximumSize = new Size(LogicalToDeviceUnits(600), 0);
		var layout = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, Dock = DockStyle.Fill };
		void Row(string label, Control control) {
			layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 8, 3) });
			layout.Controls.Add(control);
		}
		void Whole(Control control) { layout.Controls.Add(control); layout.SetColumnSpan(control, 2); }
		Row("Abbreviation:", abbreviation);
		foreach (var name in parts) part.Items.Add(name);
		Row("Part (its name):", part);
		Whole(partNote);
		// the text: typed or pasted (Ctrl+V), formatted
		var tools = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 6, 0, 0) };
		tools.Controls.Add(new Label { Text = "Text:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 8, 0) });
		styleButtons = new RichFormat.Buttons(tools, Font, LogicalToDeviceUnits, (flag, on) => { RichFormat.Set(text, flag, on); text.Focus(); });
		var fontButton = new Button { Text = "Font...", AutoSize = true, TabStop = false, AccessibleName = "Text font" };
		fontButton.Click += (s, e) => ChooseFont();
		tools.Controls.Add(fontButton);
		tools.Controls.Add(new Label { Text = "Type it, or paste it (Ctrl+V). Fields such as {{Your Name}} are asked for when it goes in.", AutoSize = true, ForeColor = SystemColors.GrayText, Anchor = AnchorStyles.Left, Margin = new Padding(8, 7, 0, 0) });
		Whole(tools);
		text.Font = textFont;
		text.Size = new Size(LogicalToDeviceUnits(800), LogicalToDeviceUnits(170));
		Whole(text);
		Whole(new Label { Text = "Fires:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
		Whole(onEnding);
		Whole(instant);
		Whole(keep);
		Whole(exact);
		Whole(problem);
		var ok = new Button { Text = "OK", AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
		buttons.Controls.AddRange(new Control[] { cancel, ok });
		Whole(buttons);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		// (Enter in the text: a new paragraph, not OK; Ctrl+B, I, U)
		text.PreviewKeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) e.IsInputKey = true; };
		text.KeyDown += (s, e) => {
			FontStyle? flag = e.Control && !e.Alt && !e.Shift ? e.KeyCode switch { Keys.B => FontStyle.Bold, Keys.I => FontStyle.Italic, Keys.U => FontStyle.Underline, _ => null } : null;
			if (flag is { } f) { Toggle(f); e.SuppressKeyPress = e.Handled = true; }
		};
		if (old != null) {
			abbreviation.Text = old.Abbreviation;
			part.Text = old.Part;
			instant.Checked = old.Instant;
			keep.Checked = old.KeepEnding;
			exact.Checked = old.CaseSensitive;
		}
		if (rtf != null) { text.Rtf = rtf; loadedFor = "\0"; edited = true; }    // (the selection: a new part's text)
		else ShowPart();
		text.TextChanged += (s, e) => edited = true;
		text.SelectionChanged += (s, e) => styleButtons.Show(text);
		Shown += (s, e) => styleButtons.Show(text);
		part.TextChanged += (s, e) => ShowPart();
		void Enable() { keep.Enabled = onEnding.Checked; Check(false); }
		onEnding.CheckedChanged += (s, e) => Enable();
		abbreviation.TextChanged += (s, e) => { SuggestName(); Check(false); };
		Enable();
		ok.Click += (s, e) => { if (Check(true)) DialogResult = DialogResult.OK; };
	}

	// An existing part's text into the box (while it wasn't changed here); a new name: the box as it is, a new part.
	void ShowPart() {
		string name = PartName;
		bool existing = parts.Contains(name, StringComparer.OrdinalIgnoreCase);
		partNote.Text = name == "" ? "Name the part the text is kept as (in your templates' Parts folder): a new name makes a new part; a part there is shows its text below, and changing it changes the part."
			: existing ? $"The part \"{name}\" is there already: its text is below, and changing it changes the part (in templates too)."
			: $"A new part, \"{name}\", is made with the text below.";
		if (!existing || edited || string.Equals(loadedFor, name, StringComparison.OrdinalIgnoreCase)) return;
		bool was = edited;
		if (loadPart(name) is { } rtf) { try { text.Rtf = rtf; } catch (ArgumentException) { text.Text = ""; } }
		loadedFor = name;
		edited = was;
	}

	// (a new snippet: the part named after its abbreviation, until a name is typed)
	string suggested = "";
	void SuggestName() {
		if (old != null || (PartName != "" && PartName != suggested)) return;
		string name = new string(abbreviation.Text.Trim().Where(c => char.IsLetterOrDigit(c) || c is ' ' or '-' or '_').ToArray()).Trim();
		suggested = name == "" ? "" : char.ToUpper(name[0]) + name[1..];
		if (!parts.Contains(suggested, StringComparer.OrdinalIgnoreCase)) part.Text = suggested;
	}

	// (Ctrl+B, I, U: on unless all of the selection has it already, as Word does)
	void Toggle(FontStyle flag) {
		RichFormat.Set(text, flag, !RichFormat.Has(text, flag));
		styleButtons.Show(text);
	}

	void ChooseFont() {
		using var dlg = new FontDialog { ShowColor = true, ShowEffects = true, FontMustExist = true, AllowVerticalFonts = false };
		var current = text.SelectionFont ?? text.Font;
		dlg.Font = current;
		dlg.Color = text.SelectionColor;
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		text.SelectionFont = dlg.Font;
		text.SelectionColor = dlg.Color;
		text.Focus();
	}

	// The abbreviation's and the part's problems, shown; on OK, the snippet made if there are none (a word's warning
	// doesn't stop it).
	bool Check(bool ok) {
		string typed = abbreviation.Text.Trim();
		string wrong = typed == "" && !ok ? "" : Snippets.Problem(typed, all, old?.Abbreviation);
		if (wrong == "" && ok && PartName == "") wrong = "Give the part a name.";
		if (wrong == "" && PartName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) wrong = "A part's name can't have \\ / : * ? \" < > | in it.";
		if (wrong == "" && ok && text.TextLength == 0) wrong = "Type or paste the snippet's text.";
		problem.ForeColor = wrong != "" ? Color.Firebrick : SystemColors.GrayText;
		problem.Text = wrong != "" ? wrong : typed == "" ? "" : Snippets.Warning(typed, instant.Checked);
		if (!ok || wrong != "") return false;
		Result = new Snippet(typed, PartName, instant.Checked, onEnding.Checked && keep.Checked, exact.Checked);
		return true;
	}
}
