// The command palette, for AhkWord (AhkWord.cs, through #load). Iteration 57.
//
// Ctrl+Shift+P (Tools > Command Palette, and the quick access toolbar): type part of a command's name and press Enter.
// It lists the ribbon's commands, named and placed as the ribbon has them (RteRibbon.xml: the tab and group, the
// tooltip), with their shortcuts; the macros; the styles; the snippets; and the recent files. A command runs as its
// button does (the source views, macro recording); one the ribbon has greyed out is greyed out here. Used ones come
// first when nothing is typed.
//   - RibbonLabels: each command's label, tooltip and place, read from RteRibbon.xml beside the program.
//   - PaletteItem, Palette: what can be run, and the matching (every word typed, in the name first).
//   - PaletteForm: the box, the list, the keys.
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

/// The ribbon's own names for its commands (RteRibbon.xml): label, tooltip, and where it is ("Insert > Links").
public static class RibbonLabels
{
	public sealed record Label(string Text, string Tooltip, string Where);

	public static Dictionary<uint, Label> Load(string xmlPath) {
		var labels = new Dictionary<uint, Label>();
		if (!File.Exists(xmlPath)) return labels;
		XDocument doc;
		try { doc = XDocument.Load(xmlPath); } catch (Exception ex) when (ex is IOException or System.Xml.XmlException) { return labels; }
		XNamespace ns = "http://schemas.microsoft.com/windows/2009/Ribbon";
		var ids = new Dictionary<string, uint>();
		var texts = new Dictionary<string, (string Text, string Tooltip)>();
		foreach (var command in doc.Descendants(ns + "Command")) {
			string name = (string?)command.Attribute("Name") ?? "";
			if (name == "" || !uint.TryParse((string?)command.Attribute("Id"), out uint id)) continue;
			ids[name] = id;
			string text = Clean((string?)command.Attribute("LabelTitle") ?? (string?)command.Attribute("TooltipTitle") ?? "");
			string tip = (string?)command.Attribute("TooltipDescription") ?? "";
			texts[name] = (text, tip);
		}
		string TextOf(string? name) => name != null && texts.TryGetValue(name, out var t) ? t.Text : "";
		// where each command is: the File menu, or a tab's group (a contextual tab's set first)
		var places = new Dictionary<string, string>();
		foreach (var element in doc.Descendants().Where(e => e.Attribute("CommandName") != null)) {
			string name = (string)element.Attribute("CommandName")!;
			if (places.ContainsKey(name) || element.Name.LocalName is "Tab" or "Group" or "TabGroup" or "ApplicationMenu" or "QuickAccessToolbar") continue;
			if (element.Ancestors(ns + "QuickAccessToolbar").Any()) continue;    // (there too: its place on a tab is the one to tell)
			string where;
			if (element.Name.LocalName == "HelpButton") where = "the ribbon's (?) button";
			else if (element.Ancestors(ns + "ApplicationMenu").Any()) where = "File";
			else {
				var tab = element.Ancestors(ns + "Tab").FirstOrDefault();
				if (tab == null) continue;
				var set = element.Ancestors(ns + "TabGroup").FirstOrDefault();
				var group = element.Ancestors(ns + "Group").FirstOrDefault();
				where = string.Join(" > ", new[] { TextOf((string?)set?.Attribute("CommandName")), TextOf((string?)tab.Attribute("CommandName")), TextOf((string?)group?.Attribute("CommandName")) }.Where(t => t != ""));
			}
			places[name] = where;
		}
		// (a command only on the quick access toolbar: Undo, Redo)
		foreach (var element in doc.Descendants(ns + "QuickAccessToolbar").Descendants().Where(e => e.Attribute("CommandName") != null))
			places.TryAdd((string)element.Attribute("CommandName")!, "Quick access toolbar");
		foreach (var (name, id) in ids)
			labels[id] = new Label(texts[name].Text, texts[name].Tooltip, places.TryGetValue(name, out var w) ? w : "");
		return labels;
	}

	// (an & before a letter marks its access key, "C&lassic", and goes; one with spaces round it is shown, "Rows & Columns";
	// && is an &)
	static string Clean(string label) => Regex.Replace(Regex.Replace(label, @"&(&|(?=\S))", m => m.Groups[1].Value), @"\s+", " ").Trim();
}

/// Something the palette can run: its name, where it is (or what it is), a longer description, its shortcut, whether it
/// can run now, and a key for remembering it.
public sealed record PaletteItem(string Name, string Where, string Description, string Shortcut, Func<bool> Enabled, Action Run, string Key);

public static class Palette
{
	/// The items matching what was typed, the best first: every word in the item; in its name before its place or
	/// description; at a word's start before inside one. Nothing typed: the recent ones first, then the rest by name.
	public static List<PaletteItem> Match(IEnumerable<PaletteItem> items, string typed, IList<string> recent) {
		var words = typed.Split(' ', StringSplitOptions.RemoveEmptyEntries);
		if (words.Length == 0)
			return items.OrderBy(i => recent.IndexOf(i.Key) is int at && at >= 0 ? at : int.MaxValue).ThenBy(i => i.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
		var scored = new List<(PaletteItem Item, int Score)>();
		foreach (var item in items) {
			int score = 0;
			bool all = true;
			foreach (var word in words) {
				int s = Score(item.Name, word, 100);
				if (s == 0) s = Score(item.Where, word, 30);
				if (s == 0) s = Score(item.Description, word, 10);
				if (s == 0) { all = false; break; }
				score += s;
			}
			if (!all) continue;
			if (item.Name.StartsWith(typed.Trim(), StringComparison.CurrentCultureIgnoreCase)) score += 200;
			if (recent.IndexOf(item.Key) is int r && r >= 0) score += Math.Max(0, 20 - r);
			scored.Add((item, score));
		}
		return scored.OrderByDescending(s => s.Score).ThenBy(s => s.Item.Name, StringComparer.CurrentCultureIgnoreCase).Select(s => s.Item).ToList();
	}

	// a word in a text: at a word's start (weight), inside one (a third of it), not there (0)
	static int Score(string text, string word, int weight) {
		int at = text.IndexOf(word, StringComparison.CurrentCultureIgnoreCase);
		if (at < 0) return 0;
		return at == 0 || !char.IsLetterOrDigit(text[at - 1]) ? weight : weight / 3;
	}
}

/// The palette: a box to type in above the matching items; Up, Down, Page Up and Down to choose, Enter or a click to
/// run, Esc to close. Closes when it loses the focus.
sealed class PaletteForm : Form
{
	readonly TextBox box = new() { BorderStyle = BorderStyle.FixedSingle, AccessibleName = "Command" };
	readonly ListBox list = new() { DrawMode = DrawMode.OwnerDrawFixed, BorderStyle = BorderStyle.None, IntegralHeight = false, AccessibleName = "Commands" };
	readonly Label hint = new() { AutoSize = false, ForeColor = SystemColors.GrayText, TextAlign = ContentAlignment.MiddleLeft };
	readonly List<PaletteItem> items;
	readonly IList<string> recent;
	List<PaletteItem> shown = new();

	/// The item chosen (null: closed without one).
	public PaletteItem? Chosen { get; private set; }

	public PaletteForm(List<PaletteItem> items, IList<string> recent) {
		this.items = items;
		this.recent = recent;
		FormBorderStyle = FormBorderStyle.FixedToolWindow;
		ControlBox = false;
		Text = "";
		ShowInTaskbar = false;
		KeyPreview = true;
		StartPosition = FormStartPosition.Manual;
		Padding = new Padding(LogicalToDeviceUnits(6));
		BackColor = SystemColors.Window;
		box.Font = new Font(Font.FontFamily, Font.Size * 1.25f);
		list.ItemHeight = (int)Math.Ceiling(Font.Height * 2.3);
		hint.Height = Font.Height + LogicalToDeviceUnits(6);
		hint.Text = "A command, macro, style, snippet or recent file.   ↑ ↓ choose   Enter run   Esc close";
		box.Dock = DockStyle.Top;
		hint.Dock = DockStyle.Bottom;
		list.Dock = DockStyle.Fill;
		Controls.Add(list);
		Controls.Add(new Panel { Dock = DockStyle.Top, Height = LogicalToDeviceUnits(6) });
		Controls.Add(box);
		Controls.Add(hint);
		ClientSize = new Size(LogicalToDeviceUnits(640), LogicalToDeviceUnits(420));
		box.TextChanged += (s, e) => Show(box.Text);
		list.DrawItem += DrawItem;
		list.MouseClick += (s, e) => { int at = list.IndexFromPoint(e.Location); if (at >= 0) { list.SelectedIndex = at; RunSelected(); } };
		Deactivate += (s, e) => { if (Visible && !IsDisposed) Close(); };
		Show(box.Text);
	}

	/// Opens it over the top of owner's window, the typing box ready.
	public void ShowOver(Form owner) {
		var area = owner.RectangleToScreen(owner.ClientRectangle);
		Location = new Point(area.Left + (area.Width - Width) / 2, area.Top + Math.Min(LogicalToDeviceUnits(110), area.Height / 5));
		Show(owner);
		box.Focus();
	}

	void Show(string typed) {
		shown = Palette.Match(items, typed, recent);
		list.BeginUpdate();
		list.Items.Clear();
		foreach (var item in shown) list.Items.Add(item.Name);
		if (shown.Count > 0) list.SelectedIndex = Math.Max(0, shown.FindIndex(i => SafeEnabled(i)));
		list.EndUpdate();
	}

	static bool SafeEnabled(PaletteItem item) { try { return item.Enabled(); } catch (Exception) { return true; } }

	// the name (greyed when it can't run now), its shortcut at the right; below, where it is and what it does
	void DrawItem(object? sender, DrawItemEventArgs e) {
		if (e.Index < 0 || e.Index >= shown.Count) return;
		var item = shown[e.Index];
		bool selected = (e.State & DrawItemState.Selected) != 0, enabled = SafeEnabled(item);
		using var back = new SolidBrush(selected ? SystemColors.Highlight : SystemColors.Window);
		e.Graphics.FillRectangle(back, e.Bounds);
		var main = selected ? SystemColors.HighlightText : enabled ? SystemColors.WindowText : SystemColors.GrayText;
		var gray = selected ? SystemColors.HighlightText : SystemColors.GrayText;
		int pad = LogicalToDeviceUnits(8), half = e.Bounds.Height / 2;
		using var bold = new Font(Font, FontStyle.Bold);
		var top = new Rectangle(e.Bounds.Left + pad, e.Bounds.Top + LogicalToDeviceUnits(2), e.Bounds.Width - 2 * pad, half);
		var bottom = new Rectangle(top.Left, e.Bounds.Top + half - LogicalToDeviceUnits(1), top.Width, half);
		const TextFormatFlags flags = TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter;
		int keyWidth = item.Shortcut == "" ? 0 : TextRenderer.MeasureText(item.Shortcut, Font).Width + pad;
		TextRenderer.DrawText(e.Graphics, item.Name + (enabled ? "" : "  (not now)"), bold, new Rectangle(top.Left, top.Top, top.Width - keyWidth, top.Height), main, flags);
		if (keyWidth > 0) TextRenderer.DrawText(e.Graphics, item.Shortcut, Font, top, gray, flags | TextFormatFlags.Right);
		string below = string.Join("  ·  ", new[] { item.Where, item.Description }.Where(t => t != ""));
		TextRenderer.DrawText(e.Graphics, below, Font, bottom, gray, flags);
	}

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
		switch (keyData) {
			case Keys.Escape: Close(); return true;
			case Keys.Enter: RunSelected(); return true;
			case Keys.Down: Move(1); return true;
			case Keys.Up: Move(-1); return true;
			case Keys.PageDown: Move(Math.Max(1, list.ClientSize.Height / list.ItemHeight - 1)); return true;
			case Keys.PageUp: Move(-Math.Max(1, list.ClientSize.Height / list.ItemHeight - 1)); return true;
		}
		return base.ProcessCmdKey(ref msg, keyData);
	}

	void Move(int by) {
		if (shown.Count == 0) return;
		list.SelectedIndex = Math.Clamp(list.SelectedIndex + by, 0, shown.Count - 1);
	}

	void RunSelected() {
		if (list.SelectedIndex < 0 || list.SelectedIndex >= shown.Count) return;
		var item = shown[list.SelectedIndex];
		if (!SafeEnabled(item)) { hint.Text = $"{item.Name} can't run now (the ribbon has it greyed out)."; return; }
		Chosen = item;
		Close();
	}
}
