// Templates and the start page for RTE Writer (rte_writer31.cs, through #load).
//   - Templates: documents (.rtf, .docx, .odt, .md, .html, .txt) in the templates folder beside the program and in the
//     user's own Templates folder (beside the settings); a new document made from one is untitled. Fields in a template,
//     {{Your name}} or {{Recipient|Sir or Madam}} (a default after the bar), are asked for in a dialog when the document
//     is made (TemplateFieldsDialog). The values of fields about the user, named "Your ..." (Your name, Your address),
//     are remembered for the next time; the others (a paper's parties, its case number) start afresh each time. A field
//     whose name has the word "date" in it ({{Date}}, {{Hearing date}}, {{execution_date}}) gets a calendar to pick the
//     day from, and stays a box to type in ("on or about June 3"); nothing fills itself in.
//   - Parts: templates in a Parts folder (inside either templates folder) are pieces to insert into a document (Insert >
//     Insert from Template: a caption, a signature block ...), their fields filled in as a document's are.
//   - StartPage: shown in the editor's place at the start (and by File > Start Page): a blank document, the templates,
//     the recent documents and the macros; Esc goes to the blank document. TemplatePickerDialog: the templates in a list.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EllipticBit.RichEditorNET.TextObjectModel2;

public sealed record TemplateInfo(string Name, string Path, bool Own, bool Part = false);

public static class Templates
{
	public static readonly string[] Extensions = { ".rtf", ".docx", ".odt", ".md", ".markdown", ".html", ".htm", ".txt" };

	/// The templates in the folders, by name (the user's own replacing a shipped one of the same name): whole documents,
	/// and the parts in their Parts folders (pieces to insert into a document: a caption, a signature block ...).
	public static List<TemplateInfo> Find(string shipped, string own) {
		var byName = new Dictionary<(string, bool), TemplateInfo>();
		foreach (var (folder, isOwn) in new[] { (shipped, false), (own, true) })
			foreach (var (place, part) in new[] { (folder, false), (System.IO.Path.Combine(folder, "Parts"), true) }) {
				if (!Directory.Exists(place)) continue;
				foreach (var file in Directory.GetFiles(place).Where(f => Extensions.Contains(System.IO.Path.GetExtension(f).ToLowerInvariant()))) {
					string name = System.IO.Path.GetFileNameWithoutExtension(file);
					byName[(name.ToLowerInvariant(), part)] = new TemplateInfo(name, file, isOwn, part);
				}
			}
		return byName.Values.OrderBy(t => t.Part).ThenBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
	}

	// {{Name}} or {{Name|default}}
	static readonly Regex Field = new(@"\{\{\s*([^{}|]+?)\s*(?:\|\s*([^{}]*?)\s*)?\}\}");

	// "date" (or "dates") as a word of a field's name: Date, Hearing date, execution_date, Date of service; not Candidate
	static readonly Regex DateWord = new(@"(?<![\p{L}\p{N}])dates?(?![\p{L}\p{N}])", RegexOptions.IgnoreCase);

	/// Whether a field is a date (its name has the word "date" in it): the dialog gives it a calendar, and its value isn't
	/// remembered.
	public static bool IsDate(string name) => DateWord.IsMatch(name.Replace('_', ' '));

	// "Your" as a field name's first word: Your name, Your address, Your phone ...
	static readonly Regex YourWord = new(@"^\s*your(?![\p{L}\p{N}])", RegexOptions.IgnoreCase);

	/// Whether a field's value is remembered for the next time: fields about the user, named "Your ..." (Your name, Your
	/// address ...), which rarely change. Not the others (a paper's parties, its case number, a letter's recipient), so
	/// that last time's don't turn up in a new paper, nor dates.
	public static bool IsRemembered(string name) => YourWord.IsMatch(name.Replace('_', ' ')) && !IsDate(name);

	/// A day as the calendar writes it into a date field: the long date without the day of the week (October 17, 2026).
	public static string DateText(DateTime day) {
		string pattern = Regex.Replace(CultureInfo.CurrentCulture.DateTimeFormat.LongDatePattern, @"dddd,?\s*|,?\s*dddd", "").Trim(' ', ',');
		return day.ToString(pattern == "" ? "D" : pattern, CultureInfo.CurrentCulture);
	}

	/// The fields to ask for in a text, in order, each once (by name, ignoring case), with its default ("" if none).
	public static List<(string Name, string Default)> Fields(string text) {
		var fields = new List<(string, string)>();
		var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		foreach (Match m in Field.Matches(text)) {
			string name = m.Groups[1].Value;
			if (!seen.Add(name)) continue;
			fields.Add((name, m.Groups[2].Value));
		}
		return fields;
	}

	/// Puts the values into the document's fields (each keeping the field's formatting); value: a field's value, or null
	/// for its default (a field without one, and without a value, stays as it is). Lines in a value become paragraphs.
	/// Only the fields from start to end. Returns how many were filled.
	public static int Fill(ITextDocument2 doc, Func<string, string?> value, int start = 0, int end = int.MaxValue) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		int count = 0;
		foreach (Match m in Field.Matches(story).Reverse()) {
			if (m.Index < start || m.Index + m.Length > end) continue;
			string? text = value(m.Groups[1].Value);
			if (string.IsNullOrEmpty(text)) text = m.Groups[2].Success && m.Groups[2].Value != "" ? m.Groups[2].Value : null;
			if (text == null) continue;
			doc.Range2(m.Index, m.Index + m.Length).Text = text.Replace("\r\n", "\r").Replace('\n', '\r');
			count++;
		}
		return count;
	}
}

// The fields of a template (or a document), asked for: one box each, filled with the field's default, or for a field
// about the user ("Your ...") with the value given last time; "Forget remembered values" clears those. A field whose name mentions an address gets a box of several lines; a date field a calendar button
// beside its box (also Alt+Down), which writes the day chosen into the box, to change as wanted.
sealed class TemplateFieldsDialog : Form
{
	readonly List<(string Name, TextBox Box)> boxes = new();
	readonly List<(TextBox Box, string Default)> fromMemory = new();    // the boxes filled with a remembered value

	/// title: the window's ("New from Letter", "Fill In Fields" ...).
	/// forget: clears the remembered values (null: no Forget button).
	public TemplateFieldsDialog(string title, List<(string Name, string Default)> fields, IReadOnlyDictionary<string, string> remembered, Action? forget = null) {
		Text = title;
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MaximizeBox = MinimizeBox = false;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoScaleMode = AutoScaleMode.Dpi;
		Font = SystemFonts.MessageBoxFont ?? Font;
		var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(12), Dock = DockStyle.Fill };
		table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
		table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 340));
		var intro = new Label {
			Text = "Fill in the fields. Your own details (the fields named \"Your ...\") are remembered for next time; the others start empty each time. Leave a box empty to keep the field's default, or the field itself to fill in later (Insert > Fill In Fields).",
			AutoSize = true, MaximumSize = new Size(500, 0), Margin = new Padding(0, 0, 0, 10)
		};
		table.Controls.Add(intro, 0, 0);
		table.SetColumnSpan(intro, 2);
		int row = 1;
		foreach (var (name, fallback) in fields) {
			bool lines = name.Contains("address", StringComparison.OrdinalIgnoreCase), date = Templates.IsDate(name);
			var label = new Label { Text = name + ":", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(0, 6, 8, 0) };
			var box = new TextBox { Width = 340, Multiline = lines, AcceptsReturn = lines, Height = lines ? 64 : 0, ScrollBars = lines ? ScrollBars.Vertical : ScrollBars.None, AccessibleName = name };
			string start = Templates.IsRemembered(name) && remembered.TryGetValue(name, out var value) && value != "" ? value : fallback;
			if (start != fallback) fromMemory.Add((box, fallback));
			box.Text = start.Replace("\r\n", "\n").Replace("\n", "\r\n");
			if (box.Text == "" && fallback == "") box.PlaceholderText = date ? "Pick a day, or type the date" : "{{" + name + "}}";
			table.Controls.Add(label, 0, row);
			if (date) {
				// (the calendar's button inside the box, at its right, as a drop-down list has it; the text stops short of it)
				var pick = new Button { Text = "▾", Dock = DockStyle.Right, Width = 22, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Default, TabStop = false, AccessibleName = "Pick the " + name, BackColor = SystemColors.Control };
				pick.FlatAppearance.BorderSize = 0;
				new ToolTip().SetToolTip(pick, "Pick a day (Alt+Down)");
				pick.Click += (s, e) => ShowCalendar(box, pick);
				box.Controls.Add(pick);
				box.HandleCreated += (s, e) => SendMessage(box.Handle, 0xD3, (IntPtr)2, (IntPtr)(pick.Width << 16));    // EM_SETMARGINS, EC_RIGHTMARGIN
				box.KeyDown += (s, e) => { if (e.Alt && e.KeyCode == Keys.Down) { ShowCalendar(box, pick); e.Handled = true; } };
			}
			table.Controls.Add(box, 1, row);			boxes.Add((name, box));
			row++;
		}
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		buttons.Controls.Add(cancel);
		buttons.Controls.Add(ok);
		if (forget != null) {
			// (on the left, apart from OK and Cancel)
			var forgetButton = new Button { Text = "Forget remembered values", AutoSize = true, Margin = new Padding(3, 3, 60, 3), Enabled = remembered.Keys.Any(Templates.IsRemembered) };
			new ToolTip().SetToolTip(forgetButton, "Clear your details remembered from earlier (Your name, Your address ...), here and for next time.");
			forgetButton.Click += (s, e) => {
				forget();
				foreach (var (box, fallback) in fromMemory) box.Text = fallback;
				forgetButton.Enabled = false;
				forgetButton.Text = "Forgotten";
			};
			buttons.Controls.Add(forgetButton);
		}
		table.Controls.Add(buttons, 0, row);
		table.SetColumnSpan(buttons, 2);
		Controls.Add(table);
		AcceptButton = ok;
		CancelButton = cancel;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
	}

	// The calendar under a date field's box: at the date the box has (or today); choosing a day (click, or Enter) writes
	// it into the box; Esc leaves the box as it was. One calendar for the dialog, shown under whichever box asks, and
	// disposed with the dialog (not when it closes: it closes from inside the calendar's own handling of the click, which
	// goes on afterwards).
	ToolStripDropDown? popup;
	MonthCalendar? calendar;
	TextBox? dateBox;

	void ShowCalendar(TextBox box, Control under) {
		if (popup == null) {
			calendar = new MonthCalendar { MaxSelectionCount = 1, ShowTodayCircle = true };
			popup = new ToolStripDropDown { Padding = Padding.Empty, AutoClose = true };
			popup.Items.Add(new ToolStripControlHost(calendar) { Margin = Padding.Empty, Padding = Padding.Empty, AutoSize = true });
			calendar.DateSelected += (s, e) => Choose(e.Start);
			calendar.KeyDown += (s, e) => {
				if (e.KeyCode == Keys.Enter) { Choose(calendar.SelectionStart); e.Handled = true; }
				else if (e.KeyCode == Keys.Escape) { popup.Close(); dateBox?.Focus(); e.Handled = true; }
			};
		}
		dateBox = box;
		calendar!.SetDate(DateTime.TryParse(box.Text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out var current) ? current : DateTime.Today);
		popup.Show(box, new Point(0, box.Height));
		calendar.Focus();
	}

	// A day chosen in the calendar: into the box that asked, the calendar closed.
	void Choose(DateTime day) {
		if (dateBox == null) return;
		dateBox.Text = Templates.DateText(day);
		popup?.Close();
		dateBox.SelectAll();
		dateBox.Focus();
	}

	protected override void Dispose(bool disposing) {
		if (disposing) popup?.Dispose();    // (the calendar with it)
		base.Dispose(disposing);
	}
	/// The values typed, by field name ("" for an empty box).
	public Dictionary<string, string> Values => boxes.ToDictionary(b => b.Name, b => b.Box.Text.Replace("\r\n", "\n"), StringComparer.OrdinalIgnoreCase);

	/// Whether a field's box has a calendar (for tests and scripts' checks).
	public bool HasCalendar(string name) => boxes.Any(b => string.Equals(b.Name, name, StringComparison.OrdinalIgnoreCase) && b.Box.Controls.OfType<Button>().Any());

	[System.Runtime.InteropServices.DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);
}
// File > New from Template and Insert > Insert from Template: the templates in a list (for inserting, the parts first),
// with what each is; double-click or Enter chooses one.
sealed class TemplatePickerDialog : Form
{
	readonly ListBox list = new() { IntegralHeight = false, Dock = DockStyle.Fill };
	readonly List<TemplateInfo> templates;

	public TemplatePickerDialog(string title, IReadOnlyList<TemplateInfo> templates, string ownFolder) {
		this.templates = templates.ToList();
		Text = title;
		FormBorderStyle = FormBorderStyle.Sizable;
		MinimizeBox = MaximizeBox = false;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoScaleMode = AutoScaleMode.Dpi;
		Font = SystemFonts.MessageBoxFont ?? Font;
		ClientSize = new Size(440, 360);
		MinimumSize = new Size(320, 260);
		foreach (var t in this.templates)
			list.Items.Add(t.Name + "    (" + (t.Part ? "part" : "document") + (t.Own ? ", yours" : "") + ")");
		if (list.Items.Count > 0) list.SelectedIndex = 0;
		list.DoubleClick += (s, e) => { if (list.SelectedIndex >= 0) { DialogResult = DialogResult.OK; Close(); } };
		var note = new Label {
			Dock = DockStyle.Top, AutoSize = false, Height = 52, Padding = new Padding(0, 0, 0, 8),
			Text = this.templates.Count == 0 ? "There are no templates yet." : "Choose a template. Your own templates go in:\n" + ownFolder + " (parts, to insert into a document, in its Parts folder)."
		};
		var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(0, 8, 0, 0) };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, Enabled = this.templates.Count > 0 };
		buttons.Controls.Add(cancel);
		buttons.Controls.Add(ok);
		var body = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12) };
		body.Controls.Add(list);
		body.Controls.Add(note);
		body.Controls.Add(buttons);
		Controls.Add(body);
		AcceptButton = ok;
		CancelButton = cancel;
	}

	/// The template chosen (null: none).
	public TemplateInfo? Chosen => list.SelectedIndex >= 0 && list.SelectedIndex < templates.Count ? templates[list.SelectedIndex] : null;
}

// The start page: tiles for a blank document and each template, the recent documents, the macros, links to open a
// document and the templates folder, and "Show this page at the start". The keyboard: Tab between them, Enter chooses,
// Esc goes to the blank document.
sealed class StartPage : UserControl
{
	public event Action? BlankChosen, OpenChosen, TemplatesFolderChosen, Closing;
	public event Action<TemplateInfo>? TemplateChosen;
	public event Action<string>? RecentChosen, MacroChosen;
	public event Action<bool>? ShowAtStartChanged;

	readonly FlowLayoutPanel content;
	readonly Font titleFont, sectionFont, tileFont;

	public StartPage() {
		BackColor = SystemColors.Window;
		AutoScroll = true;
		titleFont = new Font("Segoe UI Light", 20f);
		sectionFont = new Font("Segoe UI Semibold", 11f);
		tileFont = new Font("Segoe UI", 9.5f);
		content = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Location = new Point(LogicalToDeviceUnits(36), LogicalToDeviceUnits(24)) };
		Controls.Add(content);
	}

	/// The first tile (for the focus).
	public Control? FirstTile { get; private set; }

	public void Fill(IReadOnlyList<TemplateInfo> templates, IReadOnlyList<string> recent, IReadOnlyList<(string Name, string Path, string? Keys)> macros, bool showAtStart) {
		SuspendLayout();
		content.SuspendLayout();
		foreach (Control c in content.Controls.Cast<Control>().ToList()) c.Dispose();
		content.Controls.Clear();
		int width = Math.Max(LogicalToDeviceUnits(480), ClientSize.Width - LogicalToDeviceUnits(80));
		content.Controls.Add(new Label { Text = "Start", Font = titleFont, AutoSize = true, Margin = new Padding(0, 0, 0, LogicalToDeviceUnits(8)) });

		Section("New");
		var tiles = new FlowLayoutPanel { AutoSize = true, MaximumSize = new Size(width, 0), WrapContents = true, Margin = new Padding(0) };
		FirstTile = Tile(tiles, "Blank document", "Start from nothing (Esc)", () => BlankChosen?.Invoke());
		foreach (var t in templates)
			Tile(tiles, t.Name, (t.Own ? "Your template" : "Template") + " · " + System.IO.Path.GetExtension(t.Path).TrimStart('.').ToUpperInvariant(), () => TemplateChosen?.Invoke(t));
		content.Controls.Add(tiles);
		var links = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, LogicalToDeviceUnits(2), 0, 0) };
		links.Controls.Add(Link("Open a document...", () => OpenChosen?.Invoke()));
		links.Controls.Add(Link("Your templates folder", () => TemplatesFolderChosen?.Invoke()));
		content.Controls.Add(links);

		if (recent.Count > 0) {
			Section("Recent");
			foreach (var path in recent.Take(8))
				content.Controls.Add(Link($"{System.IO.Path.GetFileName(path)}    {System.IO.Path.GetDirectoryName(path)}", () => RecentChosen?.Invoke(path)));
		}

		if (macros.Count > 0) {
			Section("Macros");
			content.Controls.Add(new Label { Text = "AutoHotkey scripts that work on the open document. The Macros button (Home and Tools tabs) runs them; Tools > Record Macro makes new ones.", AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 0, 0, 4) });
			foreach (var (name, path, keys) in macros.Take(10))
				content.Controls.Add(Link(name + (keys != null ? $"    ({keys})" : ""), () => MacroChosen?.Invoke(path)));
		}

		var show = new CheckBox { Text = "Show this page when RTE Writer starts", Checked = showAtStart, AutoSize = true, Margin = new Padding(3, LogicalToDeviceUnits(18), 0, LogicalToDeviceUnits(24)) };
		show.CheckedChanged += (s, e) => ShowAtStartChanged?.Invoke(show.Checked);
		content.Controls.Add(show);
		content.ResumeLayout();
		ResumeLayout();
	}

	void Section(string text) => content.Controls.Add(new Label { Text = text, Font = sectionFont, AutoSize = true, Margin = new Padding(3, LogicalToDeviceUnits(16), 0, LogicalToDeviceUnits(6)) });

	Button Tile(Control parent, string title, string note, Action chosen) {
		var tile = new Button {
			Text = title + "\n" + note, TextAlign = ContentAlignment.MiddleLeft, Font = tileFont, FlatStyle = FlatStyle.Flat,
			Size = new Size(LogicalToDeviceUnits(190), LogicalToDeviceUnits(64)), Margin = new Padding(0, 0, LogicalToDeviceUnits(10), LogicalToDeviceUnits(10)),
			Padding = new Padding(LogicalToDeviceUnits(8), 0, 0, 0), BackColor = SystemColors.Control, UseVisualStyleBackColor = false, AccessibleName = title
		};
		tile.FlatAppearance.BorderColor = SystemColors.ControlDark;
		tile.Click += (s, e) => chosen();
		parent.Controls.Add(tile);
		return tile;
	}

	LinkLabel Link(string text, Action chosen) {
		var link = new LinkLabel { Text = text, AutoSize = true, Font = tileFont, Margin = new Padding(3, 2, LogicalToDeviceUnits(18), 2), LinkBehavior = LinkBehavior.HoverUnderline };
		link.LinkClicked += (s, e) => chosen();
		return link;
	}

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
		if (keyData == Keys.Escape) { Closing?.Invoke(); return true; }
		return base.ProcessCmdKey(ref msg, keyData);
	}

	protected override void Dispose(bool disposing) {
		if (disposing) { titleFont.Dispose(); sectionFont.Dispose(); tileFont.Dispose(); }
		base.Dispose(disposing);
	}
}
