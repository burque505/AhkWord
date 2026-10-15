// Line numbers for RTE Writer (rte_writer19.cs, through #load; uses the Open XML SDK packages that rte_office6.cs lists).
// A document's line numbering is set as in Word (Layout > Line Numbers): off, continuous, or restarting on each page; the
// first line's number; which numbers are shown (every line's, or every 5th ...); the distance from the text. Lines in
// tables, page breaks and paragraphs whose numbers are suppressed have none and are not counted.
// The editor shows them left of the text (LineNumberGutter), counting the lines as it lays them out. They are kept in:
//   RTF            \linemod \linestarts \linex \lineppage or \linecont in the section (RichEdit drops these, so they are
//                  the document's LineNumbering), and \noline on paragraphs (RichEdit keeps that one: TOM's NoLineNumber)
//   Word           w:lnNumType in the section (also made by Word and LibreOffice from the RTF); suppressed paragraphs
//                  through Word and LibreOffice only (the built-in converter and pandoc go through HTML)
//   OpenDocument   text:linenumbering-configuration (it has no first number: lines start at 1)
//   HTML           <meta name="rte-writer-line-numbers" content="...">
//   Markdown       <!-- rte-writer-line-numbers: ... --> as the last line
using System.Drawing;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;
using W = DocumentFormat.OpenXml.Wordprocessing;

/// A document's line numbering. The other settings are kept while it is off, as in Word.
public sealed class LineNumbering
{
	public enum Modes { None, Continuous, EachPage }

	[JsonConverter(typeof(JsonStringEnumConverter))] public Modes Mode { get; set; }
	/// The first line's number (after each restart).
	public int Start { get; set; } = 1;
	/// Numbers shown: those that are multiples of this (1: every line's).
	public int CountBy { get; set; } = 1;
	/// From the text, in twips; 0: automatic (a quarter inch).
	public int Distance { get; set; }

	[JsonIgnore] public bool On => Mode != Modes.None;

	public LineNumbering Copy() => (LineNumbering)MemberwiseClone();

	public bool SameAs(LineNumbering other) => Mode == other.Mode && Start == other.Start && CountBy == other.CountBy && Distance == other.Distance;

	public LineNumbering Normalized() {
		Start = Math.Clamp(Start, 1, 32767);
		CountBy = Math.Clamp(CountBy, 1, 100);
		Distance = Math.Clamp(Distance, 0, 31680);
		return this;
	}

	public static Modes ParseMode(string mode) => mode.Trim().ToLowerInvariant() switch {
		"none" or "off" or "" or "0" => Modes.None,
		"continuous" or "on" or "1" => Modes.Continuous,
		"page" or "eachpage" or "each page" or "restart each page" => Modes.EachPage,
		_ => throw new ArgumentException($"\"{mode}\" is not a line numbering mode (none, continuous or page).")
	};

	public static string ModeName(Modes mode) => mode switch { Modes.Continuous => "continuous", Modes.EachPage => "page", _ => "none" };

	/// "none", or e.g. "continuous, start 1, count by 1, distance auto".
	public string Describe() => !On ? "none"
		: $"{ModeName(Mode)}, start {Start}, count by {CountBy}, distance {(Distance == 0 ? "auto" : (Distance / 20.0).ToString("0.##", CultureInfo.InvariantCulture) + "pt")}";

	static readonly JsonSerializerOptions Json = new();
	public string ToJson() => JsonSerializer.Serialize(this, Json);
	public static LineNumbering? FromJson(string json) {
		try { return JsonSerializer.Deserialize<LineNumbering>(json, Json)?.Normalized(); }
		catch (Exception ex) when (ex is JsonException or NotSupportedException) { return null; }
	}
}

public static class LineNumberFormats
{
	const string MetaName = "rte-writer-line-numbers";

	// ================================================================ RTF

	/// RTF with the line numbering in its section (after HeaderFooter.AddToRtf, which starts one with \sectd).
	public static string AddToRtf(string rtf, LineNumbering numbering) {
		if (!numbering.On) return rtf;
		string words = $@"\linemod{numbering.CountBy}" + (numbering.Start != 1 ? $@"\linestarts{numbering.Start}" : "")
			+ (numbering.Distance > 0 ? $@"\linex{numbering.Distance}" : "") + (numbering.Mode == LineNumbering.Modes.EachPage ? @"\lineppage" : @"\linecont");
		var section = Regex.Match(rtf, @"\\sectd(?![a-z])");
		if (section.Success) return rtf.Insert(section.Index + section.Length, words);
		var paragraph = Regex.Match(rtf, @"\\pard(?![a-z])");
		return paragraph.Success ? rtf.Insert(paragraph.Index, @"\sectd" + words + " ") : rtf;
	}

	/// The line numbering of RTF (null: none). Without \linecont or \linerestart it restarts on each page (Word and
	/// LibreOffice write nothing for that); restarting each section is continuous here (one section).
	public static LineNumbering? FromRtf(string rtf) {
		var countBy = Regex.Match(rtf, @"\\linemod(\d+)");
		if (!countBy.Success || int.Parse(countBy.Groups[1].Value) < 1) return null;
		int Number(string word, int otherwise) => Regex.Match(rtf, @"\\" + word + @"(\d+)") is { Success: true } m ? int.Parse(m.Groups[1].Value) : otherwise;
		return new LineNumbering {
			Mode = Regex.IsMatch(rtf, @"\\line(cont|restart)(?![a-z])") ? LineNumbering.Modes.Continuous : LineNumbering.Modes.EachPage,
			CountBy = int.Parse(countBy.Groups[1].Value), Start = Number("linestarts", 1), Distance = Number("linex", 0)
		}.Normalized();
	}

	// ================================================================ .docx

	// What may follow w:lnNumType in w:sectPr (the schema's order).
	static readonly Type[] AfterLineNumbers = {
		typeof(W.PageNumberType), typeof(W.Columns), typeof(W.FormProtection), typeof(W.VerticalTextAlignmentOnPage), typeof(W.NoEndnote),
		typeof(W.TitlePage), typeof(W.TextDirection), typeof(W.BiDi), typeof(W.GutterOnRight), typeof(W.DocGrid),
		typeof(W.PrinterSettingsReference), typeof(W.SectionPropertiesChange)
	};

	/// Sets a .docx's line numbering (in its last section, the whole document's for RTE Writer). Word's w:start is the
	/// first number less one.
	public static void AddToDocx(string path, LineNumbering numbering) {
		using var package = WordprocessingDocument.Open(path, true);
		var body = package.MainDocumentPart!.Document.Body!;
		var section = body.Elements<W.SectionProperties>().LastOrDefault();
		if (section == null) {
			if (!numbering.On) return;
			section = body.AppendChild(new W.SectionProperties());
		}
		section.RemoveAllChildren<W.LineNumberType>();
		if (numbering.On) {
			var element = new W.LineNumberType {
				CountBy = (short)numbering.CountBy,
				Restart = numbering.Mode == LineNumbering.Modes.EachPage ? W.LineNumberRestartValues.NewPage : W.LineNumberRestartValues.Continuous
			};
			if (numbering.Start != 1) element.Start = (short)(numbering.Start - 1);
			if (numbering.Distance > 0) element.Distance = numbering.Distance.ToString(CultureInfo.InvariantCulture);
			var next = section.ChildElements.FirstOrDefault(e => AfterLineNumbers.Contains(e.GetType()));
			if (next != null) section.InsertBefore(element, next); else section.AppendChild(element);
		}
		package.MainDocumentPart.Document.Save();
	}

	/// The line numbering of a .docx (null: none). Without w:restart it restarts on each page (the schema's default, and
	/// what Word writes for that).
	public static LineNumbering? FromDocx(string path) {
		try {
			using var package = WordprocessingDocument.Open(path, false);
			var body = package.MainDocumentPart?.Document.Body;
			var section = body?.Elements<W.SectionProperties>().LastOrDefault() ?? body?.Descendants<W.SectionProperties>().LastOrDefault();
			var element = section?.GetFirstChild<W.LineNumberType>();
			if (element?.CountBy?.Value is not short countBy || countBy < 1) return null;
			return new LineNumbering {
				Mode = element.Restart == null || element.Restart.Value == W.LineNumberRestartValues.NewPage ? LineNumbering.Modes.EachPage : LineNumbering.Modes.Continuous,
				CountBy = countBy, Start = (element.Start?.Value ?? 0) + 1,
				Distance = int.TryParse(element.Distance?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int d) ? d : 0
			}.Normalized();
		} catch (Exception ex) when (ex is IOException or InvalidDataException or OpenXmlPackageException) {
			return null;
		}
	}

	// ================================================================ .odt

	static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

	/// Sets an .odt's line numbering (its styles.xml's text:linenumbering-configuration).
	public static void AddToOdt(string path, LineNumbering numbering) {
		using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
		if (zip.GetEntry("styles.xml") is not { } entry) return;
		XDocument styles;
		using (var stream = entry.Open()) styles = XDocument.Load(stream);
		var root = styles.Root!;
		var officeStyles = root.Element(Office + "styles");
		if (officeStyles == null) {
			if (!numbering.On) return;
			officeStyles = new XElement(Office + "styles");
			var before = root.Element(Office + "automatic-styles") ?? root.Element(Office + "master-styles");
			if (before != null) before.AddBeforeSelf(officeStyles); else root.Add(officeStyles);
		}
		if (root.GetNamespaceOfPrefix("text") == null) root.Add(new XAttribute(XNamespace.Xmlns + "text", Text));
		if (root.GetNamespaceOfPrefix("style") == null) root.Add(new XAttribute(XNamespace.Xmlns + "style", Style));
		var config = officeStyles.Element(Text + "linenumbering-configuration");
		if (config == null) { config = new XElement(Text + "linenumbering-configuration"); officeStyles.Add(config); }
		config.SetAttributeValue(Text + "number-lines", numbering.On ? "true" : "false");
		config.SetAttributeValue(Text + "increment", numbering.CountBy);
		config.SetAttributeValue(Text + "restart-on-page", numbering.Mode == LineNumbering.Modes.EachPage ? "true" : "false");
		config.SetAttributeValue(Text + "offset", ((numbering.Distance > 0 ? numbering.Distance : 360) / 1440.0).ToString("0.####", CultureInfo.InvariantCulture) + "in");
		config.SetAttributeValue(Style + "num-format", "1");
		config.SetAttributeValue(Text + "number-position", "left");
		entry.Delete();
		var added = zip.CreateEntry("styles.xml", CompressionLevel.Optimal);
		using (var stream = added.Open()) styles.Save(stream, SaveOptions.DisableFormatting);
	}

	/// The line numbering of an .odt (null: none).
	public static LineNumbering? FromOdt(string path) {
		try {
			using var zip = ZipFile.OpenRead(path);
			if (zip.GetEntry("styles.xml") is not { } entry) return null;
			XDocument styles;
			using (var stream = entry.Open()) styles = XDocument.Load(stream);
			var config = styles.Descendants(Text + "linenumbering-configuration").FirstOrDefault();
			if (config == null || (string?)config.Attribute(Text + "number-lines") == "false") return null;
			return new LineNumbering {
				Mode = (string?)config.Attribute(Text + "restart-on-page") == "true" ? LineNumbering.Modes.EachPage : LineNumbering.Modes.Continuous,
				CountBy = int.TryParse((string?)config.Attribute(Text + "increment"), out int increment) ? increment : 1,
				Distance = RtfTables.CssLength(((string?)config.Attribute(Text + "offset") ?? "").Trim()) is int offset && Math.Abs(offset - 360) > 10 ? offset : 0
			}.Normalized();
		} catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException) {
			return null;
		}
	}

	// ================================================================ HTML and Markdown

	public static string HtmlMeta(LineNumbering numbering) =>
		numbering.On ? $"<meta name=\"{MetaName}\" content=\"{System.Net.WebUtility.HtmlEncode(numbering.ToJson())}\">\r\n" : "";

	public static LineNumbering? FromHtml(string html) {
		var m = Regex.Match(html, "<meta\\s+name=\"" + MetaName + "\"\\s+content=\"([^\"]*)\"", RegexOptions.IgnoreCase);
		return m.Success ? LineNumbering.FromJson(System.Net.WebUtility.HtmlDecode(m.Groups[1].Value)) : null;
	}

	public static string MarkdownLine(LineNumbering numbering) => numbering.On ? $"<!-- {MetaName}: {numbering.ToJson()} -->\n" : "";

	public static LineNumbering? FromMarkdown(string markdown) {
		var m = Regex.Match(markdown, @"^<!-- " + MetaName + @": (.*) -->\s*$", RegexOptions.Multiline);
		return m.Success ? LineNumbering.FromJson(m.Groups[1].Value) : null;
	}
}

// The line numbers left of the text: each line of the editor that counts gets its number (or every CountBy-th one), level
// with the line's text (on its baseline). Lines that don't count: in tables, page breaks (the break and its paragraph
// mark) and paragraphs with suppressed numbers. "Each page" restarts at the lines where printed pages start (PageStarts,
// worked out with the printer's layout by the form). The lines are the editor's own, so they follow its wrapping.
sealed class LineNumberGutter : Control
{
	readonly RichEditBox editor;
	LineNumbering numbering = new();
	List<int>? pageStarts;
	List<(int Start, int End)>? skipped;    // characters whose lines have no number (cached until the text changes)

	const int EM_GETLINECOUNT = 0x00BA, EM_LINEINDEX = 0x00BB, EM_GETFIRSTVISIBLELINE = 0x00CE, EM_EXLINEFROMCHAR = 0x0436;
	const int tomParagraph = 4;
	[DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	int Send(int msg, int wParam = 0, int lParam = 0) => (int)SendMessage(editor.Handle, msg, (IntPtr)wParam, (IntPtr)lParam);

	public LineNumberGutter(RichEditBox editor) {
		this.editor = editor;
		Dock = DockStyle.Left;
		Visible = false;
		SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.ResizeRedraw, true);
		BackColor = SystemColors.Window;
		TabStop = false;
		editor.TextChanged += (s, e) => DocumentChanged();
	}

	public LineNumbering Numbering {
		get => numbering;
		set { numbering = value.Copy(); Visible = numbering.On; FitWidth(); Invalidate(); }
	}

	/// Where printed pages start (character positions), for numbering that restarts on each page.
	public List<int>? PageStarts {
		get => pageStarts;
		set { pageStarts = value; Invalidate(); }
	}

	/// The text, or paragraphs' suppression, changed: which lines count is worked out again.
	public void DocumentChanged() {
		skipped = null;
		if (Visible) Invalidate();
	}

	/// After the editor painted (scrolled, typed, zoomed): the numbers follow.
	public void EditorPainted() {
		if (!Visible) return;
		FitWidth();
		Invalidate();
	}

	// ---- which lines count

	List<(int Start, int End)> Skipped() {
		if (skipped != null) return skipped;
		var doc = editor.TextDocument;
		int length = doc.Range2(0, 0).StoryLength;
		var ranges = new List<(int Start, int End)>();
		foreach (var table in RtfTables.Find(doc)) ranges.Add((table.Start, Math.Max(table.Start, table.End - 1)));    // (End is the paragraph after it)
		foreach (int at in PageBreaks.Positions(doc)) ranges.Add((at, Math.Min(length - 1, at + 1)));    // the break and its paragraph mark
		Suppressed(doc, 0, length, ranges);
		ranges.Sort();
		return skipped = ranges;
	}

	// Paragraphs with suppressed numbers between two paragraph boundaries: a range's NoLineNumber is the same for all its
	// paragraphs or undefined, so a document without any is one call, and others are halved until each part is uniform.
	static void Suppressed(ITextDocument2 doc, int start, int end, List<(int Start, int End)> found) {
		if (end <= start) return;
		int value = doc.Range2(start, Math.Max(start, end - 1)).Para.NoLineNumber;
		if (value == tomConstants.tomFalse) return;
		if (value == tomConstants.tomTrue) { found.Add((start, end - 1)); return; }
		var middle = doc.Range2((start + end) / 2, (start + end) / 2);
		middle.Expand(tomParagraph);
		int split = middle.Start > start ? middle.Start : middle.End;
		if (split <= start || split >= end) return;    // one paragraph (undefined: not suppressed)
		Suppressed(doc, start, split, found);
		Suppressed(doc, split, end, found);
	}

	// The editor's lines that have no number, as sorted, merged line ranges (they depend on the wrapping, so they are
	// worked out from the character ranges each time).
	List<(int First, int Last)> SkippedLines() {
		var lines = new List<(int First, int Last)>();
		foreach (var (start, end) in Skipped()) {
			int first = Send(EM_EXLINEFROMCHAR, 0, start), last = Send(EM_EXLINEFROMCHAR, 0, Math.Max(start, end));
			if (lines.Count > 0 && first <= lines[^1].Last + 1) lines[^1] = (lines[^1].First, Math.Max(lines[^1].Last, last));
			else lines.Add((first, last));
		}
		return lines;
	}

	List<int> RestartLines() => numbering.Mode == LineNumbering.Modes.EachPage && pageStarts != null
		? pageStarts.Where(p => p > 0).Select(p => Send(EM_EXLINEFROMCHAR, 0, p)).Distinct().OrderBy(l => l).ToList() : new List<int>();

	/// The number of the editor's line (from 0), or 0 if it has none (off, or a line that doesn't count).
	public int NumberOf(int line) {
		if (!numbering.On) return 0;
		var skippedLines = SkippedLines();
		if (skippedLines.Any(s => line >= s.First && line <= s.Last)) return 0;
		return Counted(line, RestartLines(), skippedLines);
	}

	// A counted line's number: the lines that count from the last restart (the start, or a page's first line) up to it.
	int Counted(int line, List<int> restarts, List<(int First, int Last)> skippedLines) {
		int from = restarts.Where(l => l <= line).DefaultIfEmpty(0).Max();
		int notCounted = skippedLines.Sum(s => Math.Max(0, Math.Min(s.Last, line - 1) - Math.Max(s.First, from) + 1));
		return numbering.Start + (line - from) - notCounted;
	}

	/// The number shown in front of a line (0: none shown: it has none, or isn't a multiple of CountBy).
	public int ShownNumberOf(int line) {
		int number = NumberOf(line);
		return number > 0 && number % numbering.CountBy == 0 ? number : 0;
	}

	/// The line (from 0) with the character at a position.
	public int LineOf(int position) => Send(EM_EXLINEFROMCHAR, 0, position);

	// ---- drawing

	Font NumberFont() {
		float zoom = 1f;
		try { zoom = editor.ZoomFactor; } catch (Exception) { }
		return new Font(editor.Font.FontFamily, editor.Font.SizeInPoints * zoom, FontStyle.Regular, GraphicsUnit.Point);
	}

	// Wide enough for the largest number the document can have, and at least two digits.
	void FitWidth() {
		if (!IsHandleCreated || !editor.IsHandleCreated) return;
		int largest = numbering.Start + Send(EM_GETLINECOUNT);
		using var font = NumberFont();
		int width = TextRenderer.MeasureText(new string('8', Math.Max(2, largest.ToString(CultureInfo.InvariantCulture).Length)), font).Width + LogicalToDeviceUnits(12);
		if (Math.Abs(width - Width) > 1) Width = width;
	}

	protected override void OnPaint(PaintEventArgs e) {
		base.OnPaint(e);
		if (!numbering.On || !editor.IsHandleCreated) return;
		try { Draw(e.Graphics); }
		catch (COMException) { }    // the document can be briefly unavailable while it is being replaced
	}

	void Draw(Graphics g) {
		var doc = editor.TextDocument;
		int count = Send(EM_GETLINECOUNT), first = Send(EM_GETFIRSTVISIBLELINE);
		var skippedLines = SkippedLines();
		var restarts = RestartLines();
		using var font = NumberFont();
		// the numbers' tops are their baselines less the font's ascent (TextRenderer draws from the top of the cell)
		float ascent = font.SizeInPoints * DeviceDpi / 72f * font.FontFamily.GetCellAscent(font.Style) / font.FontFamily.GetEmHeight(font.Style);
		int bottom = PointToScreen(new Point(0, ClientSize.Height)).Y;
		int right = ClientSize.Width - LogicalToDeviceUnits(8);
		int number = Counted(first, restarts, skippedLines);
		int nextSkip = skippedLines.FindIndex(s => s.Last >= first);
		for (int line = first; line < count; line++) {
			if (restarts.Contains(line) && line > first) number = numbering.Start;
			if (nextSkip >= 0 && nextSkip < skippedLines.Count && line >= skippedLines[nextSkip].First) {
				// a run of lines without numbers: on after it (a restart in it applies to the line after it)
				int last = skippedLines[nextSkip].Last;
				if (restarts.Any(l => l > line && l <= last + 1)) number = numbering.Start;
				line = last;
				nextSkip++;
				continue;
			}
			int at = Send(EM_LINEINDEX, line);
			if (at < 0) break;
			doc.Range2(at, at).GetPoint(0x20 | 24 | 0x200, out int x, out int baseline);    // tomStart | TA_BASELINE | tomAllowOffClient (screen)
			if (baseline - ascent > bottom) break;
			if (number % numbering.CountBy == 0) {
				int y = PointToClient(new Point(x, baseline)).Y - (int)Math.Round(ascent);
				string text = number.ToString(CultureInfo.CurrentCulture);
				int width = TextRenderer.MeasureText(g, text, font, Size.Empty, TextFormatFlags.NoPadding).Width;
				TextRenderer.DrawText(g, text, font, new Point(right - width, y), SystemColors.GrayText, TextFormatFlags.NoPadding);
			}
			number++;
		}
	}
}

// Layout > Line Numbers > Line Numbering Options: Word's Line Numbers dialog.
sealed class LineNumbersDialog : Form
{
	readonly CheckBox on = new() { Text = "&Add line numbering", AutoSize = true };
	readonly NumericUpDown start = new() { Minimum = 1, Maximum = 32767, Width = 80 };
	readonly NumericUpDown countBy = new() { Minimum = 1, Maximum = 100, Width = 80 };
	readonly TextBox distance = new() { Width = 80 };
	readonly RadioButton eachPage = new() { Text = "Restart each &page", AutoSize = true };
	readonly RadioButton continuous = new() { Text = "&Continuous", AutoSize = true };
	readonly bool metric = RegionInfo.CurrentRegion.IsMetric;

	public LineNumbersDialog(LineNumbering numbering) {
		Text = "Line Numbers";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = MaximizeBox = false;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Font = SystemFonts.MessageBoxFont;
		Padding = new Padding(10);
		on.Checked = numbering.On;
		start.Value = numbering.Start;
		countBy.Value = numbering.CountBy;
		distance.Text = numbering.Distance == 0 ? "Auto" : metric
			? (numbering.Distance * 2.54 / 1440).ToString("0.##", CultureInfo.CurrentCulture) + " cm"
			: (numbering.Distance / 1440.0).ToString("0.##", CultureInfo.CurrentCulture) + " in";
		(numbering.Mode == LineNumbering.Modes.EachPage ? eachPage : continuous).Checked = true;

		var layout = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Dock = DockStyle.Fill };
		Controls.Add(layout);
		layout.Controls.Add(on);
		var options = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Margin = new Padding(20, 3, 3, 3) };
		void Option(string label, Control control) { options.Controls.Add(Label(label)); options.Controls.Add(control); }
		Option("&Start at:", start);
		Option("From &text:", distance);
		Option("Count &by:", countBy);
		options.Controls.Add(Label("Numbering:"));
		var modes = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
		modes.Controls.AddRange(new Control[] { eachPage, continuous });
		options.Controls.Add(modes);
		layout.Controls.Add(options);
		layout.Controls.Add(new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(360), 0), ForeColor = SystemColors.GrayText, Margin = new Padding(3, 8, 3, 3),
			Text = "From text: Auto, or e.g. 0.3 in or 0.8 cm (for Word and LibreOffice). Count by 5 shows 5, 10, 15 ... "
				+ "Lines in tables don't count. Saved with the document; RTE Writer doesn't print line numbers yet (Word and LibreOffice do)."
		});
		void Enable() => options.Enabled = on.Checked;
		on.CheckedChanged += (s, e) => Enable();
		Enable();

		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var buttons = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Anchor = AnchorStyles.Right };
		buttons.Controls.AddRange(new Control[] { ok, cancel });
		layout.Controls.Add(buttons);
		AcceptButton = ok;
		CancelButton = cancel;
		FormClosing += (s, e) => {
			if (DialogResult == DialogResult.OK && on.Checked && Distance() == null) {
				MessageBox.Show(this, $"\"{distance.Text}\" is not a distance: use Auto, or e.g. 0.3 in, 0.8 cm or 20 pt.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
				e.Cancel = true;
			}
		};
	}

	static Label Label(string text) => new() { Text = text, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 3, 3) };

	// Twips (0: automatic), or null if the text isn't a distance. A plain number is in the region's unit.
	int? Distance() {
		string text = distance.Text.Trim().ToLowerInvariant().Replace(" ", "").Replace(',', '.');
		if (text is "" or "auto") return 0;
		if (Regex.IsMatch(text, @"^[\d.]+$")) text += metric ? "cm" : "in";
		return RtfTables.CssLength(text) is int twips && twips >= 0 && twips <= 31680 ? twips : null;
	}

	public LineNumbering Result(LineNumbering current) {
		var result = current.Copy();
		result.Mode = !on.Checked ? LineNumbering.Modes.None : eachPage.Checked ? LineNumbering.Modes.EachPage : LineNumbering.Modes.Continuous;
		if (on.Checked) {
			result.Start = (int)start.Value;
			result.CountBy = (int)countBy.Value;
			result.Distance = Distance() ?? current.Distance;
		}
		return result.Normalized();
	}
}
