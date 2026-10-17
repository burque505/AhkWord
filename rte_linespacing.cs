// Line spacing for RTE Writer (rte_writer23.cs, through #load): the selected paragraphs' line spacing, as in Word's
// Line and Paragraph Spacing button (1.0, 1.15, 1.5, 2.0 ... lines, or at least / exactly so many points). It is a
// paragraph format (TOM's SetLineSpacing), kept by RichEdit in RTF (\sl \slmult) and so in Word and OpenDocument files
// made by Word or LibreOffice. RichEditorNET's HTML leaves it out, so:
//   ExportHtml   gives each paragraph's block (one per paragraph, as LineBreaks.ExportHtml leaves them) a CSS
//                line-height: a percentage for lines, points for exactly, points and Word's mso-line-height-rule:
//                at-least for at least; paragraphs in tables are left alone
//   MarkHtml     turns line-height on blocks into a marker in front of their text, and FromMarkers makes the loaded
//                paragraphs so spaced (as line and page breaks are carried through loading)
// Markdown has no line spacing.
using System.Drawing;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static class LineSpacing
{
	// TOM's line spacing rules
	public const int Single = 0, OneAndHalf = 1, Double = 2, AtLeast = 3, Exactly = 4, Multiple = 5;

	/// A rule and its value as text: "1.5" (lines), "exactly 20pt", "at least 12pt".
	public static string Describe(int rule, float spacing) => rule switch {
		Single => "1",
		OneAndHalf => "1.5",
		Double => "2",
		Multiple => spacing.ToString("0.##", CultureInfo.InvariantCulture),
		Exactly => "exactly " + spacing.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
		AtLeast => "at least " + spacing.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
		_ => "mixed"
	};

	/// "1.15" (lines), "exactly 14pt", "at least 12" (points) as a rule and its value.
	public static (int Rule, float Spacing) Parse(string text) {
		string t = text.Trim().ToLowerInvariant().Replace(',', '.');
		var m = Regex.Match(t, @"^(exactly|at\s*least|atleast)?\s*([\d.]+)\s*(pt|lines?)?$");
		if (!m.Success || !float.TryParse(m.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out float value) || value <= 0)
			throw new ArgumentException($"\"{text}\" is not a line spacing: use a number of lines (1, 1.15, 1.5, 2), or \"exactly 14pt\" / \"at least 12pt\".");
		if (m.Groups[1].Success) {
			if (value > 1584) throw new ArgumentException("Line spacing in points goes up to 1584.");
			return (m.Groups[1].Value.StartsWith("exactly") ? Exactly : AtLeast, value);
		}
		if (m.Groups[3].Value == "pt") return (Exactly, value);
		if (value > 132) throw new ArgumentException("Line spacing goes up to 132 lines.");
		return Math.Abs(value - 1) < 0.001 ? (Single, 0) : (Multiple, value);
	}

	/// Sets the line spacing of the paragraphs a range touches (one undo step).
	public static void Apply(ITextRange2 range, int rule, float spacing) => range.Para.SetLineSpacing(rule, rule == Single ? 1 : spacing);

	/// The line spacing of the paragraphs a range touches ("mixed" if they differ).
	public static string Of(ITextRange2 range) {
		var para = range.Para;
		int rule = para.LineSpacingRule;
		return rule < 0 ? "mixed" : Describe(rule, para.LineSpacing);
	}

	// ================================================================ HTML

	static readonly Regex BlockStart = new(@"<(p|h[1-6]|li|pre)\b[^>]*>", RegexOptions.IgnoreCase);
	const char CellEnd = '\u0007';

	static string Story(ITextDocument2 doc) => doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";

	// A paragraph's line spacing as CSS (null: single, nothing to write).
	static string? Css(int rule, float spacing) => rule switch {
		OneAndHalf => "line-height: 150%",
		Double => "line-height: 200%",
		Multiple when Math.Abs(spacing - 1) > 0.001 => "line-height: " + Math.Round(spacing * 100).ToString(CultureInfo.InvariantCulture) + "%",
		Exactly => "line-height: " + spacing.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
		AtLeast => "line-height: " + spacing.ToString("0.##", CultureInfo.InvariantCulture) + "pt; mso-line-height-rule: at-least",
		_ => null
	};

	/// RichEditorNET's HTML (with one block per paragraph, as LineBreaks.ExportHtml leaves it) with each paragraph's
	/// line spacing as CSS. If its blocks don't match the document's paragraphs, html comes back as it was.
	public static string ExportHtml(RichEditBox editor, string html) {
		var doc = editor.TextDocument;
		string story = Story(doc);
		// the paragraphs as the library writes them: the text up to each paragraph mark, cell end and page break
		var starts = new List<int>();
		int start = 0;
		for (int i = 0; i < story.Length; i++)
			if (story[i] is '\r' or CellEnd or '\f') { starts.Add(start); start = i + 1; }
		var blocks = BlockStart.Matches(html).ToList();
		if (blocks.Count != starts.Count) return html;
		var tables = RtfTables.Find(doc);
		var output = new StringBuilder();
		int position = 0;
		bool any = false;
		for (int k = 0; k < blocks.Count; k++) {
			int at = starts[k];
			if (tables.Any(t => at >= t.Start && at < t.End)) continue;    // (cells: the table's own formatting)
			var para = doc.Range2(at, at).Para;
			if (Css(para.LineSpacingRule, para.LineSpacing) is not { } css) continue;
			var block = blocks[k];
			output.Append(html, position, block.Index - position);
			var style = Regex.Match(block.Value, "\\sstyle=\"([^\"]*)\"", RegexOptions.IgnoreCase);
			output.Append(style.Success
				? block.Value.Insert(style.Groups[1].Index + style.Groups[1].Length, (style.Groups[1].Value.TrimEnd().EndsWith(";") || style.Groups[1].Value.Trim() == "" ? " " : "; ") + css)
				: block.Value.Insert(block.Value.Length - 1, $" style=\"{css}\""));
			position = block.Index + block.Length;
			any = true;
		}
		if (!any) return html;
		output.Append(html, position, html.Length - position);
		return output.ToString();
	}

	const string MarkerStart = "@@RTE-LINE-SPACING:", MarkerEnd = "@@";
	static readonly Regex Marker = new(Regex.Escape(MarkerStart) + @"(\d)([\d.]+)" + Regex.Escape(MarkerEnd));
	static readonly Regex StyledBlock = new(@"<(p|h[1-6]|li|div|pre|td|th)\b[^>]*\sstyle=""([^""]*line-height[^""]*)""[^>]*>", RegexOptions.IgnoreCase);

	/// HTML to load, with the line-height of its blocks as markers in front of their text (FromMarkers applies them
	/// once it is loaded). Lengths: percentages and plain numbers are lines; pt, px and em, points (em: of 12 pt).
	public static string MarkHtml(string html) => StyledBlock.Replace(html, m => {
		string css = m.Groups[2].Value;
		var height = Regex.Match(css, @"line-height\s*:\s*([\d.]+)\s*(%|pt|px|em)?", RegexOptions.IgnoreCase);
		if (!height.Success || !float.TryParse(height.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out float value) || value <= 0) return m.Value;
		bool atLeast = Regex.IsMatch(css, @"mso-line-height-rule\s*:\s*at-least", RegexOptions.IgnoreCase);
		(int rule, float spacing) = height.Groups[2].Value.ToLowerInvariant() switch {
			"%" => (Multiple, value / 100),
			"pt" => (atLeast ? AtLeast : Exactly, value),
			"px" => (atLeast ? AtLeast : Exactly, value * 0.75f),
			"em" => (Multiple, value),
			_ => (Multiple, value)
		};
		if (rule == Multiple && Math.Abs(spacing - 1) < 0.001) return m.Value;
		return m.Value + MarkerStart + rule + spacing.ToString("0.###", CultureInfo.InvariantCulture) + MarkerEnd;
	});

	/// Gives the paragraphs with the markers MarkHtml left their line spacing, and removes the markers; returns how many
	/// there were.
	public static int FromMarkers(RichEditBox editor) {
		var doc = editor.TextDocument;
		int count = 0;
		while (true) {
			var m = Marker.Matches(Story(doc)).LastOrDefault();
			if (m == null) break;
			int rule = m.Groups[1].Value[0] - '0';
			float spacing = float.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
			var marker = doc.Range2(m.Index, m.Index + m.Length);
			marker.Text = "";
			Apply(doc.Range2(m.Index, m.Index), rule, spacing);
			count++;
		}
		return count;
	}
}

// The built-in .docx converter goes through HTML: HtmlToOpenXml writes line-height as Word line spacing but "at least"
// as exactly, and the PowerTools read every paragraph without spacing of its own as 1.08 lines (Word's newer default,
// but no spacing means single). So the .docx's own values are used, paragraph by paragraph, where the paragraphs on
// both sides can be paired up in order (otherwise the HTML's spacing is left out, rather than guessed).
public static class DocxLineSpacing
{
	static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

	// A paragraph's spacing in Word's terms (line in 240ths of a line or twips, and its rule), or null for single.
	static (int Line, string Rule)? Spacing(XElement? spacing) {
		if (spacing?.Attribute(W + "line") is not { } line || !int.TryParse(line.Value, out int value)) return null;
		return (value, (string?)spacing.Attribute(W + "lineRule") ?? "auto");
	}

	/// Each paragraph's line spacing in a .docx, in document order (w:p anywhere in the body): its own, else its style's
	/// (and the styles that one is based on), else the document's default; null: single.
	static List<(int Line, string Rule)?> Paragraphs(string path) {
		using var zip = ZipFile.OpenRead(path);
		XDocument Load(string name) { using var s = zip.GetEntry(name)!.Open(); return XDocument.Load(s); }
		var document = Load("word/document.xml");
		var styles = zip.GetEntry("word/styles.xml") != null ? Load("word/styles.xml") : null;
		var byId = styles?.Descendants(W + "style").Where(s => (string?)s.Attribute(W + "type") == "paragraph" && s.Attribute(W + "styleId") != null)
			.GroupBy(s => (string)s.Attribute(W + "styleId")!).ToDictionary(g => g.Key, g => g.First()) ?? new();
		var defaultStyle = styles?.Descendants(W + "style").FirstOrDefault(s => (string?)s.Attribute(W + "type") == "paragraph" && (string?)s.Attribute(W + "default") is "1" or "true");
		var docDefault = Spacing(styles?.Root?.Element(W + "docDefaults")?.Element(W + "pPrDefault")?.Element(W + "pPr")?.Element(W + "spacing"));
		(int, string)? FromStyle(string? id) {
			for (int depth = 0; id != null && depth < 20 && byId.TryGetValue(id, out var style); depth++) {
				if (Spacing(style.Element(W + "pPr")?.Element(W + "spacing")) is { } s) return s;
				id = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
			}
			return null;
		}
		var list = new List<(int Line, string Rule)?>();
		foreach (var p in document.Root!.Element(W + "body")!.Descendants(W + "p")) {
			var pPr = p.Element(W + "pPr");
			string? styleId = (string?)pPr?.Element(W + "pStyle")?.Attribute(W + "val") ?? (string?)defaultStyle?.Attribute(W + "styleId");
			list.Add(Spacing(pPr?.Element(W + "spacing")) ?? FromStyle(styleId) ?? docDefault);
		}
		return list;
	}

	static string? Css((int Line, string Rule)? spacing) {
		if (spacing is not var (line, rule)) return null;
		string points = (line / 20.0).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
		if (rule == "auto") return Math.Abs(line - 240) <= 1 ? null : "line-height: " + Math.Round(line / 2.4).ToString(CultureInfo.InvariantCulture) + "%";
		return rule == "atLeast" ? "line-height: " + points + "; mso-line-height-rule: at-least" : "line-height: " + points;
	}

	static readonly Regex LineHeightCss = new(@"\s*(line-height|mso-line-height-rule)\s*:[^;]*;?", RegexOptions.IgnoreCase);

	/// The built-in reader's HTML of a .docx (its paragraphs in document order) with the .docx's own line spacing.
	public static void IntoHtml(string path, XElement html) {
		XNamespace x = html.Name.Namespace;
		var blocks = html.Descendants().Where(e => e.Name.Namespace == x && Regex.IsMatch(e.Name.LocalName, "^(p|h[1-6])$")).ToList();
		List<(int Line, string Rule)?>? spacings = null;
		try { spacings = Paragraphs(path); } catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException or NullReferenceException) { }
		bool paired = spacings != null && spacings.Count == blocks.Count;
		for (int i = 0; i < blocks.Count; i++) {
			var block = blocks[i];
			string style = LineHeightCss.Replace((string?)block.Attribute("style") ?? "", "").Trim();
			if (paired && Css(spacings![i]) is { } css) style = (style == "" || style.EndsWith(";") ? style : style + ";") + css + ";";
			if (style == "") block.Attribute("style")?.Remove(); else block.SetAttributeValue("style", style);
		}
	}

	/// Gives a .docx made by the built-in converter the editor's line spacing, on the paragraphs outside tables (paired
	/// up in order with the body's own paragraphs; left as the converter wrote them if they don't pair up).
	public static void FromEditor(string path, RichEditBox editor) {
		var doc = editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var tables = RtfTables.Find(doc);
		var starts = new List<int>();
		int start = 0;
		for (int i = 0; i < story.Length; i++) {
			if (story[i] != '\r') continue;
			if (!tables.Any(t => start >= t.Start && start < t.End)) starts.Add(start);
			start = i + 1;
		}
		XDocument document;
		using (var zip = ZipFile.OpenRead(path)) using (var s = zip.GetEntry("word/document.xml")!.Open()) document = XDocument.Load(s);
		var paragraphs = document.Root!.Element(W + "body")!.Elements(W + "p").ToList();
		// (RichEdit keeps an empty paragraph after a table that ends the document; the converter leaves it out)
		while (starts.Count > paragraphs.Count && story.IndexOf('\r', starts[^1]) == starts[^1]) starts.RemoveAt(starts.Count - 1);
		if (paragraphs.Count != starts.Count) return;
		bool changed = false;
		for (int i = 0; i < paragraphs.Count; i++) {
			var para = doc.Range2(starts[i], starts[i]).Para;
			(int line, string rule)? want = para.LineSpacingRule switch {
				LineSpacing.OneAndHalf => (360, "auto"),
				LineSpacing.Double => (480, "auto"),
				LineSpacing.Multiple => ((int)Math.Round(para.LineSpacing * 240), "auto"),
				LineSpacing.Exactly => ((int)Math.Round(para.LineSpacing * 20), "exact"),
				LineSpacing.AtLeast => ((int)Math.Round(para.LineSpacing * 20), "atLeast"),
				_ => null
			};
			var pPr = paragraphs[i].Element(W + "pPr");
			var spacing = pPr?.Element(W + "spacing");
			if (want is not var (line, rule)) {
				if (spacing?.Attribute(W + "line") != null) { spacing.Attribute(W + "line")!.Remove(); spacing.Attribute(W + "lineRule")?.Remove(); changed = true; }
				continue;
			}
			if (pPr == null) { pPr = new XElement(W + "pPr"); paragraphs[i].AddFirst(pPr); }
			if (spacing == null) {
				spacing = new XElement(W + "spacing");
				// (after pStyle, keepNext ... numPr, pBdr, shd, tabs, suppressAutoHyphens ...: before ind, jc and the rest)
				var next = pPr.Elements().FirstOrDefault(e => e.Name.LocalName is "ind" or "contextualSpacing" or "mirrorIndents" or "suppressOverlap" or "jc"
					or "textDirection" or "textAlignment" or "textboxTightWrap" or "outlineLvl" or "divId" or "cnfStyle" or "rPr" or "sectPr" or "pPrChange");
				if (next != null) next.AddBeforeSelf(spacing); else pPr.Add(spacing);
			}
			spacing.SetAttributeValue(W + "line", line);
			spacing.SetAttributeValue(W + "lineRule", rule);
			changed = true;
		}
		if (!changed) return;
		using var update = ZipFile.Open(path, ZipArchiveMode.Update);
		update.GetEntry("word/document.xml")!.Delete();
		using var stream = update.CreateEntry("word/document.xml", CompressionLevel.Optimal).Open();
		document.Save(stream, SaveOptions.DisableFormatting);
	}
}

// Home > Paragraph > Line Spacing > Line Spacing Options: the rule and its value, as in Word's Paragraph dialog.
sealed class LineSpacingDialog : Form
{
	readonly ComboBox rule = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };
	readonly NumericUpDown value = new() { DecimalPlaces = 2, Increment = 0.25m, Minimum = 0.25m, Maximum = 132, Width = 80 };
	readonly Label unit = new() { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 3, 3) };
	static readonly string[] Rules = { "Single", "1.5 lines", "Double", "At least", "Exactly", "Multiple" };

	public LineSpacingDialog(string current) {
		Text = "Line Spacing";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = MaximizeBox = false;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Font = SystemFonts.MessageBoxFont;
		Padding = new Padding(10);
		rule.Items.AddRange(Rules);
		(int r, float v) = current == "mixed" ? (LineSpacing.Single, 0f) : LineSpacing.Parse(current);
		if (r == LineSpacing.Multiple && Math.Abs(v - 1.5f) < 0.001) r = LineSpacing.OneAndHalf;
		else if (r == LineSpacing.Multiple && Math.Abs(v - 2) < 0.001) r = LineSpacing.Double;
		rule.SelectedIndex = r;
		value.Value = (decimal)Math.Clamp(r is LineSpacing.AtLeast or LineSpacing.Exactly or LineSpacing.Multiple ? v : 1, 0.25f, 132f);
		rule.SelectedIndexChanged += (s, e) => Fit(changed: true);
		Fit(changed: false);

		var layout = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Dock = DockStyle.Fill };
		Controls.Add(layout);
		var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
		row.Controls.AddRange(new Control[] { Label("Line &spacing:"), rule, Label("&At:"), value, unit });
		layout.Controls.Add(row);
		layout.Controls.Add(new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(380), 0), ForeColor = SystemColors.GrayText, Margin = new Padding(3, 8, 3, 3),
			Text = "For the selected paragraphs. Multiple: a number of lines, such as 1.15. At least and Exactly: points. "
				+ "Kept in RTF, Word, OpenDocument and HTML files; Markdown has no line spacing."
		});
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var buttons = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Anchor = AnchorStyles.Right };
		buttons.Controls.AddRange(new Control[] { ok, cancel });
		layout.Controls.Add(buttons);
		AcceptButton = ok;
		CancelButton = cancel;
	}

	static Label Label(string text) => new() { Text = text, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 3, 3) };

	// The value box for the rule: lines for Multiple, points for At least and Exactly, nothing for the fixed ones.
	void Fit(bool changed) {
		int r = rule.SelectedIndex;
		bool points = r is LineSpacing.AtLeast or LineSpacing.Exactly;
		value.Enabled = r is LineSpacing.AtLeast or LineSpacing.Exactly or LineSpacing.Multiple;
		value.Maximum = points ? 1584 : 132;
		value.Increment = points ? 1 : 0.25m;
		unit.Text = points ? "pt" : r == LineSpacing.Multiple ? "lines" : "";
		if (changed && points && value.Value < 6) value.Value = 12;
		if (changed && r == LineSpacing.Multiple && value.Value > 10) value.Value = 1.15m;
	}

	/// The chosen spacing as SetLineSpacing takes it ("1.5", "exactly 14pt" ...).
	public string Result => rule.SelectedIndex switch {
		LineSpacing.Single => "1",
		LineSpacing.OneAndHalf => "1.5",
		LineSpacing.Double => "2",
		LineSpacing.AtLeast => "at least " + value.Value.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
		LineSpacing.Exactly => "exactly " + value.Value.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
		_ => value.Value.ToString("0.##", CultureInfo.InvariantCulture)
	};
}
