// Line and paragraph spacing for RTE Writer (rte_writer36.cs, through #load): the selected paragraphs' line spacing,
// as in Word's Line and Paragraph Spacing button (1.0, 1.15, 1.5, 2.0 ... lines, or at least / exactly so many points),
// and (version 2) their space before and after. They are paragraph formats (TOM's SetLineSpacing, SpaceBefore,
// SpaceAfter), kept by RichEdit in RTF (\sl \slmult \sb \sa) and so in Word and OpenDocument files made by Word or
// LibreOffice. RichEditorNET's HTML leaves them out, so:
//   ExportHtml   gives each paragraph's block (one per paragraph, as LineBreaks.ExportHtml leaves them) a CSS
//                line-height (a percentage for lines, points for exactly, points and Word's mso-line-height-rule:
//                at-least for at least) and margin-top / margin-bottom; paragraphs in tables are left alone
//   MarkHtml     turns those on blocks (and margin-left, margin-right, text-indent and data-rte-tabs, version 2) into
//                markers in front of their text, and FromMarkers gives the loaded paragraphs that format (as line and
//                page breaks are carried through loading); list items keep their list's indents
// Version 2, paragraph formats from Word files (DocxLineSpacing): a .docx's paragraphs get their spacing as Word works
// it out: the document's defaults, the default paragraph style, a table's style (which wins over the default
// paragraph style in its cells, as in Word), the paragraph's own style and its own formatting; "Don't add space between
// paragraphs of the same style" (contextualSpacing) leaves out the space between such paragraphs, which RichEdit can't
// do itself. The built-in reader's HTML gets them as CSS (and tab stops as data-rte-tabs), files read through RTF
// (Word, LibreOffice) get the spacing afterwards, paragraph by paragraph, where the paragraphs pair up.
// Markdown has no line or paragraph spacing.
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
	static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

	/// A rule and its value as text: "1.5" (lines), "exactly 20pt", "at least 12pt".
	public static string Describe(int rule, float spacing) => rule switch {
		Single => "1",
		OneAndHalf => "1.5",
		Double => "2",
		Multiple => spacing.ToString("0.##", Invariant),
		Exactly => "exactly " + spacing.ToString("0.##", Invariant) + "pt",
		AtLeast => "at least " + spacing.ToString("0.##", Invariant) + "pt",
		_ => "mixed"
	};

	/// "1.15" (lines), "exactly 14pt", "at least 12" (points) as a rule and its value.
	public static (int Rule, float Spacing) Parse(string text) {
		string t = text.Trim().ToLowerInvariant().Replace(',', '.');
		var m = Regex.Match(t, @"^(exactly|at\s*least|atleast)?\s*([\d.]+)\s*(pt|lines?)?$");
		if (!m.Success || !float.TryParse(m.Groups[2].Value, NumberStyles.Float, Invariant, out float value) || value <= 0)
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

	/// The space before or after the paragraphs a range touches, in points (-1: they differ).
	public static float SpaceOf(ITextRange2 range, bool before) {
		float value = before ? range.Para.SpaceBefore : range.Para.SpaceAfter;
		return value < 0 ? -1 : value;
	}

	/// Sets the space before or after the paragraphs a range touches, in points (one undo step).
	public static void ApplySpace(ITextRange2 range, bool before, float points) {
		if (points < 0 || points > 1584) throw new ArgumentException("Paragraph spacing is from 0 to 1584 points.");
		if (before) range.Para.SpaceBefore = points; else range.Para.SpaceAfter = points;
	}

	// ================================================================ HTML

	static readonly Regex BlockStart = new(@"<(p|h[1-6]|li|pre)\b[^>]*>", RegexOptions.IgnoreCase);
	const char CellEnd = '\u0007';

	static string Story(ITextDocument2 doc) => doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";

	static string Points(float value) => value.ToString("0.##", Invariant) + "pt";

	// A paragraph's line spacing as CSS (null: single, nothing to write).
	static string? Css(int rule, float spacing) => rule switch {
		OneAndHalf => "line-height: 150%",
		Double => "line-height: 200%",
		Multiple when Math.Abs(spacing - 1) > 0.001 => "line-height: " + Math.Round(spacing * 100).ToString(Invariant) + "%",
		Exactly => "line-height: " + Points(spacing),
		AtLeast => "line-height: " + Points(spacing) + "; mso-line-height-rule: at-least",
		_ => null
	};

	/// RichEditorNET's HTML (with one block per paragraph, as LineBreaks.ExportHtml leaves it) with each paragraph's
	/// line spacing, space before and after, indents and tab stops as CSS (and data-rte-tabs), table cells' paragraphs
	/// included (RtfTables.ExportHtml then puts them in their cells). If its blocks don't match the document's paragraphs, html comes
	/// back as it was.
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
		var output = new StringBuilder();
		int position = 0;
		bool any = false;
		for (int k = 0; k < blocks.Count; k++) {
			int at = starts[k];
			if (at < story.Length && story[at] is '￹' or '￻') continue;    // (the marks that start and end table rows: the row's own format)
			var para = doc.Range2(at, at).Para;
			var parts = new List<string>();
			if (Css(para.LineSpacingRule, para.LineSpacing) is { } lineHeight) parts.Add(lineHeight);
			if (para.SpaceBefore > 0.05f) parts.Add("margin-top: " + Points(para.SpaceBefore));
			if (para.SpaceAfter > 0.05f) parts.Add("margin-bottom: " + Points(para.SpaceAfter));
			// indents and tab stops (not list items: their indents are the list's)
			string tabs = "";
			var block = blocks[k];
			// the line and page breaks: for RTE Writer (data-rte-breaks) and as CSS, which browsers keep when printing
			string breaks = Breaks(para);
			foreach (char flag in breaks)
				parts.Add(flag switch {
					'n' => "page-break-after: avoid; break-after: avoid",
					't' => "page-break-inside: avoid; break-inside: avoid",
					'p' => "page-break-before: always; break-before: page",
					'o' => "widows: 1; orphans: 1",
					_ => ""
				});
			parts.RemoveAll(p => p == "");
			if ((para.ListType & 0xFFFF) == tomConstants.tomListNone && !block.Value.StartsWith("<li", StringComparison.OrdinalIgnoreCase)) {
				if (Math.Abs(para.LeftIndent) > 0.05f) parts.Add("margin-left: " + Points(para.LeftIndent));
				if (Math.Abs(para.RightIndent) > 0.05f) parts.Add("margin-right: " + Points(para.RightIndent));
				if (Math.Abs(para.FirstLineIndent) > 0.05f) parts.Add("text-indent: " + Points(para.FirstLineIndent));
				tabs = TabStops(para);
			}
			if (parts.Count == 0 && tabs == "" && breaks == "") continue;
			string css = string.Join("; ", parts);
			output.Append(html, position, block.Index - position);
			var style = Regex.Match(block.Value, "\\sstyle=\"([^\"]*)\"", RegexOptions.IgnoreCase);
			string tag = css == "" ? block.Value : style.Success
				? block.Value.Insert(style.Groups[1].Index + style.Groups[1].Length, (style.Groups[1].Value.TrimEnd().EndsWith(";") || style.Groups[1].Value.Trim() == "" ? " " : "; ") + css)
				: block.Value.Insert(block.Value.Length - 1, $" style=\"{css}\"");
			if (tabs != "") tag = tag.Insert(tag.Length - 1, $" data-rte-tabs=\"{tabs}\"");
			if (breaks != "") tag = tag.Insert(tag.Length - 1, $" data-rte-breaks=\"{breaks}\"");
			output.Append(tag);
			position = block.Index + block.Length;
			any = true;
		}
		if (!any) return html;
		output.Append(html, position, html.Length - position);
		return output.ToString();
	}

	const string MarkerStart = "@@RTE-LINE-SPACING:", MarkerEnd = "@@", ParaStart = "@@RTE-PARA:";
	static readonly Regex Marker = new(Regex.Escape(MarkerStart) + @"(\d)([\d.]+)" + Regex.Escape(MarkerEnd));
	static readonly Regex ParaMarker = new(Regex.Escape(ParaStart) + @"([^@]*)" + Regex.Escape(MarkerEnd));
	static readonly Regex Tag = new(@"<(/?)(p|h[1-6]|li|ol|ul|div|pre|td|th)\b([^>]*)>", RegexOptions.IgnoreCase);

	// A CSS length in points (null: none, or not a length).
	static float? Length(string css, string property) {
		var m = Regex.Match(css, @"(?:^|;)\s*" + property + @"\s*:\s*(-?[\d.]+)\s*(pt|px|in|cm|mm|em|pc)?\s*(?:;|$)", RegexOptions.IgnoreCase);
		if (!m.Success || !float.TryParse(m.Groups[1].Value, NumberStyles.Float, Invariant, out float value)) return null;
		return value * m.Groups[2].Value.ToLowerInvariant() switch { "px" => 0.75f, "in" => 72f, "cm" => 72 / 2.54f, "mm" => 72 / 25.4f, "em" => 12f, "pc" => 12f, _ => 1f };
	}

	// margin: top right bottom left (one to four lengths) as its longhands
	static (float? Top, float? Right, float? Bottom, float? Left) Margin(string css) {
		var m = Regex.Match(css, @"(?:^|;)\s*margin\s*:\s*([^;]+)", RegexOptions.IgnoreCase);
		if (!m.Success) return (null, null, null, null);
		var values = m.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(v => Length("x:" + v, "x")).ToList();
		return values.Count switch {
			1 => (values[0], values[0], values[0], values[0]),
			2 => (values[0], values[1], values[0], values[1]),
			3 => (values[0], values[1], values[2], values[1]),
			4 => (values[0], values[1], values[2], values[3]),
			_ => (null, null, null, null)
		};
	}

	/// HTML to load, with the line-height, margins, text-indent and tab stops (data-rte-tabs) of its blocks as markers
	/// in front of their text (FromMarkers applies them once it is loaded). Lengths: percentages and plain numbers are
	/// lines; pt, px, in, cm, mm and em (of 12 pt) lengths. List items keep their list's indents.
	public static string MarkHtml(string html) {
		var output = new StringBuilder();
		int position = 0, lists = 0;
		foreach (Match m in Tag.Matches(html)) {
			string name = m.Groups[2].Value.ToLowerInvariant();
			bool closing = m.Groups[1].Value == "/";
			if (name is "ol" or "ul" or "li") { lists += closing ? -1 : 1; if (lists < 0) lists = 0; }
			if (closing || name is "ol" or "ul" or "td" or "th" or "pre") continue;
			string attributes = m.Groups[3].Value;
			string css = Regex.Match(attributes, "\\sstyle=\"([^\"]*)\"", RegexOptions.IgnoreCase) is { Success: true } s ? WebUtilityDecode(s.Groups[1].Value) : "";
			string tabs = Regex.Match(attributes, "\\sdata-rte-tabs=\"([^\"]*)\"", RegexOptions.IgnoreCase) is { Success: true } t ? t.Groups[1].Value : "";
			string breaks = Regex.Match(attributes, "\\sdata-rte-breaks=\"([ntpxo]*)\"", RegexOptions.IgnoreCase) is { Success: true } k ? k.Groups[1].Value : "";
			string markers = "";
			// line spacing (any block)
			var height = Regex.Match(css, @"line-height\s*:\s*([\d.]+)\s*(%|pt|px|em)?", RegexOptions.IgnoreCase);
			if (height.Success && float.TryParse(height.Groups[1].Value, NumberStyles.Float, Invariant, out float value) && value > 0) {
				bool atLeast = Regex.IsMatch(css, @"mso-line-height-rule\s*:\s*at-least", RegexOptions.IgnoreCase);
				(int rule, float spacing) = height.Groups[2].Value.ToLowerInvariant() switch {
					"%" => (Multiple, value / 100), "pt" => (atLeast ? AtLeast : Exactly, value), "px" => (atLeast ? AtLeast : Exactly, value * 0.75f), _ => (Multiple, value)
				};
				if (!(rule == Multiple && Math.Abs(spacing - 1) < 0.001)) markers += MarkerStart + rule + spacing.ToString("0.###", Invariant) + MarkerEnd;
			}
			// space before and after (paragraphs, headings, list items), indents and tab stops (paragraphs and headings
			// outside lists): divs are page breaks and other wrappers
			if (name != "div") {
				var (top, right, bottom, left) = Margin(css);
				top = Length(css, "margin-top") ?? top;
				bottom = Length(css, "margin-bottom") ?? bottom;
				var fields = new List<string>();
				if (top is float b && b > 0.05f) fields.Add("b" + b.ToString("0.##", Invariant));
				if (breaks != "") fields.Add("k" + breaks);
				if (bottom is float a && a > 0.05f) fields.Add("a" + a.ToString("0.##", Invariant));
				if (name != "li" && lists == 0) {
					left = Length(css, "margin-left") ?? left;
					right = Length(css, "margin-right") ?? right;
					float? first = Length(css, "text-indent");
					if (left is float l && Math.Abs(l) > 0.05f) fields.Add("l" + l.ToString("0.##", Invariant));
					if (right is float rr && Math.Abs(rr) > 0.05f) fields.Add("r" + rr.ToString("0.##", Invariant));
					if (first is float f && Math.Abs(f) > 0.05f) fields.Add("f" + f.ToString("0.##", Invariant));
					// (positions left of the paragraph's edge too, LibreOffice's: AddTab leaves those out, and keeps the rest)
					if (Regex.IsMatch(tabs, @"^-?[\d.]+/\d/\d(,-?[\d.]+/\d/\d)*$")) fields.Add("t" + tabs);
				}
				if (fields.Count > 0 || attributes.Contains("data-rte-keep=", StringComparison.OrdinalIgnoreCase)) markers += ParaStart + string.Join(";", fields) + MarkerEnd;
			}
			if (markers == "") continue;
			output.Append(html, position, m.Index + m.Length - position).Append(markers);
			position = m.Index + m.Length;
		}
		if (position == 0) return html;
		output.Append(html, position, html.Length - position);
		return output.ToString();
	}

	/// A paragraph's line and page breaks as data-rte-breaks has them: n keep with next, t keep lines together, p page
	/// break before, x no line numbers, o no widow and orphan control ("": none of them, widow control on).
	public static string Breaks(ITextPara para) =>
		(para.KeepWithNext == tomConstants.tomTrue ? "n" : "") + (para.KeepTogether == tomConstants.tomTrue ? "t" : "")
		+ (para.PageBreakBefore == tomConstants.tomTrue ? "p" : "") + (para.NoLineNumber == tomConstants.tomTrue ? "x" : "")
		+ (StyleEngine.WidowControlOf(para) ? "" : "o");

	/// Adds a tab stop, unless RichEdit can't have it (at or left of the paragraph's edge: LibreOffice's, relative to the
	/// paragraph's indent, can be): returns whether it was added.
	public static bool AddTab(ITextPara para, float position, int alignment, int leader) {
		if (position <= 0 || position > 1584) return false;
		try { para.AddTab(position, alignment, leader); return true; }
		catch (ArgumentException) { return false; }
	}

	/// A paragraph's tab stops as data-rte-tabs has them: position (points) / TOM alignment / TOM leader, ... ("": none).
	public static string TabStops(ITextPara para) {
		var tabs = new List<string>();
		for (int i = 0; i < para.TabCount; i++) {
			para.GetTab(i, out float position, out int alignment, out int leader);
			tabs.Add(position.ToString("0.##", Invariant) + "/" + alignment + "/" + leader);
		}
		return string.Join(",", tabs);
	}

	static string WebUtilityDecode(string s) => System.Net.WebUtility.HtmlDecode(s);

	/// Gives the paragraphs with the markers MarkHtml left their line spacing and paragraph format, and removes the
	/// markers; returns how many there were.
	public static int FromMarkers(RichEditBox editor) {
		var doc = editor.TextDocument;
		// (all of them found once, and taken out from the last: the text before each stays where it was)
		string story = Story(doc);
		var markers = Marker.Matches(story).Select(m => (Match: m, Line: true)).Concat(ParaMarker.Matches(story).Select(m => (Match: m, Line: false)))
			.OrderByDescending(m => m.Match.Index).ToList();
		foreach (var (m, line) in markers) {
			doc.Range2(m.Index, m.Index + m.Length).Text = "";
			var at = doc.Range2(m.Index, m.Index);
			// (a format the editor can't take is left out, not the rest of the document's)
			try {
				if (line) Apply(at, m.Groups[1].Value[0] - '0', float.Parse(m.Groups[2].Value, Invariant));
				else ApplyFormat(at, m.Groups[1].Value);
			} catch (Exception ex) when (ex is ArgumentException or FormatException or System.Runtime.InteropServices.COMException) { }
		}
		return markers.Count;
	}

	// "b12;a0;l36;r0;f-18;t45/0/0,468/2/0;knt": before, after, left, right, first-line indent (points), tab stops
	// (position / alignment / leader, TOM's values), the line and page breaks (as Breaks gives them)
	static void ApplyFormat(ITextRange2 at, string fields) {
		var para = at.Para;
		float? left = null, right = null, first = null;
		foreach (string field in fields.Split(';', StringSplitOptions.RemoveEmptyEntries)) {
			char kind = field[0];
			string value = field[1..];
			if (kind == 'k') {
				// (each set only if it differs: each change makes RichEdit format the paragraph again)
				void Flag(int now, bool on, Action<int> set) { int want = on ? tomConstants.tomTrue : tomConstants.tomFalse; if (now != want) set(want); }
				Flag(para.KeepWithNext, value.Contains('n'), v => para.KeepWithNext = v);
				Flag(para.KeepTogether, value.Contains('t'), v => para.KeepTogether = v);
				Flag(para.PageBreakBefore, value.Contains('p'), v => para.PageBreakBefore = v);
				Flag(para.NoLineNumber, value.Contains('x'), v => para.NoLineNumber = v);
				if (StyleEngine.WidowControlOf(para) == value.Contains('o')) StyleEngine.SetWidowControl(para, !value.Contains('o'));
				continue;
			}
			if (kind == 't') {
				para.ClearAllTabs();
				foreach (var tab in value.Split(',')) {
					var p = tab.Split('/');
					AddTab(para, float.Parse(p[0], Invariant), int.Parse(p[1]), int.Parse(p[2]));
				}
				continue;
			}
			if (!float.TryParse(value, NumberStyles.Float, Invariant, out float points)) continue;
			switch (kind) {
				case 'b': if (Math.Abs(para.SpaceBefore - points) > 0.01f) para.SpaceBefore = points; break;
				case 'a': if (Math.Abs(para.SpaceAfter - points) > 0.01f) para.SpaceAfter = points; break;
				case 'l': left = points; break;
				case 'r': right = points; break;
				case 'f': first = points; break;
			}
		}
		if (left != null || right != null || first != null) {
			// (SetIndents: the first line's indent relative to the left one, the left and the right indent)
			float f = first ?? para.FirstLineIndent, l = left ?? para.LeftIndent, r = right ?? para.RightIndent;
			if (Math.Abs(f - para.FirstLineIndent) > 0.01f || Math.Abs(l - para.LeftIndent) > 0.01f || Math.Abs(r - para.RightIndent) > 0.01f) para.SetIndents(f, l, r);
		}
	}
}

// Paragraph formats from Word files. The built-in .docx converter goes through HTML: HtmlToOpenXml writes line-height
// as Word line spacing but "at least" as exactly, and the PowerTools read every paragraph without spacing of its own
// as 1.08 lines and without its space before and after; Word's own RTF (and LibreOffice's) has the spacing, but
// RichEdit can't leave out the space between paragraphs of the same style (contextualSpacing), nor give a table's
// paragraphs its style's spacing. So each paragraph's format is worked out from the .docx, as Word does it, and used
// where the paragraphs pair up in order (otherwise left as the converter made it, rather than guessed).
public static class DocxLineSpacing
{
	static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

	/// A paragraph's format as Word works it out (twips; Line in 240ths of a line or twips, by Rule).
	public sealed class Format
	{
		public int Before, After;
		public int? Line;
		public string Rule = "auto";
		public bool Contextual;
		public int? Left, Right, FirstLine;    // (FirstLine: negative for a hanging indent)
		public List<(int Position, string Alignment, string Leader)> Tabs = new();
		public bool TabsSet, List;
		public bool KeepNext, KeepLines, PageBreakBefore, WidowControl;    // (the line and page breaks)
		public string Style = "";
		public XElement? Container;            // the body or the table cell the paragraph is in
		public bool InTable;
	}

	// One layer of formatting (a style's pPr or a paragraph's own) over what the layers below gave.
	static void Merge(Format format, XElement? pPr) {
		if (pPr == null) return;
		if (pPr.Element(W + "spacing") is { } spacing) {
			if (Int(spacing, "before") is int before) format.Before = before;
			if (Int(spacing, "after") is int after) format.After = after;
			if (Int(spacing, "line") is int line) { format.Line = line; format.Rule = (string?)spacing.Attribute(W + "lineRule") ?? "auto"; }
			if ((string?)spacing.Attribute(W + "beforeAutospacing") is "1" or "true") format.Before = 280;    // (HTML's 14 pt)
			if ((string?)spacing.Attribute(W + "afterAutospacing") is "1" or "true") format.After = 280;
		}
		if (pPr.Element(W + "contextualSpacing") is { } contextual) format.Contextual = (string?)contextual.Attribute(W + "val") is not ("0" or "false");
		if (pPr.Element(W + "ind") is { } ind) {
			if ((Int(ind, "left") ?? Int(ind, "start")) is int left) format.Left = left;
			if ((Int(ind, "right") ?? Int(ind, "end")) is int right) format.Right = right;
			if (Int(ind, "hanging") is int hanging) format.FirstLine = -hanging;
			else if (Int(ind, "firstLine") is int first) format.FirstLine = first;
		}
		if (pPr.Element(W + "tabs") is { } tabs) {
			format.TabsSet = true;
			foreach (var tab in tabs.Elements(W + "tab")) {
				if (Int(tab, "pos") is not int pos) continue;
				string val = (string?)tab.Attribute(W + "val") ?? "left";
				format.Tabs.RemoveAll(t => t.Position == pos);
				if (val != "clear") format.Tabs.Add((pos, val, (string?)tab.Attribute(W + "leader") ?? "none"));
			}
		}
		if (pPr.Element(W + "numPr") != null) format.List = true;
		if (pPr.Element(W + "keepNext") is { } keepNext) format.KeepNext = On(keepNext);
		if (pPr.Element(W + "keepLines") is { } keepLines) format.KeepLines = On(keepLines);
		if (pPr.Element(W + "pageBreakBefore") is { } breakBefore) format.PageBreakBefore = On(breakBefore);
		if (pPr.Element(W + "widowControl") is { } widow) format.WidowControl = On(widow);
	}

	/// A WordprocessingML on/off element: on unless its val says off.
	public static bool On(XElement element) => (string?)element.Attribute(W + "val") is not ("0" or "false" or "off");

	/// A paragraph's (or style's) pPr over a format (DocxStyles works out styles with it).
	public static void MergeFormat(Format format, XElement pPr) => Merge(format, pPr);

	/// Puts a child into a w:pPr in its place, in the schema's order.
	public static void PlaceIn(XElement pPr, XElement child) => Place(pPr, child);

	static int? Int(XElement element, string name) => int.TryParse((string?)element.Attribute(W + name), NumberStyles.Integer, Invariant, out int value) ? value : null;

	/// Each paragraph's format in a .docx, in document order (w:p in the body, its tables included; text boxes left
	/// out), with contextualSpacing applied.
	public static List<Format> Paragraphs(string path) {
		using var zip = ZipFile.OpenRead(path);
		XDocument Load(string name) { using var s = zip.GetEntry(name)!.Open(); return XDocument.Load(s); }
		var document = Load("word/document.xml");
		var styles = zip.GetEntry("word/styles.xml") != null ? Load("word/styles.xml") : null;
		Dictionary<string, XElement> Styles(string type) => styles?.Descendants(W + "style").Where(s => (string?)s.Attribute(W + "type") == type && s.Attribute(W + "styleId") != null)
			.GroupBy(s => (string)s.Attribute(W + "styleId")!).ToDictionary(g => g.Key, g => g.First()) ?? new();
		var paragraphStyles = Styles("paragraph");
		var tableStyles = Styles("table");
		string? defaultStyle = (string?)styles?.Descendants(W + "style").FirstOrDefault(s => (string?)s.Attribute(W + "type") == "paragraph" && (string?)s.Attribute(W + "default") is "1" or "true")?.Attribute(W + "styleId");
		var docDefault = styles?.Root?.Element(W + "docDefaults")?.Element(W + "pPrDefault")?.Element(W + "pPr");
		// a style and the styles it is based on, the base first
		List<XElement> Chain(Dictionary<string, XElement> byId, string? id) {
			var chain = new List<XElement>();
			for (int depth = 0; id != null && depth < 20 && byId.TryGetValue(id, out var style); depth++) {
				chain.Insert(0, style);
				id = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
			}
			return chain;
		}
		var list = new List<Format>();
		foreach (var p in document.Root!.Element(W + "body")!.Descendants(W + "p")) {
			if (p.Ancestors(W + "txbxContent").Any()) continue;
			var pPr = p.Element(W + "pPr");
			string style = (string?)pPr?.Element(W + "pStyle")?.Attribute(W + "val") ?? defaultStyle ?? "";
			var table = p.Ancestors(W + "tbl").FirstOrDefault();
			var format = new Format { Style = style, InTable = table != null, Container = (XElement?)p.Ancestors(W + "tc").FirstOrDefault() ?? p.Parent };
			Merge(format, docDefault);
			// the default paragraph style, then the table's style (which wins over it), then the paragraph's own style
			if (defaultStyle != null) foreach (var s in Chain(paragraphStyles, defaultStyle)) Merge(format, s.Element(W + "pPr"));
			if (table?.Element(W + "tblPr")?.Element(W + "tblStyle")?.Attribute(W + "val")?.Value is { } tableStyle)
				foreach (var s in Chain(tableStyles, tableStyle)) Merge(format, s.Element(W + "pPr"));
			if (style != defaultStyle) foreach (var s in Chain(paragraphStyles, style)) if ((string?)s.Attribute(W + "styleId") != defaultStyle) Merge(format, s.Element(W + "pPr"));
			Merge(format, pPr);
			list.Add(format);
		}
		// Don't add space between paragraphs of the same style: a paragraph that says so gets no space after (before)
		// when the next (previous) one, in the same place, has its style
		var after = list.Select(f => f.After).ToList();
		var before = list.Select(f => f.Before).ToList();
		for (int i = 0; i + 1 < list.Count; i++) {
			if (list[i].Container != list[i + 1].Container || list[i].Style != list[i + 1].Style) continue;
			if (list[i].Contextual) after[i] = 0;
			if (list[i + 1].Contextual) before[i + 1] = 0;
		}
		for (int i = 0; i < list.Count; i++) { list[i].After = after[i]; list[i].Before = before[i]; }
		return list;
	}

	static string Points(int twips) => (twips / 20.0).ToString("0.##", Invariant) + "pt";

	// A paragraph's spacing as CSS, for the editor's HTML import (MarkHtml reads it).
	static string Css(Format f) {
		var parts = new List<string> { "margin-top: " + Points(f.Before), "margin-bottom: " + Points(f.After) };
		if (f.Line is int line) {
			if (f.Rule == "auto") { if (Math.Abs(line - 240) > 1) parts.Add("line-height: " + Math.Round(line / 2.4).ToString(Invariant) + "%"); }
			else parts.Add("line-height: " + Points(line) + (f.Rule == "atLeast" ? "; mso-line-height-rule: at-least" : ""));
		}
		return string.Join("; ", parts);
	}

	// Tab stops as data-rte-tabs: position (points) / TOM alignment / TOM leader
	static string Tabs(Format f) => string.Join(",", f.Tabs.OrderBy(t => t.Position).Select(t =>
		(t.Position / 20.0).ToString("0.##", Invariant) + "/" + (t.Alignment switch { "center" => 1, "right" or "end" => 2, "decimal" => 3, "bar" => 4, _ => 0 })
		+ "/" + (t.Leader switch { "dot" => 1, "hyphen" => 2, "underscore" or "heavy" => 3, "middleDot" => 1, _ => 0 })));

	static readonly Regex IndentCss = new(@"\s*(margin-left|margin-right|text-indent)\s*:[^;]*;?", RegexOptions.IgnoreCase);
	static readonly Regex SpacingCss = new(@"\s*(line-height|mso-line-height-rule|margin-top|margin-bottom)\s*:[^;]*;?", RegexOptions.IgnoreCase);

	/// The built-in reader's HTML of a .docx (its paragraphs in document order) with the .docx's own spacing, indents
	/// (list items' are the reader's) and tab stops. Without a pairing, the reader's spacing is left out rather than
	/// trusted, and its indents are kept.
	public static void IntoHtml(string path, XElement html) {
		XNamespace x = html.Name.Namespace;
		var blocks = html.Descendants().Where(e => e.Name.Namespace == x && Regex.IsMatch(e.Name.LocalName, "^(p|h[1-6])$")).ToList();
		List<Format>? formats = null;
		try { formats = Paragraphs(path); } catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException or NullReferenceException) { }
		bool paired = formats != null && formats.Count == blocks.Count;
		for (int i = 0; i < blocks.Count; i++) {
			var block = blocks[i];
			string style = SpacingCss.Replace((string?)block.Attribute("style") ?? "", "").Trim();
			if (paired) {
				// (and the exact indents: the PowerTools round them to hundredths of an inch; list items keep the list's)
				if (!formats![i].List) style = IndentCss.Replace(style, "").Trim();
				style = (style == "" || style.EndsWith(";") ? style : style + ";") + Css(formats[i]) + ";";
				if (!formats[i].List) style += $" margin-left: {Points(formats[i].Left ?? 0)}; margin-right: {Points(formats[i].Right ?? 0)}; text-indent: {Points(formats[i].FirstLine ?? 0)};";
				if (formats[i].TabsSet && formats[i].Tabs.Count > 0 && !formats[i].List) block.SetAttributeValue("data-rte-tabs", Tabs(formats[i]));
			}
			if (style == "") block.Attribute("style")?.Remove(); else block.SetAttributeValue("style", style);
			// (Word's empty paragraphs are blank lines; the editor's HTML import drops a paragraph with only a space)
			if (block.Value.Trim() == "" && !block.Descendants(x + "img").Any()) block.SetAttributeValue("data-rte-keep", "1");
		}
		// (and the paragraphs' styles, and their line and page breaks)
		if (paired) DocxStyles.IntoHtml(path, html, blocks, formats!);
	}

	/// A document opened from a .docx through RTF (Word, LibreOffice): its paragraphs' spacing as Word has it (their
	/// indents and tab stops are right in the RTF). Paired up in order: the editor's paragraphs (cells included, the
	/// marks that end table rows left out) with the .docx's; nothing is changed if they don't pair up. Returns whether
	/// they did.
	public static bool IntoEditor(RichEditBox editor, string path) {
		List<Format> formats;
		try { formats = Paragraphs(path); } catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException or NullReferenceException) { return false; }
		var doc = editor.TextDocument;
		var starts = Pieces(doc, formats.Count);
		while (formats.Count > starts.Count && formats.Count > 0) formats.RemoveAt(formats.Count - 1);    // (a last empty paragraph RichEdit merged)
		if (starts.Count != formats.Count) return false;
		doc.BeginEditCollection();
		try {
			for (int i = 0; i < starts.Count; i++) {
				var f = formats[i];
				var para = doc.Range2(starts[i], starts[i]).Para;
				para.SpaceBefore = f.Before / 20f;
				para.SpaceAfter = f.After / 20f;
				if (f.Line is int line) {
					if (f.Rule == "auto") para.SetLineSpacing(Math.Abs(line - 240) <= 1 ? LineSpacing.Single : LineSpacing.Multiple, Math.Abs(line - 240) <= 1 ? 1 : line / 240f);
					else para.SetLineSpacing(f.Rule == "atLeast" ? LineSpacing.AtLeast : LineSpacing.Exactly, line / 20f);
				}
			}
		} finally {
			doc.EndEditCollection();
		}
		return true;
	}

	// Where the editor's paragraphs start, in order as a .docx has them: split at paragraph marks and cell ends, without
	// the marks that start and end table rows; empty ones at the end are left out while there are more than count.
	static List<int> Pieces(ITextDocument2 doc, int count) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var starts = new List<int>();
		int start = 0;
		for (int i = 0; i < story.Length; i++) {
			if (story[i] is not ('\r' or '\u0007')) continue;
			string piece = story[start..i];
			if (!(story[i] == '\r' && piece.Length > 0 && piece.Trim('\uFFF9', '\uFFFB') == "")) starts.Add(start);    // (not a row's start or end)
			start = i + 1;
		}
		while (starts.Count > count && story.IndexOfAny(new[] { '\r', '\u0007' }, starts[^1]) == starts[^1]) starts.RemoveAt(starts.Count - 1);
		return starts;
	}

	// The order of a w:pPr's children, as the schema has it
	static readonly string[] PPrOrder = { "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers",
		"pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi",
		"adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
		"textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange" };

	// Puts a child into a w:pPr in its place, instead of one of the same name
	static void Place(XElement pPr, XElement child) {
		pPr.Element(child.Name)?.Remove();
		int rank = Array.IndexOf(PPrOrder, child.Name.LocalName);
		var next = pPr.Elements().FirstOrDefault(e => Array.IndexOf(PPrOrder, e.Name.LocalName) > rank);
		if (next != null) next.AddBeforeSelf(child); else pPr.Add(child);
	}

	static int Twips(float points) => (int)Math.Round(points * 20);

	/// A .docx's paragraphs (its document.xml's w:p) paired up in order with the editor's (their starts): all the
	/// document's paragraphs, the tables' included; failing that, those outside tables; null if neither pairs up.
	public static (List<XElement> Paragraphs, List<int> Starts)? Pair(ITextDocument2 doc, XDocument document) {
		var body = document.Root!.Element(W + "body")!;
		var paragraphs = body.Descendants(W + "p").Where(p => !p.Ancestors(W + "txbxContent").Any()).ToList();
		var starts = Pieces(doc, paragraphs.Count);
		if (paragraphs.Count != starts.Count) {
			// the paragraphs outside tables
			string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
			var tables = RtfTables.Find(doc);
			starts = new List<int>();
			int start = 0;
			for (int i = 0; i < story.Length; i++) {
				if (story[i] != '\r') continue;
				if (!tables.Any(t => start >= t.Start && start < t.End)) starts.Add(start);
				start = i + 1;
			}
			paragraphs = body.Elements(W + "p").ToList();
			// (RichEdit keeps an empty paragraph after a table that ends the document; the converter leaves it out)
			while (starts.Count > paragraphs.Count && story.IndexOf('\r', starts[^1]) == starts[^1]) starts.RemoveAt(starts.Count - 1);
			if (paragraphs.Count != starts.Count) return null;
		}
		if (paragraphs.Count == 0) return null;
		return (paragraphs, starts);
	}

	/// Gives a .docx made by the built-in converter the editor's paragraph formats: line spacing, space before and after,
	/// indents and tab stops (HtmlToOpenXml writes "at least" as exactly, and has no tab stops). Paired up in order with
	/// all the document's paragraphs, the tables' included; failing that with those outside tables; left as the converter
	/// wrote them if neither pairs up.
	public static void FromEditor(string path, RichEditBox editor) {
		var doc = editor.TextDocument;
		XDocument document;
		using (var zip = ZipFile.OpenRead(path)) using (var s = zip.GetEntry("word/document.xml")!.Open()) document = XDocument.Load(s);
		if (Pair(doc, document) is not { } pair) return;
		var (paragraphs, starts) = pair;
		for (int i = 0; i < paragraphs.Count; i++) {
			var para = doc.Range2(starts[i], starts[i]).Para;
			var pPr = paragraphs[i].Element(W + "pPr");
			if (pPr == null) { pPr = new XElement(W + "pPr"); paragraphs[i].AddFirst(pPr); }
			// line spacing, space before and after
			(int line, string rule) = para.LineSpacingRule switch {
				LineSpacing.OneAndHalf => (360, "auto"),
				LineSpacing.Double => (480, "auto"),
				LineSpacing.Multiple => ((int)Math.Round(para.LineSpacing * 240), "auto"),
				LineSpacing.Exactly => (Twips(para.LineSpacing), "exact"),
				LineSpacing.AtLeast => (Twips(para.LineSpacing), "atLeast"),
				_ => (240, "auto")
			};
			Place(pPr, new XElement(W + "spacing", new XAttribute(W + "before", Twips(Math.Max(0, para.SpaceBefore))), new XAttribute(W + "after", Twips(Math.Max(0, para.SpaceAfter))),
				new XAttribute(W + "line", line), new XAttribute(W + "lineRule", rule)));
			// indents and tab stops (list items keep their list's indents)
			if ((para.ListType & 0xFFFF) != tomConstants.tomListNone || pPr.Element(W + "numPr") != null) continue;
			int left = Twips(para.LeftIndent), right = Twips(para.RightIndent), first = Twips(para.FirstLineIndent);
			pPr.Element(W + "ind")?.Remove();
			if (left != 0 || right != 0 || first != 0) {
				var ind = new XElement(W + "ind");
				if (left != 0) ind.SetAttributeValue(W + "left", left);
				if (right != 0) ind.SetAttributeValue(W + "right", right);
				if (first > 0) ind.SetAttributeValue(W + "firstLine", first);
				else if (first < 0) ind.SetAttributeValue(W + "hanging", -first);
				Place(pPr, ind);
			}
			if (para.TabCount > 0) {
				var tabs = new XElement(W + "tabs");
				for (int t = 0; t < para.TabCount; t++) {
					para.GetTab(t, out float position, out int alignment, out int leader);
					var tab = new XElement(W + "tab", new XAttribute(W + "val", alignment switch { 1 => "center", 2 => "right", 3 => "decimal", 4 => "bar", _ => "left" }));
					string? dots = leader switch { 1 => "dot", 2 => "hyphen", 3 => "underscore", 4 => "heavy", _ => null };
					if (dots != null) tab.SetAttributeValue(W + "leader", dots);
					tab.SetAttributeValue(W + "pos", Twips(position));
					tabs.Add(tab);
				}
				Place(pPr, tabs);
			}
		}
		using var update = ZipFile.Open(path, ZipArchiveMode.Update);
		update.GetEntry("word/document.xml")!.Delete();
		using var stream = update.CreateEntry("word/document.xml", CompressionLevel.Optimal).Open();
		document.Save(stream, SaveOptions.DisableFormatting);
	}
}
