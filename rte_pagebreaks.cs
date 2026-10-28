// Manual page breaks for RichEditBox (used by AhkWord, AhkWord.cs, through #load).
// In the document a page break is a paragraph of its own: a form feed (U+000C) and its paragraph mark, which is how
// RichEdit reads Word's "\page\par" from RTF. RichEdit's printing (EM_FORMATRANGE) starts a new page after the form
// feed, and its RTF writes it as \page, which Word and LibreOffice read as a page break. (TOM sees the form feed as a
// paragraph of its own, sharing its paragraph format with the paragraph mark after it.)
// In HTML and Markdown a page break is
//     <div style="page-break-after: always; break-after: page"></div>
// which browsers obey when printing (or saving as PDF), and which the built-in .docx converter (HtmlToOpenXml) turns
// into a Word page break. Read back from HTML: that, and any element with page-break-before / page-break-after: always
// or break-before / break-after: page in its style, such as Word's <br style="page-break-before: always">.
using System.Text;
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static class PageBreaks
{
	public const char Break = '\f';
	/// A page break in HTML and Markdown.
	public const string Html = "<div style=\"page-break-after: always; break-after: page\"></div>";
	const string Marker = "@@RTE-PAGE-BREAK@@";

	static string Story(ITextDocument2 doc) => doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";

	/// The positions of the document's page breaks (their form feeds).
	public static List<int> Positions(ITextDocument2 doc) {
		string s = Story(doc);
		var positions = new List<int>();
		for (int i = s.IndexOf(Break); i >= 0; i = s.IndexOf(Break, i + 1)) positions.Add(i);
		return positions;
	}

	// ================================================================ editing

	/// Inserts a page break at the selection (which it replaces), as a paragraph of its own, and puts the caret at the
	/// start of the paragraph after it. Returns the break's position. The caller groups this into one undo step.
	public static int Insert(RichEditBox editor) {
		var doc = editor.TextDocument;
		var selection = doc.Selection2;
		if (RtfTables.At(editor, selection.Start) != null || (selection.End > selection.Start && RtfTables.At(editor, selection.End - 1) != null))
			throw new InvalidOperationException("A page break cannot go in a table.");
		if (selection.End > selection.Start) selection.Text = "";
		int at = selection.Start;
		var paragraph = doc.Range2(at, at);
		paragraph.Expand(tomConstants.tomParagraph);
		int storyEnd = doc.Range2(0, 0).StoryLength;
		// in a legal item, the caret in its number counts as its start (the number stays with the item)
		if (LegalLists.IsLegal(paragraph.Para.ListType) && at <= paragraph.Start + LegalLists.Prefix.Match(paragraph.Text ?? "").Length) at = paragraph.Start;
		int breakAt;
		if (at == paragraph.Start) {
			doc.Range2(at, at).Text = "\f\r";                    // before the paragraph
			breakAt = at;
		} else if (at == paragraph.End - 1 && paragraph.End < storyEnd) {
			doc.Range2(at, at).Text = "\r\f";                    // after it: the break takes its paragraph mark
			breakAt = at + 1;
		} else {
			doc.Range2(at, at).Text = "\r\f\r";                  // in it, or at the end of the document: it is split
			breakAt = at + 1;
		}
		MakePlain(editor, breakAt);
		selection.SetRange(breakAt + 2, breakAt + 2);
		return breakAt;
	}

	// A page break's paragraph: not a list item, not indented, in the editor's font (it is still a paragraph in Word).
	static void MakePlain(RichEditBox editor, int at) {
		var range = editor.TextDocument.Range2(at, at + 2);
		var para = range.Para;
		if ((para.ListType & 0xFFFF) != tomConstants.tomListNone) para.ListType = tomConstants.tomListNone;
		para.SetIndents(0, 0, 0);
		para.Alignment = tomConstants.tomAlignLeft;
		var font = range.Font;
		font.Name = editor.Font.Name;
		font.Size = editor.Font.SizeInPoints;
		font.Bold = font.Italic = font.StrikeThrough = font.Superscript = font.Subscript = tomConstants.tomFalse;
		font.Underline = tomConstants.tomNone;
		font.ForeColor = font.BackColor = tomConstants.tomAutoColor;
	}

	/// Makes every form feed a page break paragraph of its own (RTF and plain text can have them anywhere, e.g. Word's
	/// "text\page more"). Form feeds in tables are left alone. Returns how many there are.
	public static int Normalize(RichEditBox editor) {
		var doc = editor.TextDocument;
		var positions = Positions(doc);
		if (positions.Count == 0) return 0;
		var tables = RtfTables.Find(doc);
		string s = Story(doc);
		int count = 0;
		for (int k = positions.Count - 1; k >= 0; k--) {
			int at = positions[k];
			if (tables.Any(t => at >= t.Start && at < t.End)) continue;
			if (at + 1 >= s.Length || s[at + 1] != '\r') doc.Range2(at + 1, at + 1).Text = "\r";
			if (at > 0 && s[at - 1] != '\r' && !Sections.IsMarkerBefore(s, at)) { doc.Range2(at, at).Text = "\r"; at++; }    // (iteration 63: a section break's marker stays with it)
			MakePlain(editor, at);
			count++;
		}
		return count;
	}

	/// Where printing goes on after a page that ended at a page break: after the break's paragraph mark too (or that
	/// empty line would start the next page).
	public static int AfterBreak(ITextDocument2 doc, int next) {
		if (next <= 0) return next;
		var range = doc.Range2(next - 1, next + 1);
		string text = range.Text ?? "";
		return text.Length == 2 && text[0] == Break && text[1] == '\r' ? next + 1 : next;
	}

	// ================================================================ RTF

	static readonly Regex RtfToken = new(@"\\[a-zA-Z]+-?\d* ?|\\[^a-zA-Z]|[{}]|[^\\{}]+");
	static readonly string[] SkippedGroups = { @"\*", @"\stylesheet", @"\fonttbl", @"\colortbl", @"\listtable", @"\listoverridetable", @"\info", @"\pict" };

	/// RTF with LibreOffice's and Word's "page break before" paragraphs (\pagebb, which RichEdit ignores) given a page
	/// break (\page) at their start instead. (Styles and other tables are left alone.)
	public static string FromPageBreakBefore(string rtf) {
		if (!rtf.Contains(@"\pagebb")) return rtf;
		var sb = new StringBuilder(rtf.Length + 64);
		var skip = new Stack<bool>();
		bool skipping = false;
		foreach (Match token in RtfToken.Matches(rtf)) {
			string t = token.Value;
			if (t == "{") {
				skip.Push(skipping);
				string next = rtf.Substring(token.Index + 1, Math.Min(20, rtf.Length - token.Index - 1));
				skipping = skipping || SkippedGroups.Any(g => next.StartsWith(g) && (g == @"\*" || next.Length == g.Length || !char.IsLetter(next[g.Length])));
			} else if (t == "}") {
				skipping = skip.Count > 0 && skip.Pop();
			} else if (!skipping && Regex.IsMatch(t, @"^\\pagebb ?$")) {
				sb.Append(@"\page ");
				continue;
			}
			sb.Append(t);
		}
		return sb.ToString();
	}

	// ================================================================ HTML

	// Empty blocks of RichEditorNET's HTML (a paragraph with nothing but spaces and inline formatting in it), lists,
	// and table row markers (rte_tables.cs), in document order.
	static readonly Regex HtmlUnits = new(
		@"(?<list></?(?:ol|ul)\b[^>]*>)|(?<row>[￹￻])|(?<empty><(?<tag>p|div|pre|h[1-6])\b[^>]*>(?:[ \t\r\n]|</?(?:span|b|i|u|s|strong|em|font|sup|sub|a)\b[^>]*>)*</\k<tag>>)",
		RegexOptions.IgnoreCase);

	/// RichEditorNET's HTML with the page breaks written as Html. It writes a page break as an empty paragraph, like an
	/// empty paragraph of text: its empty blocks outside lists and tables are the document's empty paragraphs there, in
	/// the same order, so those that are page breaks are known. If that ever does not hold, the HTML is left as it is.
	public static string ExportHtml(RichEditBox editor, string html) {
		var doc = editor.TextDocument;
		if (Story(doc).IndexOf(Break) < 0) return html;
		// the document's empty paragraphs outside lists and tables: 1 a page break, 2 its paragraph mark, 0 others
		var empties = new List<int>();
		var tables = RtfTables.Find(doc);
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var range = doc.Range2(0, 0);
		range.Expand(tomConstants.tomParagraph);
		while (true) {
			int start = range.Start;
			if (tables.FirstOrDefault(t => start >= t.Start && start < t.End) is { } table) {
				if (table.End >= storyEnd) break;
				range.SetRange(table.End, table.End);
				range.Expand(tomConstants.tomParagraph);
				continue;
			}
			string text = range.Text ?? "";
			if ((range.Para.ListType & 0xFFFF) == tomConstants.tomListNone) {
				string content = text.TrimEnd('\r');
				if (content == Break.ToString()) empties.Add(1);
				else if (content.Trim(' ', '\t') == "") empties.Add(text == "\r" && empties.Count > 0 && empties[^1] == 1 && LastWasBreak(doc, start) ? 2 : 0);
			}
			if (range.End >= storyEnd || range.End <= start) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}

		var blocks = new List<Match>();
		int lists = 0;
		bool inRow = false;
		foreach (Match unit in HtmlUnits.Matches(html)) {
			if (unit.Groups["list"].Success) { lists += unit.Value[1] == '/' ? -1 : 1; continue; }
			if (unit.Groups["row"].Success) { inRow = unit.Value[0] == '￹'; continue; }
			if (lists <= 0 && !inRow) blocks.Add(unit);
		}
		if (blocks.Count != empties.Count) return html;    // the HTML does not map onto the document

		var sb = new StringBuilder();
		int position = 0;
		for (int b = 0; b < blocks.Count; b++) {
			if (empties[b] == 0) continue;
			sb.Append(html, position, blocks[b].Index - position);
			if (empties[b] == 1) sb.Append(Html);    // (and the break's own paragraph mark goes: it is part of the break)
			position = blocks[b].Index + blocks[b].Length;
		}
		sb.Append(html, position, html.Length - position);
		return sb.ToString();
	}

	static bool LastWasBreak(ITextDocument2 doc, int start) => start > 0 && doc.Range2(start - 1, start).Text == Break.ToString();

	static readonly Regex BreakStyle = new(@"\b(?:page-break-(?<side>before|after)\s*:\s*always|break-(?<side>before|after)\s*:\s*page)\b", RegexOptions.IgnoreCase);
	static readonly Regex OpeningTag = new(@"<([a-zA-Z][a-zA-Z0-9]*)\b([^>]*)>");
	static readonly Regex StyleAttribute = new(@"\bstyle\s*=\s*(""[^""]*""|'[^']*')", RegexOptions.IgnoreCase);

	/// HTML to load, with a marker where each page break is (FromMarkers makes them page breaks after loading).
	public static string MarkHtml(string html) {
		if (!BreakStyle.IsMatch(html)) return html;
		var sb = new StringBuilder();
		int position = 0;
		foreach (Match tag in OpeningTag.Matches(html)) {
			if (tag.Index < position) continue;
			var style = StyleAttribute.Match(tag.Groups[2].Value);
			if (!style.Success || BreakStyle.Match(style.Value) is not { Success: true } breakStyle) continue;
			// (AhkWord's own "page break before" paragraph: the paragraph's setting, which data-rte-breaks gives back)
			if (Regex.IsMatch(tag.Groups[2].Value, "data-rte-breaks=\"[^\"]*p", RegexOptions.IgnoreCase)) continue;
			string name = tag.Groups[1].Value.ToLowerInvariant();
			bool before = breakStyle.Groups["side"].Value.Equals("before", StringComparison.OrdinalIgnoreCase);
			sb.Append(html, position, tag.Index - position);
			position = tag.Index + tag.Length;
			if (name is "br") { sb.Append(Marker); continue; }                   // a break in a paragraph
			if (name is "hr" or "img") { sb.Append("<p>").Append(Marker).Append("</p>"); continue; }
			sb.Append(tag.Value);
			if (before || tag.Value.EndsWith("/>")) { sb.Append(Marker); continue; }    // at the start of the element's text
			int close = ClosingTag(html, name, position);                        // after: at the end of its text
			if (close < 0) { sb.Append(Marker); continue; }
			sb.Append(MarkHtml(html[position..close])).Append(Marker);
			position = close;
		}
		sb.Append(html, position, html.Length - position);
		return sb.ToString();
	}

	// The position of the tag that closes the element whose content starts at from, or -1.
	static int ClosingTag(string html, string name, int from) {
		int depth = 1;
		foreach (Match tag in new Regex(@"<(/?)" + Regex.Escape(name) + @"\b[^>]*>", RegexOptions.IgnoreCase).Matches(html, from)) {
			if (tag.Groups[1].Value == "/") { if (--depth == 0) return tag.Index; }
			else if (!tag.Value.EndsWith("/>")) depth++;
		}
		return -1;
	}

	/// After loading HTML from MarkHtml: each marker becomes a page break before its paragraph (a marker at the start;
	/// a paragraph with nothing else in it becomes the break itself), after it (a marker at the end), or splitting it.
	/// Returns how many there were.
	public static int FromMarkers(RichEditBox editor) {
		var doc = editor.TextDocument;
		int count = 0;
		while (true) {
			string s = Story(doc);
			int m = s.LastIndexOf(Marker, StringComparison.Ordinal);
			if (m < 0) break;
			var paragraph = doc.Range2(m, m);
			paragraph.Expand(tomConstants.tomParagraph);
			int start = paragraph.Start, end = paragraph.End;
			int contentEnd = end > start && s[end - 1] is '\r' or '\v' or '\f' ? end - 1 : end;
			string before = s[start..m], after = s[(m + Marker.Length)..contentEnd];
			doc.Range2(m, m + Marker.Length).Text = "";
			contentEnd -= Marker.Length;
			bool atStart = Invisible(before) || (LegalLists.IsLegal(paragraph.Para.ListType) && LegalLists.Prefix.Match(before) is { Success: true } number && Invisible(before[number.Length..]));
			bool atEnd = Invisible(after);
			// (what is left on the break's side is invisible, such as the PowerTools' left-to-right mark: it goes)
			if (atEnd && after != "" && !(atStart && (paragraph.Para.ListType & 0xFFFF) == tomConstants.tomListNone)) {
				doc.Range2(m, m + after.Length).Text = "";
				contentEnd -= after.Length;
			}
			int prefix = LegalLists.IsLegal(paragraph.Para.ListType) ? LegalLists.Prefix.Match(before).Length : 0;
			if (atStart && !atEnd && m > start + prefix) {
				doc.Range2(start + prefix, m).Text = "";
				contentEnd -= m - start - prefix;
				m = start + prefix;
			}
			int storyEnd = doc.Range2(0, 0).StoryLength;
			int breakAt;
			if (atStart && atEnd && (paragraph.Para.ListType & 0xFFFF) == tomConstants.tomListNone) {
				doc.Range2(start, contentEnd).Text = Break.ToString();       // the paragraph is the break
				breakAt = start;
				if (doc.Range2(start + 1, start + 2).Text != "\r") doc.Range2(start + 1, start + 1).Text = "\r";
			} else if (atStart) {
				doc.Range2(start, start).Text = "\f\r";
				breakAt = start;
			} else if (atEnd && contentEnd + 1 < storyEnd) {
				doc.Range2(contentEnd + 1, contentEnd + 1).Text = "\f\r";    // at the start of the next paragraph
				breakAt = contentEnd + 1;
			} else {
				doc.Range2(m, m).Text = "\r\f\r";
				breakAt = m + 1;
			}
			MakePlain(editor, breakAt);
			count++;
		}
		return count;
	}

	static bool Invisible(string text) => text.Trim(' ', '\t', '\v', '‎', '‏', ' ') == "";

	// ================================================================ Markdown

	/// A line of Markdown that is a page break: Html (or another element with a page break style, alone on its line),
	/// or pandoc's \newpage and \pagebreak.
	public static readonly Regex MarkdownLine = new(
		@"^ {0,3}(?:\\newpage|\\pagebreak|<(?<tag>div|p|hr|br|span)\b[^>]*(?:page-break-(?:before|after)\s*:\s*always|break-(?:before|after)\s*:\s*page)[^>]*>(?:\s*</\k<tag>>)?)\s*$",
		RegexOptions.IgnoreCase);
}
