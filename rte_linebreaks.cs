// RTE Writer: line breaks (Shift+Enter) in HTML, as other editors have them: a <br> inside the paragraph, list item,
// heading or table cell, and <br> read back as a line break, not a new paragraph.
// RichEditorNET's own HTML writes the text on each side of a line break (RichEdit's vertical tab) as a block of its own
// (a heading's first line even as a <p>), and reads <br> as a new paragraph. ExportHtml merges those blocks back into one
// per paragraph (with the paragraph's own tag) before the tables, lists and page breaks are worked out, which count one
// block per paragraph; MarkHtml and FromMarkers carry <br> through loading as a marker, as page breaks are.
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static class LineBreaks
{
	public const char Break = '\v';
	const string Marker = "@@RTE-LINE-BREAK@@";
	const char CellEnd = '\u0007';

	static string Story(ITextDocument2 doc) => doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";

	// A block's start, and where its own text ends: at the next block tag (a list item's text ends where a list nested in
	// it starts, before its </li>).
	static readonly Regex BlockStart = new(@"<(p|h[1-6]|li|pre)\b[^>]*>", RegexOptions.IgnoreCase);
	static readonly Regex BlockTag = new(@"</?(p|h[1-6]|li|pre|ol|ul|table|tr|td|th|div|blockquote)\b[^>]*>", RegexOptions.IgnoreCase);
	static readonly Regex AnyTag = new(@"<[^>]+>");

	sealed record HtmlBlock(int Start, string Open, int ContentStart, int ContentEnd)
	{
		public string Content(string html) => html[ContentStart..ContentEnd];
	}

	static List<HtmlBlock> Blocks(string html) {
		var blocks = new List<HtmlBlock>();
		foreach (Match open in BlockStart.Matches(html)) {
			int contentStart = open.Index + open.Length;
			var next = BlockTag.Match(html, contentStart);
			blocks.Add(new HtmlBlock(open.Index, open.Value, contentStart, next.Success ? next.Index : html.Length));
		}
		return blocks;
	}

	/// RichEditorNET's HTML for the document (html), with the blocks it made of each line of a paragraph with line breaks
	/// merged into the paragraph's block, the lines separated by <br>. If its blocks don't match the document's text as
	/// expected, html comes back as it was.
	public static string ExportHtml(RichEditBox editor, string html) {
		string story = Story(editor.TextDocument);
		if (story.IndexOf(Break) < 0) return html;
		// the document's pieces as the library writes them: the text between paragraph marks, line breaks, cell ends and
		// page breaks (a page break's form feed and its paragraph mark are an empty block each)
		var pieces = new List<(string Text, char End)>();
		var text = new StringBuilder();
		foreach (char c in story) {
			if (c is '\r' or Break or CellEnd or '\f') { pieces.Add((text.ToString(), c)); text.Clear(); }
			else text.Append(c);
		}
		var blocks = Blocks(html);
		if (blocks.Count != pieces.Count) return html;
		var output = new StringBuilder();
		int position = 0, i = 0;
		while (i < blocks.Count) {
			int last = i;
			while (last < blocks.Count - 1 && pieces[last].End == Break) last++;    // the blocks of one paragraph's lines
			if (last == i) { i++; continue; }
			for (int k = i; k <= last; k++) {
				if (!SameText(blocks[k].Content(html), pieces[k].Text)) return html;
				// between two lines' blocks: only the first one's closing tag and the next one's opening tag
				if (k > i && !Regex.IsMatch(html[blocks[k - 1].ContentEnd..blocks[k].Start], @"^\s*</(p|h[1-6]|li|pre)\s*>\s*$", RegexOptions.IgnoreCase)) return html;
			}
			// the paragraph's formatting (and so its tag) is the last line's: RichEdit keeps it with the paragraph mark;
			// what follows the last line's text (its closing tag, or a nested list and then it) stays as it is
			output.Append(html, position, blocks[i].Start - position);
			output.Append(blocks[last].Open).Append(string.Join("<br>", Enumerable.Range(i, last - i + 1).Select(k => blocks[k].Content(html))));
			position = blocks[last].ContentEnd;
			i = last + 1;
		}
		output.Append(html, position, html.Length - position);
		return output.ToString();
	}

	// Whether a block's content is the piece's text (tags left out, entities decoded, pictures and spacing aside).
	static bool SameText(string content, string text) {
		static string Plain(string s) => Regex.Replace(s.Replace("￼", ""), @"\s+", " ").Trim();
		return Plain(WebUtility.HtmlDecode(AnyTag.Replace(content, ""))) == Plain(text);
	}

	// <br> (and the spacing after it); not one that ends its block, which shows nothing in a browser.
	static readonly Regex LineBreak = new(@"<br\b[^>]*>\s*(?!\s*</(p|h[1-6]|li|td|th|div|pre|blockquote)\s*>)", RegexOptions.IgnoreCase);
	static readonly Regex EndingLineBreak = new(@"<br\b[^>]*>(?=\s*</(p|h[1-6]|li|td|th|div|pre|blockquote)\s*>)", RegexOptions.IgnoreCase);

	/// HTML to load, with its line breaks as markers (FromMarkers makes them line breaks once it is loaded). Page breaks
	/// written as <br style="page-break-..."> must have been marked before (PageBreaks.MarkHtml).
	public static string MarkHtml(string html) => LineBreak.Replace(EndingLineBreak.Replace(html, ""), Marker);

	/// Turns the markers MarkHtml left in the loaded document into line breaks; returns how many there were.
	public static int FromMarkers(RichEditBox editor) {
		var doc = editor.TextDocument;
		int count = 0;
		while (true) {
			int m = Story(doc).LastIndexOf(Marker, StringComparison.Ordinal);
			if (m < 0) break;
			doc.Range2(m, m + Marker.Length).Text = Break.ToString();
			count++;
		}
		return count;
	}
}
