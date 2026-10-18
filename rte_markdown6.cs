// Markdown for RichEditBox: a replacement for its Markdown property, used by RTE Writer (rte_writer32.cs, through #load).
// Version 6: line breaks (Shift+Enter) where Markdown has no "\" + newline for them are written as <br>, as GitHub
// does: in table cells and headings (they were written as spaces); and a list item line that ends in two spaces before
// another line is read as a line break (it was read as a space). Tables are written with their columns lined up
// (AlignTable, which the source view also uses to tidy a table being edited).
// Version 5: page breaks (rte_pagebreaks.cs, also #load-ed) are written as <div style="page-break-after: always;
// break-after: page"></div> on a line of its own, which Markdown viewers pass on as HTML and browsers obey when
// printing; read back from that (or any element with a page break style alone on its line), and pandoc's \newpage
// and \pagebreak.
// Version 4: legal numbering (rte_legal.cs): a legal list is a numbered list with [type=legal] at the end of its first
// item (nested lists in it are legal too); the numbers are written as Markdown's own (1., 2., ...), without the
// "1.2.1" text, and are put back when the Markdown is loaded.
// Version 3: GitHub tables become real tables in the document (rte_tables.cs, also #load-ed), and real tables are
// written as GitHub tables (a cell with several paragraphs gets <br> between them).
// Version 2: also keeps the formatting that plain Markdown cannot express, as HTML and M-ArkDown extensions:
//   alignment       <div align="center"> ... </div> around paragraphs and headings (also read: <p align>, <center>,
//                   style="text-align: ...", and single-line <p align="right">text</p> or <h1 align="center">...</h1>)
//   colors, fonts   <span style="color: #C00000; background-color: #FFFF00; font-family: Georgia; font-size: 14pt">
//                   (also read: M-ArkDown <c=#F00>, <font color face size>, <mark>, <b> <i> <s> <em> <strong> <del>)
//   list styles     M-ArkDown [type=a] (1 a A i I) at the end of a list's first item when it is not GitHub's default
//                   for its level (1., then i., then a.); "1)" punctuation; start numbers
//   image sizes     M-ArkDown ![alt](src)(width=200) when the size is not the image's own (at most 600 wide)
//   restarted lists <!-- --> between two lists that follow each other but are numbered separately (Markdown would
//                   join them); in the document they are kept apart by an empty paragraph if they start at the same
//                   number (RichEdit numbers on through items with the same start value)
// Self-contained, so that it can later replace the library's MarkdownFormatter.
//
// Load (Markdown -> document): CommonMark blocks and inlines, plus GitHub tables and ~~strikethrough~~:
//   paragraphs (separated by blank lines; single line breaks are spaces, "two spaces" or "\" at a line end is a line
//   break), # and underlined headings, bullet and numbered lists (nested, with their start number), > quotes (indented
//   paragraphs), --- rules, ``` and indented code blocks, tables (tab-separated rows with aligned tab stops),
//   **bold** *italic* ~~strike~~ `code` [links](url) [reference links][ref] <autolinks> ![images](src) and the inline
//   HTML tags <u> <ins> <sup> <sub> <br>. Images are embedded from data: URIs and local files (relative to the
//   Markdown file); images on the web become links, as the editor cannot download them.
// Save (document -> Markdown): the same constructs, with blank lines between blocks, real list numbers, escaping of
//   characters that Markdown would otherwise interpret, and images as files (or data: URIs). Underline, superscript
//   and subscript are written as inline HTML; colors, fonts and alignment have no Markdown form and are left out.
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.Formatting;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static class MarkdownConverter
{
	[Flags]
	enum Style { None = 0, Bold = 1, Italic = 2, Strike = 4, Code = 8, Underline = 16, Superscript = 32, Subscript = 64 }

	// Formatting kept as HTML: colors as COLORREF (0x00BBGGRR), font name, size in points; null = default.
	readonly record struct Css(int? Fore, int? Back, string? Font, float? Size)
	{
		public bool IsEmpty => Fore == null && Back == null && Font == null && Size == null;
		public Css With(Css more) => new(more.Fore ?? Fore, more.Back ?? Back, more.Font ?? Font, more.Size ?? Size);
		public override string ToString() => string.Join("; ", new[] {
			Fore is int f ? "color: " + Hex(f) : null,
			Back is int b ? "background-color: " + Hex(b) : null,
			Font is string n ? "font-family: " + (n.Contains(' ') ? "'" + n + "'" : n) : null,
			Size is float z ? "font-size: " + z.ToString("0.#", CultureInfo.InvariantCulture) + "pt" : null
		}.Where(x => x != null));
	}

	static string Hex(int colorRef) => $"#{colorRef & 0xFF:X2}{(colorRef >> 8) & 0xFF:X2}{(colorRef >> 16) & 0xFF:X2}";

	sealed class Span
	{
		public string Text = "";
		public Style Style;
		public Css Css;
		public string? Url;
		public string? ImageSource;    // an image: Text is its alt text
		public int ImageWidth, ImageHeight;
	}

	enum Kind { Paragraph, Heading, ListItem, Code, Rule, TableRow, PageBreak }

	sealed class Block
	{
		public Kind Kind;
		public int Level;            // heading level 1-6, or list level 0..
		public bool Ordered;
		public int Start = 1;        // number of the first item of an ordered list
		public int Quote;            // > nesting
		public string Raw = "";      // inline Markdown (or the code line)
		public List<string>? Cells;  // table row
		public bool Header;          // table header row
		public int Table;            // which table a row belongs to
		public int Group;            // which code block a code line belongs to
		public int Align;            // TOM paragraph alignment: 0 left, 1 center, 2 right, 3 justify
		public int NumberStyle = tomConstants.tomListNumberAsArabic;    // ordered list: TOM numbering (Arabic, letters, Roman)
		public int Punctuation = tomConstants.tomListPeriod;            // ordered list: 1. (period) or 1) (0)
		public bool FirstOfList;     // the first item of a list
		public bool BreakAfter;      // a list item whose last line so far ends in two spaces: a line break if a line follows
		public bool Legal;           // an item of a legal list (1.1.1)
	}

	const string RuleText = "────────────────────────────────────────";    // a horizontal rule in the document
	const char ImagePlaceholder = '\uE000';
	const int QuoteIndent = 36, ListIndent = 18;                             // points
	static readonly string[] MonospaceFonts = { "Consolas", "Courier New", "Courier", "Lucida Console", "Cascadia Code", "Cascadia Mono", "monospace" };

	// ================================================================ Load: Markdown -> document

	/// Replaces the editor's content with the Markdown. baseFolder resolves relative image paths.
	public static void Load(RichEditBox editor, string markdown, string? baseFolder) {
		var references = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		var blocks = ParseBlocks(markdown, references);
		var paragraphs = new List<(Block Block, List<Span> Spans, int Start)>();
		var text = new StringBuilder();
		Block? last = null, lastTopItem = null;
		bool anyLegal = false;
		var tables = new List<List<Block>>();    // the rows of each table, in document order
		foreach (var block in blocks) {
			if (block.Kind == Kind.TableRow && !block.Header) {    // the rows after the header go with their table
				tables[^1].Add(block);
				last = block;
				continue;
			}
			// Two lists in a row that the editor would number as one (same kind, same start): an empty paragraph between.
			if (last?.Kind == Kind.ListItem && block.Kind == Kind.ListItem && block.Level == 0 && block.FirstOfList && lastTopItem != null
				&& lastTopItem.Ordered == block.Ordered && lastTopItem.NumberStyle == block.NumberStyle && lastTopItem.Punctuation == block.Punctuation
				&& lastTopItem.Legal == block.Legal    // (a legal list and an ordinary one are never numbered as one)
				&& lastTopItem.Start == block.Start && block.Ordered) {
				text.Append('\r');
				paragraphs.Add((new Block { Kind = Kind.Paragraph }, new List<Span>(), text.Length));
			}
			if (block.Kind == Kind.ListItem && block.Level == 0) lastTopItem = block;
			// Two code blocks (or tables) in a row would become one in the document: keep an empty paragraph between them.
			if (last != null && ((last.Kind == Kind.Code && block.Kind == Kind.Code && last.Group != block.Group)
				|| (last.Kind == Kind.TableRow && block.Kind == Kind.TableRow && last.Table != block.Table))) {
				text.Append('\r');
				paragraphs.Add((new Block { Kind = Kind.Paragraph }, new List<Span>(), text.Length));
			}
			last = block;
			if (block.Kind == Kind.TableRow) tables.Add(new List<Block> { block });
			List<Span> spans = block.Kind switch {
				Kind.Code => new List<Span> { new() { Text = block.Raw } },
				Kind.Rule => new List<Span> { new() { Text = RuleText } },
				Kind.PageBreak => new List<Span> { new() { Text = PageBreaks.Break.ToString() } },    // a paragraph of its own
				Kind.TableRow => new List<Span> { new() { Text = TablePlaceholder.ToString() } },    // replaced by the table below
				_ => ParseInline(block.Raw, references)
			};
			if (paragraphs.Count > 0) text.Append('\r');
			paragraphs.Add((block, spans, text.Length));
			foreach (var span in spans) text.Append(span.ImageSource != null ? ImagePlaceholder.ToString() : span.Text.Replace('\r', ' ').Replace('\n', ' '));
		}

		editor.Html = "";
		var doc = editor.TextDocument;
		if (text.Length == 0) return;
		doc.Range2(0, 0).Text = text.ToString();
		var all = doc.Range2(0, text.Length).Font;
		all.Name = editor.Font.Name;
		all.Size = editor.Font.SizeInPoints;
		all.Bold = all.Italic = all.StrikeThrough = all.Superscript = all.Subscript = tomConstants.tomFalse;
		all.Underline = 0;

		var images = new List<(int Position, string Source, string Alt, int Width, int Height)>();
		var links = new List<(int Start, int End, string Url)>();
		var tableStops = new Dictionary<int, List<float>>();
		foreach (var (block, spans, start) in paragraphs) {
			int length = spans.Sum(s => s.ImageSource != null ? 1 : s.Text.Length);
			var para = doc.Range2(start, start).Para;
			var whole = doc.Range2(start, start + length);
			switch (block.Kind) {
				case Kind.Heading:
					whole.Font.Bold = tomConstants.tomTrue;
					whole.Font.Size = HeadingSize(block.Level);
					StyleEngine.Tag(para, -(Math.Clamp(block.Level, 1, 6) + 1));    // (the paragraph's style: Heading N)
					break;
				case Kind.Code:
					StyleEngine.Tag(para, StyleEngine.IdOf("Code"));
					whole.Font.Name = BlockConstant("PreformattedFontName", "Courier New");
					whole.Font.Size = BlockConstant("PreformattedPointSize", 10f);
					break;
				case Kind.ListItem when block.Legal:
					// the first item of a top-level list starts a list at its number (LegalLists.Normalize decides whether
					// that goes on from the list before); the others go on
					LegalLists.SetLegal(para, block.Level, block.FirstOfList && block.Level == 0 ? LegalLists.Loaded + Math.Max(1, block.Start) : 0);
					anyLegal = true;
					break;
				case Kind.ListItem:
					para.ListType = block.Ordered ? block.NumberStyle | block.Punctuation : tomConstants.tomListBullet;
					para.ListLevelIndex = block.Level;
					para.ListStart = block.Ordered ? Math.Max(1, block.Start) : 1;
					para.SetIndents(ListIndent, block.Level * ListIndent, 0);
					break;
				case Kind.Rule:
					whole.Font.ForeColor = ColorRef(Color.Gray);
					break;
			}
			if (block.Align != 0) para.Alignment = block.Align;
			if (block.Quote > 0 && block.Kind != Kind.ListItem) {
				para.SetIndents(0, block.Quote * QuoteIndent, 0);
				whole.Font.ForeColor = ColorRef(Color.FromArgb(0x55, 0x55, 0x55));
			}

			int position = start;
			foreach (var span in spans) {
				int spanLength = span.ImageSource != null ? 1 : span.Text.Length;
				if (span.ImageSource != null) images.Add((position, span.ImageSource, span.Text, span.ImageWidth, span.ImageHeight));
				else if ((span.Style != Style.None || !span.Css.IsEmpty) && spanLength > 0) {
					var font = doc.Range2(position, position + spanLength).Font;
					ApplyStyle(font, span.Style);
					if (span.Css.Fore is int fore) font.ForeColor = fore;
					if (span.Css.Back is int back) font.BackColor = back;
					if (span.Css.Font is string name) font.Name = name;
					if (span.Css.Size is float size && block.Kind != Kind.Heading) font.Size = size;
				}
				if (span.Url != null && spanLength > 0) links.Add((position, position + spanLength, span.Url));
				position += spanLength;
			}
		}

		// Images (one character each, so positions do not change), then links from the end backwards: a link adds a
		// hidden field code before its text, which would move everything after it.
		foreach (var (position, source, alt, width, height) in images.OrderByDescending(i => i.Position)) {
			editor.Select(position, 1);
			if (ImageData(source, baseFolder) is { } data) {
				editor.SelectedText = "";
				InsertImage(editor, data, alt, width, height);
			} else {
				editor.SelectedText = alt != "" ? alt : source;    // not available here (e.g. on the web): a link to it
				links.Add((position, position + (alt != "" ? alt : source).Length, source));
			}
		}
		foreach (var (start, end, url) in links.OrderByDescending(l => l.Start))
			doc.Range2(start, end).URL = "\"" + url.Replace("\"", "%22") + "\"";

		// Real tables in place of the placeholders, from the last one back. Plain cells go straight into the table's RTF;
		// cells with Markdown in them are loaded into a second editor and copied in with their formatting.
		if (tables.Count > 0) {
			string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
			var positions = story.Select((ch, i) => (ch, i)).Where(x => x.ch == TablePlaceholder).Select(x => x.i).ToList();
			int width = editor.ClientSize.Width > 200 ? Math.Min(9360, (editor.ClientSize.Width - 40) * 1440 / Math.Max(96, editor.DeviceDpi)) : 9360;
			for (int t = Math.Min(tables.Count, positions.Count) - 1; t >= 0; t--) {
				var rows = tables[t];
				int columns = rows.Max(r => r.Cells!.Count);
				editor.Select(positions[t], 1);
				editor.SelectedText = "";
				string? Cell(int r, int c) => c < rows[r].Cells!.Count ? rows[r].Cells![c] : null;
				bool Plain(string? cell) => cell == null || !Regex.IsMatch(cell, @"[*_`\[\]<\\~!&]");
				int tableStart = RtfTables.Insert(editor, rows.Count, columns, header: true, width, (r, c) => Plain(Cell(r, c)) ? Cell(r, c) : null);
				var formatted = rows.Select((row, r) => Enumerable.Range(0, columns).Select(c => Plain(Cell(r, c)) ? null : Cell(r, c)).ToList()).ToList();
				RtfTables.FillCells(editor, tableStart, formatted, header: true, (scratch, cell) => Load(scratch, cell, baseFolder));
			}
		}
		// legal lists: their numbers, last (they add text, which would move the positions used above)
		if (anyLegal) {
			LegalLists.Normalize(doc);
			LegalLists.Renumber(editor);
		}
		editor.ClearUndo();
		editor.Select(0, 0);
	}

	const char TablePlaceholder = '\uE003';

	static List<Span> TableSpans(Block row, Dictionary<string, string> references) {
		var spans = new List<Span>();
		for (int i = 0; i < row.Cells!.Count; i++) {
			if (i > 0) spans.Add(new Span { Text = "\t" });
			spans.AddRange(ParseInline(row.Cells[i], references));
		}
		return spans;
	}

	// Tab stops for a table, from the widest cell of each column.
	static List<float> TabStops(RichEditBox editor, IEnumerable<Block> rows, Dictionary<string, string> references) {
		var widths = new List<float>();
		using var bold = new Font(editor.Font, FontStyle.Bold);
		foreach (var row in rows)
			for (int i = 0; i < row.Cells!.Count; i++) {
				string plain = string.Concat(ParseInline(row.Cells[i], references).Select(s => s.Text));
				float points = TextRenderer.MeasureText(plain, row.Header ? bold : editor.Font).Width * 72f / Math.Max(96, editor.DeviceDpi);
				while (widths.Count <= i) widths.Add(0);
				widths[i] = Math.Max(widths[i], points);
			}
		var stops = new List<float>();
		float x = 0;
		for (int i = 0; i < widths.Count - 1; i++) {
			x += widths[i] + 18;
			stops.Add(x);
		}
		return stops;
	}

	static void ApplyStyle(ITextFont font, Style style) {
		if (style.HasFlag(Style.Bold)) font.Bold = tomConstants.tomTrue;
		if (style.HasFlag(Style.Italic)) font.Italic = tomConstants.tomTrue;
		if (style.HasFlag(Style.Strike)) font.StrikeThrough = tomConstants.tomTrue;
		if (style.HasFlag(Style.Underline)) font.Underline = 1;    // tomSingle
		if (style.HasFlag(Style.Superscript)) font.Superscript = tomConstants.tomTrue;
		if (style.HasFlag(Style.Subscript)) font.Subscript = tomConstants.tomTrue;
		if (style.HasFlag(Style.Code)) {
			font.Name = "Consolas";
			font.BackColor = ColorRef(Color.FromArgb(0xEE, 0xEE, 0xEE));
		}
	}

	static int ColorRef(Color c) => c.R | (c.G << 8) | (c.B << 16);

	// data: URIs, and local files (absolute, relative to baseFolder, or file: URIs).
	static byte[]? ImageData(string source, string? baseFolder) {
		try {
			var data = Regex.Match(source, @"^data:image/[a-z+.-]+;base64,(.*)$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
			if (data.Success) return Convert.FromBase64String(data.Groups[1].Value.Trim());
			if (Regex.IsMatch(source, @"^[a-z][a-z0-9+.-]*://", RegexOptions.IgnoreCase) && !source.StartsWith("file:", StringComparison.OrdinalIgnoreCase)) return null;
			string path = source.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? new Uri(source).LocalPath : Uri.UnescapeDataString(source);
			if (!Path.IsPathRooted(path)) path = Path.Combine(baseFolder ?? Environment.CurrentDirectory, path);
			return File.Exists(path) ? File.ReadAllBytes(path) : null;
		} catch (Exception ex) when (ex is FormatException or IOException or UriFormatException or ArgumentException or UnauthorizedAccessException) {
			return null;
		}
	}

	// Inserts an image at the selection: at the size asked for (a missing width or height follows the image's
	// proportions), or at its own size, at most 600 pixels wide. RichEditorNET shows images at DefaultImageWidth x
	// DefaultImageHeight, so those are set for the insert.
	static void InsertImage(RichEditBox editor, byte[] data, string alt, int wantWidth = 0, int wantHeight = 0) {
		int width, height;
		using (var buffer = new MemoryStream(data))
		using (var image = Image.FromStream(buffer)) {
			double aspect = (double)image.Height / image.Width;
			// (+1 for a size asked for: RichEditorNET rounds sizes down on the way in and out, which would lose a pixel each time)
			(width, height) = (wantWidth, wantHeight) switch {
				( > 0, > 0) => (wantWidth + 1, wantHeight + 1),
				( > 0, _) => (wantWidth + 1, (int)Math.Round(wantWidth * aspect) + 1),
				(_, > 0) => ((int)Math.Round(wantHeight / aspect) + 1, wantHeight + 1),
				_ => (Math.Min(image.Width, 600), (int)Math.Round(Math.Min(image.Width, 600) * aspect))
			};
			height = Math.Max(1, height);
		}
		int oldWidth = editor.DefaultImageWidth, oldHeight = editor.DefaultImageHeight;
		editor.DefaultImageWidth = Math.Max(1, width);
		editor.DefaultImageHeight = height;
		try {
			using var stream = new MemoryStream(data);
			editor.InsertImage(stream, alt);
		} finally {
			editor.DefaultImageWidth = oldWidth;
			editor.DefaultImageHeight = oldHeight;
		}
	}

	// ---------------------------------------------------------------- block parser

	sealed class OpenList { public int MarkerIndent, ContentColumn; public bool Ordered; public int Start; public int NumberStyle, Punctuation; public bool Legal; }

	static readonly Regex AlignOpen = new(@"^\s*<(div|p|center)(\s[^>]*)?>\s*$", RegexOptions.IgnoreCase);
	static readonly Regex AlignClose = new(@"^\s*</(div|p|center)>\s*$", RegexOptions.IgnoreCase);
	static readonly Regex AlignedLine = new(@"^\s*<(p|div|h([1-6]))(\s[^>]*)?>(.*)</\1>\s*$", RegexOptions.IgnoreCase);
	static readonly Regex ListType = new(@"\s*\[type=([1AaIi]|\([1AaIi]\)|disc|circle|square|none|legal)\]\s*$");

	// The alignment of an HTML tag's align="..." or style="text-align: ...", or null.
	static int? AlignOf(string tag, string attributes) {
		if (tag.Equals("center", StringComparison.OrdinalIgnoreCase)) return 1;
		var m = Regex.Match(attributes, @"(?:align\s*=\s*[""']?|text-align\s*:\s*)(left|center|right|justify)", RegexOptions.IgnoreCase);
		return m.Success ? m.Groups[1].Value.ToLowerInvariant() switch { "center" => 1, "right" => 2, "justify" => 3, _ => 0 } : null;
	}

	// GitHub's (and M-ArkDown's) numbering when a list has no [type=]: 1. at the first level, i. at the second, a. below.
	static int DefaultNumberStyle(int level) => level switch { 0 => tomConstants.tomListNumberAsArabic, 1 => tomConstants.tomListNumberAsLCRoman, _ => tomConstants.tomListNumberAsLCLetter };

	static int NumberStyleOf(string type) => type switch {
		"a" => tomConstants.tomListNumberAsLCLetter, "A" => tomConstants.tomListNumberAsUCLetter,
		"i" => tomConstants.tomListNumberAsLCRoman, "I" => tomConstants.tomListNumberAsUCRoman, _ => tomConstants.tomListNumberAsArabic
	};

	static string TypeOf(int numberStyle) =>
		numberStyle == tomConstants.tomListNumberAsLCLetter ? "a" : numberStyle == tomConstants.tomListNumberAsUCLetter ? "A"
		: numberStyle == tomConstants.tomListNumberAsLCRoman ? "i" : numberStyle == tomConstants.tomListNumberAsUCRoman ? "I" : "1";

	static readonly Regex ReferenceDefinition = new(@"^ {0,3}\[([^\]]+)\]:\s*<?([^\s>]+)>?(?:\s+(?:""[^""]*""|'[^']*'|\([^)]*\)))?\s*$");
	static readonly Regex Fence = new(@"^( {0,3})(`{3,}|~{3,})\s*([^`]*)$");
	static readonly Regex AtxHeading = new(@"^ {0,3}(#{1,6})(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$");
	static readonly Regex SetextUnderline = new(@"^ {0,3}(=+|-+)[ \t]*$");
	static readonly Regex Rule = new(@"^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$");
	static readonly Regex ListMarker = new(@"^( *)([-+*]|(\d{1,9})([.)]))( +|$)(.*)$");
	static readonly Regex TableDelimiter = new(@"^ {0,3}\|?[ \t]*:?-+:?[ \t]*(\|[ \t]*:?-+:?[ \t]*)*\|?[ \t]*$");

	static List<Block> ParseBlocks(string markdown, Dictionary<string, string> references) {
		var lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Select(ExpandTabs).ToList();
		// Reference definitions ([label]: url) are collected and removed (but not inside fenced code).
		bool inFence = false;
		for (int i = 0; i < lines.Count; i++) {
			if (Fence.IsMatch(StripQuote(lines[i], out _))) inFence = !inFence;
			if (!inFence && ReferenceDefinition.Match(lines[i]) is { Success: true } m) {
				references.TryAdd(NormalizeLabel(m.Groups[1].Value), m.Groups[2].Value);
				lines.RemoveAt(i--);
			}
		}

		var blocks = new List<Block>();
		var paragraph = new List<string>();
		int paragraphQuote = 0, table = 0, codeBlocks = 0;
		var lists = new List<OpenList>();
		Block? item = null;            // the list item that continuation lines belong to
		bool afterBlank = false;
		var aligns = new List<int> { 0 };    // <div align> nesting

		void Add(Block block) {
			if (block.Align == 0) block.Align = aligns[^1];
			blocks.Add(block);
		}

		void FlushParagraph() {
			if (paragraph.Count == 0) return;
			Add(new Block { Kind = Kind.Paragraph, Raw = JoinLines(paragraph), Quote = paragraphQuote });
			paragraph.Clear();
		}

		void CloseLists() {
			lists.Clear();
			item = null;
		}

		for (int i = 0; i < lines.Count; i++) {
			string content = StripQuote(lines[i], out int quote);
			bool blank = content.Trim() == "";

			// fenced code block
			var fence = Fence.Match(content);
			if (fence.Success && !(fence.Groups[2].Value[0] == '`' && fence.Groups[3].Value.Contains('`'))) {
				FlushParagraph();
				CloseLists();
				string marker = fence.Groups[2].Value;
				int group = ++codeBlocks;
				int fenceIndent = fence.Groups[1].Value.Length;
				for (i++; i < lines.Count; i++) {
					string line = StripQuote(lines[i], out _);
					var close = Regex.Match(line, @"^ {0,3}(" + Regex.Escape(marker[0].ToString()) + @"{" + marker.Length + @",})[ \t]*$");
					if (close.Success) break;
					Add(new Block { Kind = Kind.Code, Raw = RemoveIndent(line, fenceIndent), Quote = quote, Group = group });
				}
				afterBlank = false;
				continue;
			}

			if (blank) {
				FlushParagraph();
				afterBlank = true;
				continue;
			}

			// a page break (HTML on a line of its own, or \newpage): ends paragraphs and lists, like other HTML blocks
			if (quote == 0 && PageBreaks.MarkdownLine.IsMatch(content)) {
				FlushParagraph();
				CloseLists();
				blocks.Add(new Block { Kind = Kind.PageBreak });
				afterBlank = false;
				continue;
			}

			// alignment as HTML: <div align="center"> ... </div>, or <p align="right">text</p> on one line
			if (AlignedLine.Match(content) is { Success: true } alignedLine) {
				FlushParagraph();
				CloseLists();
				int align = AlignOf(alignedLine.Groups[1].Value, alignedLine.Groups[3].Value) ?? aligns[^1];
				Add(alignedLine.Groups[2].Success
					? new Block { Kind = Kind.Heading, Level = int.Parse(alignedLine.Groups[2].Value), Raw = alignedLine.Groups[4].Value.Trim(), Quote = quote, Align = align }
					: new Block { Kind = Kind.Paragraph, Raw = alignedLine.Groups[4].Value.Trim(), Quote = quote, Align = align });
				afterBlank = false;
				continue;
			}
			if (AlignOpen.Match(content) is { Success: true } alignOpen) {
				FlushParagraph();
				CloseLists();
				aligns.Add(AlignOf(alignOpen.Groups[1].Value, alignOpen.Groups[2].Value) ?? aligns[^1]);
				continue;
			}
			// an HTML comment on its own line (e.g. <!-- --> between two lists): ends the list, shows nothing
			if (Regex.IsMatch(content, @"^\s*<!--.*-->\s*$")) {
				FlushParagraph();
				CloseLists();
				continue;
			}
			if (AlignClose.IsMatch(content)) {
				FlushParagraph();
				CloseLists();
				if (aligns.Count > 1) aligns.RemoveAt(aligns.Count - 1);
				continue;
			}

			int indent = content.Length - content.TrimStart(' ').Length;
			var marker2 = ListMarker.Match(content);
			bool isRule = Rule.IsMatch(content);
			bool isListItem = marker2.Success && !isRule
				&& !(paragraph.Count > 0 && lists.Count == 0 && marker2.Groups[3].Success && marker2.Groups[3].Value != "1");    // only "1." may interrupt a paragraph

			// list continuation lines
			if (lists.Count > 0 && item != null && !isListItem) {
				bool startsBlock = AtxHeading.IsMatch(content) || isRule || Fence.IsMatch(content);
				if (indent >= lists[^1].ContentColumn || (!afterBlank && !startsBlock)) {
					item.Raw = JoinLines(new List<string> { item.Raw + (item.BreakAfter ? "  " : ""), content.TrimStart() });
					item.BreakAfter = content.EndsWith("  ");
					afterBlank = false;
					continue;
				}
				CloseLists();
			}

			// setext heading: a paragraph followed by === or ---
			if (paragraph.Count > 0 && quote == paragraphQuote && SetextUnderline.Match(content) is { Success: true } underline) {
				Add(new Block { Kind = Kind.Heading, Level = underline.Groups[1].Value[0] == '=' ? 1 : 2, Raw = JoinLines(paragraph), Quote = quote });
				paragraph.Clear();
				afterBlank = false;
				continue;
			}

			if (AtxHeading.Match(content) is { Success: true } heading) {
				FlushParagraph();
				CloseLists();
				Add(new Block { Kind = Kind.Heading, Level = heading.Groups[1].Length, Raw = heading.Groups[2].Value.Trim(), Quote = quote });
				afterBlank = false;
				continue;
			}

			if (isRule) {
				FlushParagraph();
				CloseLists();
				Add(new Block { Kind = Kind.Rule, Quote = quote });
				afterBlank = false;
				continue;
			}

			// GitHub table: a row of cells, then a delimiter row
			if (paragraph.Count == 0 && content.Contains('|') && i + 1 < lines.Count && TableDelimiter.IsMatch(StripQuote(lines[i + 1], out _))
				&& SplitCells(StripQuote(lines[i + 1], out _)).Count == SplitCells(content).Count) {
				CloseLists();
				table++;
				Add(new Block { Kind = Kind.TableRow, Cells = SplitCells(content), Header = true, Table = table, Quote = quote });
				for (i += 2; i < lines.Count; i++) {
					string row = StripQuote(lines[i], out _);
					if (row.Trim() == "" || !row.Contains('|')) { i--; break; }
					Add(new Block { Kind = Kind.TableRow, Cells = SplitCells(row), Table = table, Quote = quote });
				}
				afterBlank = false;
				continue;
			}

			if (isListItem) {
				FlushParagraph();
				bool ordered = marker2.Groups[3].Success;
				int number = ordered ? int.Parse(marker2.Groups[3].Value, CultureInfo.InvariantCulture) : 1;
				int spaces = marker2.Groups[5].Value.Length;
				int contentColumn = indent + marker2.Groups[2].Length + (spaces is >= 1 and <= 4 ? spaces : 1);
				int punctuation = marker2.Groups[4].Value == ")" ? 0 : tomConstants.tomListPeriod;    // TOM: 0 is "1)"
				string raw = marker2.Groups[6].Value.Trim();
				var type = ListType.Match(raw);
				if (type.Success) raw = raw[..type.Index].TrimEnd();
				bool newList = true;
				while (lists.Count > 0 && indent < lists[^1].MarkerIndent) lists.RemoveAt(lists.Count - 1);
				if (lists.Count > 0 && indent >= lists[^1].ContentColumn) {
					lists.Add(new OpenList { MarkerIndent = indent, ContentColumn = contentColumn, Ordered = ordered, Start = number });    // nested
				} else if (lists.Count > 0) {
					// bullets after numbers (or the reverse), or "1)" after "1.", start a new list (CommonMark)
					if (lists[^1].Ordered != ordered || (ordered && (lists[^1].Punctuation == tomConstants.tomListParentheses ? 0 : lists[^1].Punctuation) != punctuation))
						lists[^1] = new OpenList { MarkerIndent = indent, ContentColumn = contentColumn, Ordered = ordered, Start = number };
					else { lists[^1].ContentColumn = contentColumn; newList = false; }
				} else {
					lists.Add(new OpenList { MarkerIndent = indent, ContentColumn = contentColumn, Ordered = ordered, Start = number });
				}
				var open = lists[^1];
				if (newList) {
					// M-ArkDown: [type=a] on the first item sets the numbering of the whole list
					string typeValue = type.Success ? type.Groups[1].Value : "";
					bool parens = typeValue.StartsWith('(');    // [type=(a)]: "(a)" numbering (written "1)")
					open.NumberStyle = type.Success && ordered ? NumberStyleOf(parens ? typeValue[1..^1] : typeValue) : DefaultNumberStyle(lists.Count - 1);
					open.Punctuation = parens && ordered ? tomConstants.tomListParentheses : punctuation;
					// a legal list: [type=legal], or a numbered list nested in one
					open.Legal = ordered && (type.Success ? type.Groups[1].Value == "legal" : lists.Count > 1 && lists[^2].Legal);
				}
				item = new Block { Kind = Kind.ListItem, Level = lists.Count - 1, Ordered = ordered, Start = open.Start, Raw = raw, Quote = quote, FirstOfList = newList,
					NumberStyle = open.NumberStyle, Punctuation = open.Punctuation, Legal = open.Legal, BreakAfter = marker2.Groups[6].Value.EndsWith("  ") };
				Add(item);
				afterBlank = false;
				continue;
			}

			// indented code block
			if (indent >= 4 && paragraph.Count == 0) {
				int group = ++codeBlocks;
				for (; i < lines.Count; i++) {
					string line = StripQuote(lines[i], out _);
					if (line.Trim() != "" && line.Length - line.TrimStart(' ').Length < 4) { i--; break; }
					Add(new Block { Kind = Kind.Code, Raw = line.Length >= 4 ? line[4..] : "", Quote = quote, Group = group });
				}
				while (blocks.Count > 0 && blocks[^1].Kind == Kind.Code && blocks[^1].Raw.Trim() == "") blocks.RemoveAt(blocks.Count - 1);    // trailing blank lines
				afterBlank = false;
				continue;
			}

			if (paragraph.Count == 0) paragraphQuote = quote;
			paragraph.Add(content);
			afterBlank = false;
		}
		FlushParagraph();
		return blocks;
	}

	// The lines of a paragraph: a line ending in two spaces or a backslash is a line break (\v), others become spaces.
	static string JoinLines(List<string> lines) {
		var sb = new StringBuilder();
		for (int i = 0; i < lines.Count; i++) {
			string line = lines[i].TrimStart();
			bool last = i == lines.Count - 1;
			if (!last && line.EndsWith("  ")) sb.Append(line.TrimEnd()).Append('\v');
			else if (!last && line.EndsWith('\\') && !line.EndsWith("\\\\")) sb.Append(line[..^1]).Append('\v');
			else sb.Append(last ? line.TrimEnd() : line.TrimEnd() + " ");
		}
		return sb.ToString();
	}

	static string StripQuote(string line, out int quote) {
		quote = 0;
		while (Regex.Match(line, @"^ {0,3}> ?") is { Success: true } m) {
			quote++;
			line = line[m.Length..];
		}
		return line;
	}

	static string ExpandTabs(string line) {
		if (!line.Contains('\t')) return line;
		var sb = new StringBuilder();
		int i = 0;
		for (; i < line.Length && (line[i] == ' ' || line[i] == '\t'); i++)    // leading tabs only: they matter for indentation
			if (line[i] == '\t') sb.Append(' ', 4 - sb.Length % 4); else sb.Append(' ');
		return sb + line[i..];
	}

	static string RemoveIndent(string line, int count) {
		int i = 0;
		while (i < count && i < line.Length && line[i] == ' ') i++;
		return line[i..];
	}

	static List<string> SplitCells(string row) {
		row = row.Trim();
		if (row.StartsWith('|')) row = row[1..];
		if (row.EndsWith('|') && !row.EndsWith("\\|")) row = row[..^1];
		var cells = new List<string>();
		var cell = new StringBuilder();
		for (int i = 0; i < row.Length; i++) {
			if (row[i] == '\\' && i + 1 < row.Length && row[i + 1] == '|') { cell.Append('|'); i++; }
			else if (row[i] == '|') { cells.Add(cell.ToString().Trim()); cell.Clear(); }
			else cell.Append(row[i]);
		}
		cells.Add(cell.ToString().Trim());
		return cells;
	}

	/// A GitHub table's lines (header, delimiter row, rows) with their columns lined up: each cell padded to its column's
	/// width, the delimiter row's dashes as wide (keeping its ":" alignment), right-aligned columns padded on the left;
	/// a table with a cell over 40 columns wide, or rows over 120, compactly (no padding, "| --- |").
	/// Cells keep their text as written (escaped "\|" included); a row with fewer cells is filled up.
	public static List<string> AlignTable(IReadOnlyList<string> lines) {
		var rows = lines.Select(RawCells).ToList();
		int delimiter = lines.Count > 1 && TableDelimiter.IsMatch(lines[1]) ? 1 : -1;
		int count = rows.Max(r => r.Count);
		foreach (var row in rows) while (row.Count < count) row.Add("");
		var widths = Enumerable.Range(0, count).Select(c => Math.Max(3, rows.Where((row, i) => i != delimiter).Select(row => DisplayWidth(row[c])).DefaultIfEmpty(0).Max())).ToList();
		var right = Enumerable.Range(0, count).Select(c => delimiter >= 0 && rows[delimiter][c].EndsWith(':') && !rows[delimiter][c].StartsWith(':')).ToList();
		var result = new List<string>();
		// A table with long cells (paragraphs, addresses) can't be lined up readably, since a row is one line: it is
		// written compactly, without padding, and with a short delimiter row.
		if (widths.Max() > 40 || widths.Sum() + 3 * count + 1 > 120) {
			for (int r = 0; r < rows.Count; r++) {
				var cells = r == delimiter
					? rows[r].Select(cell => (cell.StartsWith(':') ? ":" : "") + "---" + (cell.EndsWith(':') && cell.Length > 1 ? ":" : ""))
					: rows[r];
				result.Add("| " + string.Join(" | ", cells) + " |");
			}
			return result;
		}
		for (int r = 0; r < rows.Count; r++) {
			var cells = new List<string>();
			for (int c = 0; c < count; c++) {
				string cell = rows[r][c];
				if (r == delimiter) {
					bool left = cell.StartsWith(':'), end = cell.EndsWith(':') && cell.Length > 1;
					cells.Add((left ? ":" : "") + new string('-', widths[c] - (left ? 1 : 0) - (end ? 1 : 0)) + (end ? ":" : ""));
				} else {
					string padding = new string(' ', Math.Max(0, widths[c] - DisplayWidth(cell)));
					cells.Add(right[c] ? padding + cell : cell + padding);
				}
			}
			result.Add("| " + string.Join(" | ", cells) + " |");
		}
		return result;
	}

	/// Whether a line is a GitHub table's delimiter row (| --- | :-: |).
	public static bool IsTableDelimiter(string line) => TableDelimiter.IsMatch(line);

	// A table row's cells as written: split at "|" that is not escaped ("\|" stays as it is), trimmed.
	static List<string> RawCells(string row) {
		row = row.Trim();
		if (row.StartsWith('|')) row = row[1..];
		if (row.EndsWith('|') && !row.EndsWith("\\|")) row = row[..^1];
		var cells = new List<string>();
		int start = 0;
		for (int i = 0; i < row.Length; i++) {
			if (row[i] == '\\') { i++; continue; }
			if (row[i] == '|') { cells.Add(row[start..i].Trim()); start = i + 1; }
		}
		cells.Add(row[start..].Trim());
		return cells;
	}

	// How many columns text takes in a monospaced font: East Asian wide characters (and most emoji) take two.
	public static int DisplayWidth(string text) {
		int width = 0;
		var elements = System.Globalization.StringInfo.GetTextElementEnumerator(text);
		while (elements.MoveNext()) {
			int c = char.ConvertToUtf32((string)elements.Current, 0);
			bool wide = c is >= 0x1100 and <= 0x115F or >= 0x2E80 and <= 0xA4CF or >= 0xAC00 and <= 0xD7A3 or >= 0xF900 and <= 0xFAFF
				or >= 0xFE30 and <= 0xFE4F or >= 0xFF00 and <= 0xFF60 or >= 0xFFE0 and <= 0xFFE6 or >= 0x1F300 and <= 0x1FAFF or >= 0x20000 and <= 0x3FFFD;
			width += wide ? 2 : 1;
		}
		return width;
	}

	static string NormalizeLabel(string label) => Regex.Replace(label.Trim(), @"\s+", " ");

	// ---------------------------------------------------------------- inline parser

	sealed class Node
	{
		public string Text = "";
		public Style Style;
		public Css Css;
		public int ImageWidth, ImageHeight;
		public string? Url, ImageSource;
		public char Delimiter;           // '*', '_' or '~' for an emphasis delimiter run
		public int Count, OriginalCount;
		public bool CanOpen, CanClose;
	}

	const string Punctuation = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~";

	static List<Span> ParseInline(string text, Dictionary<string, string> references) {
		var nodes = ParseNodes(text, references, Style.None, default);
		ProcessEmphasis(nodes);
		var spans = new List<Span>();
		foreach (var node in nodes) {
			string nodeText = node.Delimiter != '\0' ? new string(node.Delimiter, node.Count) : node.Text;
			if (nodeText.Length == 0 && node.ImageSource == null) continue;
			if (node.ImageSource == null && spans.Count > 0 && spans[^1].ImageSource == null && spans[^1].Style == node.Style && spans[^1].Url == node.Url && spans[^1].Css == node.Css)
				spans[^1].Text += nodeText;
			else
				spans.Add(new Span { Text = nodeText, Style = node.Style, Css = node.Css, Url = node.Url, ImageSource = node.ImageSource,
					ImageWidth = node.ImageWidth, ImageHeight = node.ImageHeight });
		}
		return spans;
	}

	static List<Node> ParseNodes(string s, Dictionary<string, string> references, Style htmlStyle, Css css) {
		var nodes = new List<Node>();
		var text = new StringBuilder();
		var cssStack = new List<(string Tag, Css Previous)>();
		void FlushText() {
			if (text.Length > 0) nodes.Add(new Node { Text = text.ToString(), Style = htmlStyle, Css = css });
			text.Clear();
		}

		for (int i = 0; i < s.Length; i++) {
			char c = s[i];
			if (c == '\\' && i + 1 < s.Length && Punctuation.Contains(s[i + 1])) {
				text.Append(s[++i]);
				continue;
			}
			if (c == '`') {
				int run = CountRun(s, i, '`');
				int close = FindBacktickRun(s, i + run, run);
				if (close >= 0) {
					FlushText();
					string code = s[(i + run)..close].Replace('\v', ' ');
					if (code.Length >= 2 && code[0] == ' ' && code[^1] == ' ' && code.Trim() != "") code = code[1..^1];
					nodes.Add(new Node { Text = code, Style = htmlStyle | Style.Code, Css = css });
					i = close + run - 1;
				} else {
					text.Append('`', run);
					i += run - 1;
				}
				continue;
			}
			if ((c == '[' || (c == '!' && i + 1 < s.Length && s[i + 1] == '[')) && ParseLink(s, c == '!' ? i + 1 : i, references) is { } link) {
				FlushText();
				if (c == '!') {
					string alt = string.Concat(ParseInline(link.Label, references).Select(sp => sp.Text));
					var image = new Node { Text = alt, ImageSource = link.Url, Style = htmlStyle, Css = css };
					// M-ArkDown: ![alt](src)(width=200) or (height=100)
					var size = Regex.Match(s[link.End..], @"^\(\s*((?:width|height)\s*=\s*\d+(?:\s*[ ,;]\s*(?:width|height)\s*=\s*\d+)?)\s*\)", RegexOptions.IgnoreCase);
					if (size.Success) {
						foreach (Match part in Regex.Matches(size.Groups[1].Value, @"(width|height)\s*=\s*(\d+)", RegexOptions.IgnoreCase))
							if (part.Groups[1].Value.Equals("width", StringComparison.OrdinalIgnoreCase)) image.ImageWidth = int.Parse(part.Groups[2].Value);
							else image.ImageHeight = int.Parse(part.Groups[2].Value);
						link = (link.Label, link.Url, link.End + size.Length);
					}
					nodes.Add(image);
				} else {
					var inner = ParseNodes(link.Label, references, htmlStyle, css);
					ProcessEmphasis(inner);
					foreach (var node in inner) {
						if (node.Delimiter != '\0') { node.Text = new string(node.Delimiter, node.Count); node.Delimiter = '\0'; }
						node.Url ??= link.Url;
					}
					nodes.AddRange(inner);
				}
				i = link.End - 1;
				continue;
			}
			if (c == '<') {
				var auto = Regex.Match(s[i..], @"^<([a-zA-Z][a-zA-Z0-9+.-]{1,31}:[^\s<>]*|[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+)>");
				if (auto.Success) {
					FlushText();
					string target = auto.Groups[1].Value;
					nodes.Add(new Node { Text = target, Url = target.Contains(':') ? target : "mailto:" + target, Style = htmlStyle, Css = css });
					i += auto.Length - 1;
					continue;
				}
				var tag = Regex.Match(s[i..], @"^<(/?)(u|ins|sup|sub|br|b|strong|i|em|s|del|strike)\s*/?>", RegexOptions.IgnoreCase);
				if (tag.Success) {
					FlushText();
					string name = tag.Groups[2].Value.ToLowerInvariant();
					if (name == "br") {
						text.Append('\v');
					} else {
						var flag = name switch {
							"sup" => Style.Superscript, "sub" => Style.Subscript, "b" or "strong" => Style.Bold, "i" or "em" => Style.Italic,
							"s" or "del" or "strike" => Style.Strike, _ => Style.Underline
						};
						htmlStyle = tag.Groups[1].Value == "/" ? htmlStyle & ~flag : htmlStyle | flag;
					}
					i += tag.Length - 1;
					continue;
				}
				// colors and fonts: <span style="...">, <font ...>, M-ArkDown <c=#F00>, <mark>; and their closing tags
				var cssTag = Regex.Match(s[i..], @"^<(span|font|mark|c)(?:\s*=\s*(#?[0-9a-fA-F]{3,6})|(\s[^>]*))?>", RegexOptions.IgnoreCase);
				if (cssTag.Success && CssOf(cssTag.Groups[1].Value.ToLowerInvariant(), cssTag.Groups[2].Value, cssTag.Groups[3].Value) is Css added) {
					FlushText();
					cssStack.Add((cssTag.Groups[1].Value.ToLowerInvariant(), css));
					css = css.With(added);
					i += cssTag.Length - 1;
					continue;
				}
				var cssClose = Regex.Match(s[i..], @"^</(span|font|mark|c)>", RegexOptions.IgnoreCase);
				if (cssClose.Success && cssStack.FindLastIndex(t => t.Tag == cssClose.Groups[1].Value.ToLowerInvariant()) is var at and >= 0) {
					FlushText();
					css = cssStack[at].Previous;
					cssStack.RemoveRange(at, cssStack.Count - at);
					i += cssClose.Length - 1;
					continue;
				}
			}
			if (c == '*' || c == '_' || (c == '~' && i + 1 < s.Length && s[i + 1] == '~')) {
				int run = CountRun(s, i, c);
				char before = i > 0 ? s[i - 1] : ' ', after = i + run < s.Length ? s[i + run] : ' ';
				bool afterSpace = char.IsWhiteSpace(after), beforeSpace = char.IsWhiteSpace(before);
				bool afterPunct = Punctuation.Contains(after) || char.IsPunctuation(after), beforePunct = Punctuation.Contains(before) || char.IsPunctuation(before);
				bool left = !afterSpace && (!afterPunct || beforeSpace || beforePunct);
				bool right = !beforeSpace && (!beforePunct || afterSpace || afterPunct);
				FlushText();
				nodes.Add(new Node {
					Delimiter = c, Count = run, OriginalCount = run, Style = htmlStyle, Css = css,
					CanOpen = c == '_' ? left && (!right || beforePunct) : left,
					CanClose = c == '_' ? right && (!left || afterPunct) : right
				});
				i += run - 1;
				continue;
			}
			text.Append(c);
		}
		FlushText();
		return nodes;
	}

	// The formatting of a <span style>, <font>, <mark> or <c=#...> tag, or null if it has none we use.
	static Css? CssOf(string tag, string cValue, string attributes) {
		if (tag == "mark") return new Css(null, ColorRef(Color.Yellow), null, null);
		if (tag == "c") return ColorOf(cValue) is int c ? new Css(c, null, null, null) : null;
		if (tag == "font") {
			int? color = Regex.Match(attributes, @"color\s*=\s*[""']?([^""'\s>]+)", RegexOptions.IgnoreCase) is { Success: true } m1 ? ColorOf(m1.Groups[1].Value) : null;
			string? face = Regex.Match(attributes, @"face\s*=\s*[""']([^""']+)", RegexOptions.IgnoreCase) is { Success: true } m2 ? m2.Groups[1].Value.Split(',')[0].Trim() : null;
			float? size = Regex.Match(attributes, @"size\s*=\s*[""']?([1-7])", RegexOptions.IgnoreCase) is { Success: true } m3
				? new[] { 0f, 8f, 10f, 12f, 14f, 18f, 24f, 36f }[int.Parse(m3.Groups[1].Value)] : null;
			return new Css(color, null, face, size);
		}
		var style = Regex.Match(attributes, @"style\s*=\s*(""([^""]*)""|'([^']*)')", RegexOptions.IgnoreCase);
		if (!style.Success) return new Css();
		int? fore = null, back = null;
		string? font = null;
		float? points = null;
		foreach (var declaration in (style.Groups[2].Success ? style.Groups[2].Value : style.Groups[3].Value).Split(';')) {
			int colon = declaration.IndexOf(':');
			if (colon < 0) continue;
			string property = declaration[..colon].Trim().ToLowerInvariant(), value = declaration[(colon + 1)..].Trim();
			switch (property) {
				case "color": fore = ColorOf(value); break;
				case "background-color" or "background": back = ColorOf(value); break;
				case "font-family": font = value.Split(',')[0].Trim().Trim('\'', '"'); break;
				case "font-size":
					var number = Regex.Match(value, @"^([\d.]+)\s*(pt|px)?$", RegexOptions.IgnoreCase);
					if (number.Success && float.TryParse(number.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out float n))
						points = number.Groups[2].Value.Equals("px", StringComparison.OrdinalIgnoreCase) ? n * 0.75f : n;
					break;
			}
		}
		return new Css(fore, back, font, points);
	}

	// #RGB, #RRGGBB, rgb(r, g, b) or a color name, as a COLORREF.
	static int? ColorOf(string value) {
		value = value.Trim();
		try {
			var rgb = Regex.Match(value, @"^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$", RegexOptions.IgnoreCase);
			if (rgb.Success) return ColorRef(Color.FromArgb(int.Parse(rgb.Groups[1].Value), int.Parse(rgb.Groups[2].Value), int.Parse(rgb.Groups[3].Value)));
			if (Regex.IsMatch(value, "^[0-9a-fA-F]{3}$|^[0-9a-fA-F]{6}$")) value = "#" + value;
			if (Regex.IsMatch(value, "^#[0-9a-fA-F]{3}$")) value = "#" + string.Concat(value[1..].Select(ch => $"{ch}{ch}"));
			return ColorRef(ColorTranslator.FromHtml(value));
		} catch (Exception ex) when (ex is ArgumentException or FormatException or OverflowException) {
			return null;
		}
	}

	// CommonMark's "process emphasis", in short: each closing run is matched with the nearest opening run of the same
	// character; the text between gets bold (two characters) or italic (one), or strikethrough for ~~.
	static void ProcessEmphasis(List<Node> nodes) {
		for (int c = 0; c < nodes.Count; c++) {
			var closer = nodes[c];
			if (closer.Delimiter == '\0' || !closer.CanClose || closer.Count == 0) continue;
			for (int o = c - 1; o >= 0; o--) {
				var opener = nodes[o];
				if (opener.Delimiter != closer.Delimiter || !opener.CanOpen || opener.Count == 0) continue;
				if (closer.Delimiter != '~' && (opener.CanClose || closer.CanOpen)
					&& (opener.OriginalCount + closer.OriginalCount) % 3 == 0 && !(opener.OriginalCount % 3 == 0 && closer.OriginalCount % 3 == 0)) continue;
				int use;
				Style style;
				if (closer.Delimiter == '~') {
					if (opener.Count < 2 || closer.Count < 2) continue;
					use = 2;
					style = Style.Strike;
				} else {
					use = opener.Count >= 2 && closer.Count >= 2 ? 2 : 1;
					style = use == 2 ? Style.Bold : Style.Italic;
				}
				for (int k = o + 1; k < c; k++) {
					nodes[k].Style |= style;
					if (nodes[k].Delimiter != '\0') nodes[k].CanOpen = nodes[k].CanClose = false;    // runs in between can no longer match
				}
				opener.Count -= use;
				closer.Count -= use;
				if (closer.Count > 0) c--;    // the rest of this run may close another one
				break;
			}
		}
	}

	static int CountRun(string s, int i, char c) {
		int n = 0;
		while (i + n < s.Length && s[i + n] == c) n++;
		return n;
	}

	static int FindBacktickRun(string s, int from, int length) {
		for (int i = from; i < s.Length; i++) {
			if (s[i] != '`') continue;
			int run = CountRun(s, i, '`');
			if (run == length) return i;
			i += run - 1;
		}
		return -1;
	}

	// [label](url "title"), [label][ref], [label][] or [label], starting at the '['.
	static (string Label, string Url, int End)? ParseLink(string s, int open, Dictionary<string, string> references) {
		int depth = 0, close = -1;
		for (int i = open; i < s.Length; i++) {
			if (s[i] == '\\') { i++; continue; }
			if (s[i] == '`') { int run = CountRun(s, i, '`'); int end = FindBacktickRun(s, i + run, run); if (end >= 0) { i = end + run - 1; continue; } i += run - 1; continue; }
			if (s[i] == '[') depth++;
			else if (s[i] == ']' && --depth == 0) { close = i; break; }
		}
		if (close < 0) return null;
		string label = s[(open + 1)..close];
		int pos = close + 1;
		if (pos < s.Length && s[pos] == '(') {
			int p = pos + 1;
			while (p < s.Length && char.IsWhiteSpace(s[p])) p++;
			string url;
			if (p < s.Length && s[p] == '<') {
				int end = s.IndexOf('>', p);
				if (end < 0) return null;
				url = s[(p + 1)..end];
				p = end + 1;
			} else {
				int start = p, parens = 0;
				for (; p < s.Length; p++) {
					if (s[p] == '\\' && p + 1 < s.Length) { p++; continue; }
					if (char.IsWhiteSpace(s[p])) break;
					if (s[p] == '(') parens++;
					else if (s[p] == ')') { if (parens == 0) break; parens--; }
				}
				url = Regex.Replace(s[start..p], @"\\([!-/:-@\[-`{-~])", "$1");
			}
			while (p < s.Length && char.IsWhiteSpace(s[p])) p++;
			if (p < s.Length && (s[p] == '"' || s[p] == '\'' || s[p] == '(')) {    // title: ignored
				char end = s[p] == '(' ? ')' : s[p];
				int titleEnd = s.IndexOf(end, p + 1);
				if (titleEnd < 0) return null;
				p = titleEnd + 1;
				while (p < s.Length && char.IsWhiteSpace(s[p])) p++;
			}
			if (p >= s.Length || s[p] != ')') return null;
			return (label, url, p + 1);
		}
		// reference links
		if (pos < s.Length && s[pos] == '[') {
			int end = s.IndexOf(']', pos + 1);
			if (end >= 0) {
				string reference = s[(pos + 1)..end];
				if (references.TryGetValue(NormalizeLabel(reference == "" ? label : reference), out var target)) return (label, target, end + 1);
			}
		}
		if (references.TryGetValue(NormalizeLabel(label), out var shortcut)) return (label, shortcut, close + 1);
		return null;
	}

	// ================================================================ Save: document -> Markdown

	/// The document as Markdown. saveImage(number, bytes, extension) stores the n-th image (1-based) and returns the
	/// path to write in the Markdown; without it, images are written as data: URIs.
	public static string ToMarkdown(RichEditBox editor, Func<int, byte[], string, string>? saveImage = null) {
		var doc = editor.TextDocument;
		// The images' data, alt texts and sizes come from the library's HTML (in document order); TOM does not give them.
		var images = Regex.Matches(editor.Html, "<img src=\"data:image/([a-z+.-]+);base64,([^\"]*)\"(?: alt=\"([^\"]*)\")?(?: width=\"(\\d+)\")?(?: height=\"(\\d+)\")?", RegexOptions.IgnoreCase)
			.Select(m => (Extension: m.Groups[1].Value.ToLowerInvariant() switch { "jpeg" => "jpg", var e => e }, Data: m.Groups[2].Value,
				Alt: System.Net.WebUtility.HtmlDecode(m.Groups[3].Value), Width: m.Groups[4].Success ? int.Parse(m.Groups[4].Value) : 0)).ToList();
		int imageNumber = 0;
		(string Source, string Alt, string Size) ImageSource() {
			if (imageNumber >= images.Count) return ("", "", "");
			var (extension, base64, alt, width) = images[imageNumber++];
			byte[] data = Convert.FromBase64String(base64);
			string size = "";
			try {
				using var buffer = new MemoryStream(data);
				using var image = Image.FromStream(buffer);
				if (width > 0 && Math.Abs(width - Math.Min(image.Width, 600)) > 2) size = $"(width={width})";    // M-ArkDown
			} catch (ArgumentException) { }
			return (saveImage != null ? saveImage(imageNumber, data, extension) : $"data:image/{extension};base64,{base64}", alt, size);
		}
		var context = new InlineContext(editor.Font.Name, editor.Font.SizeInPoints, !editor.EnableHtmlFontSizing, ImageSource);

		var paragraphs = ReadParagraphs(doc);
		var output = new List<string>();
		string? previous = null;    // kind of the previous block, for the blank lines between blocks
		int previousQuote = 0;
		var counters = new List<int>();
		var kinds = new List<(int, int)>();    // per level: (numbering style, punctuation) of the open list
		var runStarts = new List<int>();       // per level: the start value its items share (RichEdit numbers on through them)
		bool gap = false;                      // empty paragraphs since the last list item: the list was interrupted
		(int, int)? lastTopKind = null;        // kind of the last top-level list
		var columns = new List<int>();

		// The inline Markdown of paragraph p, joined with the lines that follow a line break in it.
		string WithLineBreaks(ref int p, bool noBold) {
			string s = Inline(doc, paragraphs[p].Start, paragraphs[p].End, noBold, false, context, paragraphs[p].Quote);
			while (paragraphs[p].LineBreakAfter && p + 1 < paragraphs.Count) {
				p++;
				s = s.TrimEnd() + "\\\n" + Inline(doc, paragraphs[p].Start, paragraphs[p].End, noBold, false, context, paragraphs[p].Quote).TrimStart();
			}
			return s;
		}

		int openAlign = 0;
		// Alignment that Markdown has no syntax for: <div align="..."> around the blocks, with blank lines inside so
		// that the Markdown in it is still read as Markdown.
		void SetAlign(int align) {
			if (align == openAlign) return;
			if (openAlign != 0) {
				output.Add("");
				output.Add("</div>");
				previous = "html";
				previousQuote = 0;
			}
			if (align != 0) {
				if (previous != null) output.Add("");
				output.Add($"<div align=\"{align switch { 1 => "center", 2 => "right", _ => "justify" }}\">");
				output.Add("");
				previous = null;
			}
			openAlign = align;
		}

		void Separate(string kind, int quote) {
			if (previous == null) return;
			if (kind == "item" && previous == "item") return;    // tight lists
			output.Add(quote > 0 && previousQuote > 0 ? string.Concat(Enumerable.Repeat(">", Math.Min(quote, previousQuote))) : "");
		}

		for (int p = 0; p < paragraphs.Count; p++) {
			var para = paragraphs[p];
			if (para.Kind == "pagebreak") {
				// (a list after it is a list of its own, which starts at its number)
				SetAlign(0);
				Separate("html", 0);
				output.Add(PageBreaks.Html);
				previous = "html";
				previousQuote = 0;
				gap = false;
				continue;
			}
			if (para.Kind == "empty") {
				if (previous == "item") gap = true;
				continue;
			}

			SetAlign(para.Kind is "paragraph" or "heading" or "rule" && para.Quote == 0 ? para.Align : 0);

			if (para.Kind == "code") {
				var lines = new List<string>();
				for (; p < paragraphs.Count && paragraphs[p].Kind == "code"; p++) lines.Add(paragraphs[p].Text.Replace("\v", "\n"));
				p--;
				string fence = "```";
				while (lines.Any(l => l.Contains(fence))) fence += "`";
				Separate("code", 0);
				output.Add(fence);
				output.AddRange(lines);
				output.Add(fence);
				previous = "code";
				previousQuote = 0;
				continue;
			}

			if (para.Kind == "rtable" && para.RealTable is { } realTable) {
				Separate("table", 0);
				int count = realTable.Rows.Max(r => r.Cells.Count);
				bool boldHeader = realTable.Rows[0].Cells.All(c => c.End > c.Start && doc.Range2(c.Start, c.End).Font.Bold == tomConstants.tomTrue);
				var tableLines = new List<string>();
				for (int r = 0; r < realTable.Rows.Count; r++) {
					var cells = realTable.Rows[r].Cells.Select(c => Inline(doc, c.Start, c.End, r == 0 && boldHeader, true, context, 0)
						.Replace("\r", "<br>").Replace("\\\n", "<br>").Trim()).ToList();
					while (cells.Count < count) cells.Add("");
					tableLines.Add("| " + string.Join(" | ", cells) + " |");
					if (r == 0) tableLines.Add("|" + string.Concat(Enumerable.Repeat(" --- |", count)));
				}
				output.AddRange(AlignTable(tableLines));
				previous = "table";
				previousQuote = 0;
				continue;
			}

			if (para.Kind == "table") {
				int first = p;
				while (p + 1 < paragraphs.Count && paragraphs[p + 1].Kind == "table") p++;
				if (p > first) {
					Separate("table", 0);
					var rows = paragraphs.GetRange(first, p - first + 1);
					int count = rows.Max(r => r.Text.Count(ch => ch == '\t') + 1);
					bool boldHeader = CellRanges(rows[0]).All(r => r.End > r.Start && doc.Range2(r.Start, r.End).Font.Bold == tomConstants.tomTrue);
					var tableLines = new List<string>();
					for (int r = 0; r < rows.Count; r++) {
						var cells = CellRanges(rows[r]).Select(c => Inline(doc, c.Start, c.End, r == 0 && boldHeader, true, context, 0).Trim()).ToList();
						while (cells.Count < count) cells.Add("");
						tableLines.Add("| " + string.Join(" | ", cells) + " |");
						if (r == 0) tableLines.Add("|" + string.Concat(Enumerable.Repeat(" --- |", count)));
					}
					output.AddRange(AlignTable(tableLines));
					previous = "table";
					previousQuote = 0;
					continue;
				}
				para.Kind = "paragraph";    // a single row is just a paragraph with tabs
			}

			if (para.Kind == "rule") {
				Separate("rule", para.Quote);
				output.Add(Prefix(para.Quote) + "---");
				previous = "rule";
				previousQuote = para.Quote;
				continue;
			}

			if (para.Kind == "item") {
				int level = para.Level;
				bool interrupted = previous == "item" && gap;
				if (previous != "item" || gap) { counters.Clear(); kinds.Clear(); columns.Clear(); runStarts.Clear(); }
				gap = false;
				level = Math.Min(level, counters.Count);    // no skipped levels
				if (level < counters.Count - 1) {           // back from a deeper level
					counters.RemoveRange(level + 1, counters.Count - level - 1);
					kinds.RemoveRange(level + 1, kinds.Count - level - 1);
					columns.RemoveRange(level + 1, columns.Count - level - 1);
					runStarts.RemoveRange(level + 1, runStarts.Count - level - 1);
				}
				// A new list starts where the kind of list changes (bullets or numbers, numbering style, punctuation: as in
				// CommonMark, where "1." and "1)" are different lists), or where the start value changes (restarted numbering).
				var kind = para.Ordered ? (para.NumberStyle, para.Punctuation) : (tomConstants.tomListBullet, 0);
				bool startsList;
				// (legal items go on where their number is the next one)
				if (level < counters.Count && kinds[level] == kind && (!para.Ordered || (para.Legal ? para.Number == counters[level] + 1 : runStarts[level] == para.Number))) {
					counters[level]++;
					startsList = false;
				} else {
					if (level == counters.Count) { counters.Add(para.Number); kinds.Add(kind); columns.Add(0); runStarts.Add(para.Number); }
					else { counters[level] = para.Number; kinds[level] = kind; runStarts[level] = para.Number; }
					startsList = true;
				}
				bool newList = (previous == "item" || interrupted) && level == 0 && startsList;
				// Markdown joins two numbered lists that follow each other when their markers end alike ("1." and "1.", even
				// with different numbering styles, or an ordinary and a legal list): an HTML comment keeps them apart.
				static char MarkerEnd((int Style, int Punctuation) k) => k.Punctuation is 0 or 0x10000 ? ')' : '.';
				bool keepApart = newList && para.Ordered && lastTopKind is { } lastKind && lastKind.Item1 != tomConstants.tomListBullet && MarkerEnd(lastKind) == MarkerEnd(kind);
				if (level == 0) lastTopKind = kind;
				int indent = level == 0 ? 0 : columns[level - 1];
				string punctuation = para.Punctuation is 0 or 0x10000 ? ")" : ".";    // "(1)" has no Markdown form: 1)
				string marker = para.Ordered ? counters[level].ToString(CultureInfo.InvariantCulture) + punctuation : "-";
				columns[level] = indent + marker.Length + 1;
				string text = WithLineBreaks(ref p, false);
				// [type=] where the numbering isn't GitHub's default for the level; [type=legal] on a legal list's top level only
				// ("(a)" has no Markdown form: its items are written "1)" with [type=(a)])
				string? typeName = para.Legal ? (level == 0 ? "legal" : null)
					: para.Punctuation == tomConstants.tomListParentheses ? "(" + TypeOf(para.NumberStyle) + ")"
					: para.NumberStyle != DefaultNumberStyle(level) ? TypeOf(para.NumberStyle) : null;
				if (para.Ordered && startsList && typeName != null) {
					int firstBreak = text.IndexOf("\\\n", StringComparison.Ordinal);
					text = firstBreak < 0 ? text + $" [type={typeName}]" : text.Insert(firstBreak, $" [type={typeName}]");
				}
				if (keepApart) { output.Add(""); output.Add("<!-- -->"); output.Add(""); }
				else if (newList) output.Add("");    // a list right after a different one: a blank line between them
				else Separate("item", 0);
				output.Add(new string(' ', indent) + marker + " " + EscapeLineStart(text).Replace("\n", "\n" + new string(' ', columns[level])));
				previous = "item";
				previousQuote = 0;
				continue;
			}

			string inline = WithLineBreaks(ref p, para.Kind == "heading").Trim();
			if (inline == "") continue;
			Separate(para.Kind, para.Quote);
			string prefix = Prefix(para.Quote);
			output.Add(para.Kind == "heading"
				? prefix + new string('#', para.Level) + " " + inline.Replace("\\\n", "<br>").Replace("\n", " ")
				: prefix + EscapeLineStart(inline).Replace("\n", "\n" + prefix));
			previous = para.Kind;
			previousQuote = para.Quote;
		}
		SetAlign(0);
		return string.Join("\n", output) + "\n";
	}

	static string Prefix(int quote) => quote == 0 ? "" : string.Concat(Enumerable.Repeat("> ", quote));

	static List<ParagraphRecord> ReadParagraphs(ITextDocument2 doc) {
		var list = new List<ParagraphRecord>();
		Dictionary<int, (string Number, int Series, int[] Counters)>? legalNumbers = null;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var tables = RtfTables.Find(doc);
		var range = doc.Range2(0, 0);
		range.Expand(tomConstants.tomParagraph);
		while (true) {
			int start = range.Start, end = range.End;
			// a real table is one record; its paragraphs are skipped
			if (tables.FirstOrDefault(t => start >= t.Start && start < t.End) is { } table) {
				list.Add(new ParagraphRecord { Start = table.Start, End = table.End, Kind = "rtable", RealTable = table });
				if (table.End + 1 >= storyEnd) break;
				range.SetRange(table.End, table.End);
				range.Expand(tomConstants.tomParagraph);
				continue;
			}
			string text = range.Text ?? "";
			bool lineBreak = false;
			if (text.EndsWith('\r')) { text = text[..^1]; end--; }
			else if (text.EndsWith('\v')) { text = text[..^1]; end--; lineBreak = true; }
			var info = new ParagraphRecord { Start = start, End = end, Text = text, LineBreakAfter = lineBreak };
			var para = range.Para;
			int listType = para.ListType & 0xFFFF;
			var font = doc.Range2(start, start + 1).Font;
			var block = BlockStyleOf(font.Size, font.Bold == tomConstants.tomTrue, font.Name);
			if (text.Contains(PageBreaks.Break) && text.Trim() == "") info.Kind = "pagebreak";
			else if (text.Trim() == "" && !text.Contains('\uFFFC')) info.Kind = "empty";
			else if (listType != tomConstants.tomListNone) {
				info.Kind = "item";
				info.Level = para.ListLevelIndex;
				info.Ordered = listType != tomConstants.tomListBullet;
				info.Number = Math.Max(1, para.ListStart);
				info.NumberStyle = listType;
				info.Punctuation = para.ListType & 0x70000;
				if (LegalLists.IsLegal(para.ListType)) {
					// a legal item: its own number at its level as Markdown's (3 for 2.3), without the text number
					info.Legal = true;
					info.NumberStyle = tomConstants.tomListNumberAsArabic;
					info.Punctuation = LegalLists.NoNumber;
					legalNumbers ??= LegalLists.Numbers(LegalLists.Paragraphs(doc));
					info.Number = legalNumbers.TryGetValue(start, out var legalNumber) ? Math.Max(1, legalNumber.Counters[Math.Min(9, info.Level)]) : 1;
					int skip = LegalLists.Prefix.Match(text).Length;
					info.Start += skip;
					info.Text = text[skip..];
				}
			}
			// (what the paragraph is: its style, when the editor says (BlockOf), else its font)
			int kind = BlockOf?.Invoke(doc, start) ?? (block is BlockStyle b ? (int)b : 0);
			if (listType != tomConstants.tomListNone || info.Kind is "pagebreak" or "empty") { }
			else if (kind == (int)BlockStyle.Preformatted) info.Kind = "code";
			else if (kind is >= 1 and <= 6) { info.Kind = "heading"; info.Level = kind; }
			else if (text.Trim().Length >= 3 && text.Trim().All(ch => ch == '─')) info.Kind = "rule";
			else if (text.Contains('\t')) info.Kind = "table";
			info.Align = para.Alignment is >= 1 and <= 3 ? para.Alignment : 0;
			if (info.Kind is "paragraph" or "rule" or "heading" && para.LeftIndent >= ListIndent)
				info.Quote = Math.Max(1, (int)Math.Round(para.LeftIndent / QuoteIndent));
			list.Add(info);
			if (end + 1 >= storyEnd || range.End <= start) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
		return list;
	}

	sealed class ParagraphRecord
	{
		public int Start, End, Level, Quote;
		public string Text = "", Kind = "paragraph";
		public bool Ordered;
		public int Number = 1;
		public bool LineBreakAfter;    // ends with a line break (\v), not a paragraph mark
		public int Align, NumberStyle, Punctuation;
		public bool Legal;
		public RtfTables.Table? RealTable;
	}

	// What Inline needs to know about the document: its default font (formatting that differs is written as HTML),
	// whether font sizes are free (not block styles), and where the next image comes from.
	sealed record InlineContext(string FontName, float FontSize, bool FreeSizes, Func<(string Source, string Alt, string Size)> ImageSource);

	static readonly int QuoteColor = 0x555555, CodeBackground = 0xEEEEEE;

	static IEnumerable<(int Start, int End)> CellRanges(ParagraphRecord row) {
		int cell = row.Start;
		for (int i = 0; i < row.Text.Length; i++)
			if (row.Text[i] == '\t') {
				yield return (cell, row.Start + i);
				cell = row.Start + i + 1;
			}
		yield return (cell, row.End);
	}

	// The Markdown of a range: runs of the same formatting (TOM tomCharFormat), with links, images and code spans.
	static string Inline(ITextDocument2 doc, int start, int end, bool noBold, bool inTable, InlineContext context, int quote) {
		var runs = new List<(string Text, Style Style, string Url, bool Image, Css Css)>();
		var range = doc.Range2(start, start);
		while (range.Start < end) {
			if (range.MoveEnd(tomConstants.tomCharFormat, 1) == 0) break;
			if (range.End > end) range.End = end;
			var font = range.Font;
			string url = CleanUrl(range.URL);
			var style = Style.None;
			if (font.Bold == tomConstants.tomTrue && !noBold) style |= Style.Bold;
			if (font.Italic == tomConstants.tomTrue) style |= Style.Italic;
			if (font.StrikeThrough == tomConstants.tomTrue) style |= Style.Strike;
			if (font.Underline > 0 && url == "") style |= Style.Underline;
			if (font.Superscript == tomConstants.tomTrue) style |= Style.Superscript;
			if (font.Subscript == tomConstants.tomTrue) style |= Style.Subscript;
			if (MonospaceFonts.Contains(font.Name ?? "", StringComparer.OrdinalIgnoreCase)) style |= Style.Code;
			// formatting that Markdown cannot express (written as <span style>): colors, other fonts, other sizes
			int foreColor = font.ForeColor, backColor = font.BackColor;
			string fontName = font.Name ?? "";
			float fontSize = font.Size;
			var css = new Css(
				foreColor == tomConstants.tomAutoColor || foreColor == 0 || (quote > 0 && foreColor == QuoteColor) || url != "" ? null : foreColor,
				backColor == tomConstants.tomAutoColor || backColor == 0xFFFFFF || (style.HasFlag(Style.Code) && backColor == CodeBackground) ? null : backColor,
				style.HasFlag(Style.Code) || fontName == "" || fontName.Equals(context.FontName, StringComparison.OrdinalIgnoreCase) ? null : fontName,
				!context.FreeSizes || noBold || fontSize <= 0 || Math.Abs(fontSize - context.FontSize) < 0.3f ? null : fontSize);
			// An image is one character (U+FFFC) and can be part of a run with the text around it.
			string text = range.Text ?? "";
			int runStart = range.Start, from = 0;
			for (int k = 0; k <= text.Length; k++) {
				if (k < text.Length && text[k] != '\uFFFC') continue;
				string piece = text[from..k];
				if (url != "" && from == 0) piece = StripFieldCode(piece);
				if (piece.Length > 0) runs.Add((piece, style, url, false, css));
				if (k < text.Length) {
					var image = doc.Range2(runStart + k, runStart + k + 1);
					runs.Add((image.GetText2(tomConstants.tomTextize) ?? "", Style.None, CleanUrl(image.URL), true, default));
				}
				from = k + 1;
			}
			range.Collapse(tomConstants.tomCollapseEnd);
		}

		var sb = new StringBuilder();
		var open = new List<Style>();
		string currentUrl = "";
		int linkStart = 0;
		Css currentCss = default;
		var order = new[] { Style.Bold, Style.Italic, Style.Strike, Style.Underline, Style.Superscript, Style.Subscript };

		void Close(Style style) {
			string marker = Marker(style, false);
			int spaces = 0;
			while (spaces < sb.Length && sb[sb.Length - 1 - spaces] == ' ') spaces++;
			sb.Insert(sb.Length - spaces, marker);
		}
		void CloseAll() { for (int i = open.Count - 1; i >= 0; i--) Close(open[i]); open.Clear(); }
		void CloseSpan() {
			if (currentCss.IsEmpty) return;
			CloseAll();
			int spaces = 0;
			while (spaces < sb.Length && sb[sb.Length - 1 - spaces] == ' ') spaces++;
			sb.Insert(sb.Length - spaces, "</span>");
			currentCss = default;
		}
		void CloseLink() {
			if (currentUrl == "") return;
			CloseSpan();
			CloseAll();
			string label = sb.ToString(linkStart + 1, sb.Length - linkStart - 1);
			if (label == currentUrl || label == Escape(currentUrl, false) || "mailto:" + label == currentUrl) {
				sb.Length = linkStart;
				sb.Append('<').Append(currentUrl.StartsWith("mailto:") && "mailto:" + label == currentUrl ? label : currentUrl).Append('>');    // <https://...>
			} else {
				sb.Append("](").Append(Destination(currentUrl)).Append(')');
			}
			currentUrl = "";
		}

		foreach (var (text, style, url, image, css) in runs) {
			if (url != currentUrl) {
				CloseSpan();
				CloseLink();
				if (url != "") { linkStart = sb.Length; sb.Append('['); currentUrl = url; }
			}
			if (image) {
				CloseSpan();
				CloseAll();
				var (source, alt, size) = context.ImageSource();
				sb.Append("![").Append(Escape(alt != "" ? alt : text.Trim(), inTable)).Append("](").Append(Destination(source)).Append(')').Append(size);
				continue;
			}
			if (css != currentCss && text.Trim() != "") {
				CloseSpan();
				if (!css.IsEmpty) {
					sb.Append("<span style=\"").Append(css.ToString()).Append("\">");
					currentCss = css;
				}
			}
			// emphasis must not start or end on a space: leading spaces go before the markers
			string body = text;
			var wanted = body.Trim() == "" ? Style.None : style & ~Style.Code;
			int keep = open.Count;
			for (int i = 0; i < open.Count; i++) if (!wanted.HasFlag(open[i])) { keep = i; break; }
			for (int i = open.Count - 1; i >= keep; i--) Close(open[i]);
			open.RemoveRange(keep, open.Count - keep);
			int lead = body.Length - body.TrimStart(' ').Length;
			sb.Append(body, 0, lead);
			body = body[lead..];
			foreach (var flag in order)
				if (wanted.HasFlag(flag) && !open.Contains(flag)) { sb.Append(Marker(flag, true)); open.Add(flag); }
			if (style.HasFlag(Style.Code) && body.Trim() != "") {
				int longest = Regex.Matches(body, "`+").Select(m => m.Length).DefaultIfEmpty(0).Max();
				string ticks = new('`', longest + 1);
				string pad = body.StartsWith('`') || body.EndsWith('`') ? " " : "";
				sb.Append(ticks).Append(pad).Append(body.Replace("\v", " ")).Append(pad).Append(ticks);
			} else {
				sb.Append(Escape(body, inTable));
			}
		}
		CloseSpan();
		CloseLink();
		CloseAll();
		return sb.ToString();
	}

	static string Marker(Style style, bool opening) => style switch {
		Style.Bold => "**",
		Style.Italic => "*",
		Style.Strike => "~~",
		Style.Underline => opening ? "<u>" : "</u>",
		Style.Superscript => opening ? "<sup>" : "</sup>",
		Style.Subscript => opening ? "<sub>" : "</sub>",
		_ => ""
	};

	// Escapes what Markdown would otherwise read as markup; \v (a line break) becomes "\" + newline.
	static string Escape(string text, bool inTable) {
		var sb = new StringBuilder();
		for (int i = 0; i < text.Length; i++) {
			char c = text[i];
			char before = i > 0 ? text[i - 1] : ' ', after = i + 1 < text.Length ? text[i + 1] : ' ';
			switch (c) {
				case '\\' or '`' or '*' or '[' or ']':
					sb.Append('\\').Append(c);
					break;
				case '_' when !(char.IsLetterOrDigit(before) && char.IsLetterOrDigit(after)):
				case '~' when before == '~' || after == '~':
				case '<' when char.IsLetter(after) || after == '/' || after == '!' || after == '?':
				case '|' when inTable:
					sb.Append('\\').Append(c);
					break;
				case '\v':
					sb.Append(inTable ? "<br>" : "\\\n");
					break;
				default:
					sb.Append(c);
					break;
			}
		}
		return sb.ToString();
	}

	// A paragraph that starts like a heading, quote, list item or rule gets that start escaped.
	static string EscapeLineStart(string text) {
		var match = Regex.Match(text, @"^(\d{1,9})([.)])(\s|$)");
		if (match.Success) return match.Groups[1].Value + "\\" + text[match.Groups[1].Length..];
		if (Regex.IsMatch(text, @"^(#{1,6}(\s|$)|>|[-+](\s|$)|=+\s*$)")) return "\\" + text;
		return text;
	}

	static string Destination(string url) =>
		url.Any(char.IsWhiteSpace) || url.Count(c => c == '(') != url.Count(c => c == ')') ? "<" + url + ">" : url;

	static string CleanUrl(string? url) =>
		string.IsNullOrEmpty(url) ? "" : url.Length >= 2 && url[0] == '"' && url[^1] == '"' ? url[1..^1] : url;

	static string StripFieldCode(string text) {
		if (!text.StartsWith("HYPERLINK \"", StringComparison.Ordinal)) return text;
		int quote = text.IndexOf('"', 11);
		return quote < 0 ? text : text[(quote + 1)..];
	}

	// ================================================================ the library's block styles (internal)

	static readonly Type? blockStyleHelper = typeof(RichEditBox).Assembly.GetType("EllipticBit.RichEditorNET.Formatting.BlockStyleHelper");
	static readonly MethodInfo? blockStyleFromFont = blockStyleHelper?.GetMethod("GetBlockStyleFromFont", BindingFlags.NonPublic | BindingFlags.Static);

	static T BlockConstant<T>(string name, T fallback) =>
		blockStyleHelper?.GetField(name, BindingFlags.NonPublic | BindingFlags.Static)?.GetValue(null) is T value ? value : fallback;

	internal static float HeadingSize(int level) => BlockConstant("HeadingPointSizes", new[] { 0f, 24f, 18f, 14f, 12f, 10f, 8f })[Math.Clamp(level, 1, 6)];

	/// What a paragraph is, when the editor knows (its style): 0 body text, 1-6 a heading, 7 code (BlockStyle.Preformatted);
	/// null: as its font says.
	public static Func<ITextDocument2, int, int?>? BlockOf;

	static BlockStyle? BlockStyleOf(float size, bool bold, string? name) {
		if (blockStyleFromFont != null) return (BlockStyle)blockStyleFromFont.Invoke(null, new object?[] { size, bold, name })!;
		if (!bold && size == 10f && string.Equals(name, "Courier New", StringComparison.OrdinalIgnoreCase)) return BlockStyle.Preformatted;
		int level = Array.IndexOf(new[] { 0f, 24f, 18f, 14f, 12f, 10f, 8f }, size);
		return bold && level > 0 ? (BlockStyle)level : null;
	}
}
