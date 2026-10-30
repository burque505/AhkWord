// Headers and footers of a document, for AhkWord (AhkWord.cs, through #load; uses the Open XML SDK packages
// that rte_office5.cs lists).
// A document's header and footer each have a left, a center and a right part, with the fields {title} {file} {page}
// {pages} {date} {time} (PageBands). (Iteration 55) A part can have several lines ("\n" between them): the band's line
// n is each part's line n (a paragraph of left <tab> center <tab> right in the files); the header is that far from
// the page's top edge and the footer from its bottom edge (Word's "Header from top", "Footer from bottom"), and one
// that reaches past the margin pushes the text in, as in Word. They are printed, and kept in the files:
//   RTF            real \header / \footer groups (one line: left <tab> center <tab> right, at center and right tab
//                  stops), with PAGE, NUMPAGES, DATE and TIME fields; the page size and margins of Page Setup; and the
//                  exact parts in {\*\rtewriterheaderfooter ...}, which other programs ignore
//   Word, OpenDocument
//                  the same, made by Word or LibreOffice from that RTF, or by the built-in converter (header and footer
//                  parts); the exact parts in a custom XML part (.docx) or a user-defined property (.odt)
//   HTML           <meta name="rte-writer-header-footer" content="...">
//   Markdown       <!-- rte-writer-header-footer: ... --> as the last line
// Read back: the exact parts if the file has them and its real header and footer are still the ones AhkWord wrote
// (if they were edited in Word or LibreOffice, those win); otherwise the real header and footer, with Word's and
// LibreOffice's page number, page count, date and time fields as {page} {pages} {date} {time}.
using System.Drawing;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

/// Iteration 55: the font of a header's or footer's text (all three parts, the first page's too): its name, size in
/// points, bold, italic and underline.
public sealed record BandFont
{
	public string Name { get; init; } = "Segoe UI";
	public float Size { get; init; } = 8;
	public bool Bold { get; init; }
	public bool Italic { get; init; }
	public bool Underline { get; init; }

	[System.Text.Json.Serialization.JsonIgnore] public FontStyle Style => (Bold ? FontStyle.Bold : 0) | (Italic ? FontStyle.Italic : 0) | (Underline ? FontStyle.Underline : 0);

	/// The font at a size of its own (points), e.g. at the zoom; a name Windows doesn't have gets GDI+'s fallback.
	public Font Create(float size) => new(Name, Math.Max(1f, size), Style, GraphicsUnit.Point);

	public static BandFont From(Font font) => new() { Name = font.Name, Size = font.SizeInPoints, Bold = font.Bold, Italic = font.Italic, Underline = font.Underline };

	/// "Times New Roman 10 pt, bold, italic"
	public string Describe() => $"{Name} {Size.ToString("0.#", System.Globalization.CultureInfo.CurrentCulture)} pt"
		+ string.Concat(new[] { (Bold, "bold"), (Italic, "italic"), (Underline, "underlined") }.Where(x => x.Item1).Select(x => ", " + x.Item2));

	/// The text as it was before fonts could be chosen (a band without a font of its own), drawn in grey.
	public static readonly BandFont Plain = new();
}

/// Iteration 55: a word's or a line's own font in a header or footer (the band's font elsewhere).
public sealed record RunFormat(string Name, float Size, bool Bold, bool Italic, bool Underline);

/// Iteration 55: a header's or footer's part with words in fonts of their own. A mark RunStart "name;size;biu" RunEnd
/// sets the font of the text after it, an empty one (RunStart RunEnd) goes back to the band's font; marks carry across
/// the part's lines. A part without marks is all in the band's font.
public static class BandText
{
	public const char RunStart = '\uE030', RunEnd = '\uE031';
	static readonly Regex Mark = new("\uE030([^\uE031]*)\uE031");

	/// The part's text without its marks.
	public static string Plain(string? part) => Mark.Replace(part ?? "", "");

	static string Spec(RunFormat? f) => f == null ? "" : string.Join(";", f.Name, f.Size.ToString(System.Globalization.CultureInfo.InvariantCulture), (f.Bold ? "b" : "") + (f.Italic ? "i" : "") + (f.Underline ? "u" : ""));
	static RunFormat? Parse(string spec) {
		var p = spec.Split(';');
		if (p.Length < 3 || p[0] == "" || !float.TryParse(p[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float size)) return null;
		return new RunFormat(p[0], size, p[2].Contains('b'), p[2].Contains('i'), p[2].Contains('u'));
	}

	/// The part's runs: text and its own font (null: the band's), neighbours in the same font joined.
	public static List<(string Text, RunFormat? Format)> Runs(string? part) {
		var runs = new List<(string Text, RunFormat? Format)>();
		RunFormat? format = null;
		int at = 0;
		part ??= "";
		void Add(string text) {
			if (text == "") return;
			if (runs.Count > 0 && Equals(runs[^1].Format, format)) runs[^1] = (runs[^1].Text + text, format);
			else runs.Add((text, format));
		}
		foreach (Match m in Mark.Matches(part)) { Add(part[at..m.Index]); format = Parse(m.Groups[1].Value); at = m.Index + m.Length; }
		Add(part[at..]);
		return runs;
	}

	public static string FromRuns(IEnumerable<(string Text, RunFormat? Format)> runs) {
		var sb = new StringBuilder();
		RunFormat? format = null;
		foreach (var (text, f) in runs) {
			if (text == "") continue;
			if (!Equals(f, format)) { sb.Append(RunStart).Append(Spec(f)).Append(RunEnd); format = f; }
			sb.Append(text);
		}
		return sb.ToString();
	}

	/// The part's lines, each its runs.
	public static List<List<(string Text, RunFormat? Format)>> Lines(string? part) {
		var lines = new List<List<(string Text, RunFormat? Format)>> { new() };
		foreach (var (text, f) in Runs(part)) {
			var pieces = text.Split('\n');
			for (int i = 0; i < pieces.Length; i++) {
				if (i > 0) lines.Add(new());
				if (pieces[i] != "") lines[^1].Add((pieces[i], f));
			}
		}
		return lines;
	}

	public static BandFont Resolve(BandFont band, RunFormat? f) => f == null ? band : new BandFont { Name = f.Name, Size = f.Size, Bold = f.Bold, Italic = f.Italic, Underline = f.Underline };
	public static RunFormat? FormatOf(BandFont band, BandFont font) => font == band ? null : new RunFormat(font.Name, font.Size, font.Bold, font.Italic, font.Underline);

	/// The band's height: its lines' (each its tallest font's line height, from lineHeight; sizes times scale).
	public static float Height(string[] parts, BandFont band, Func<Font, float> lineHeight, float scale = 1) {
		var cells = PageBands.Parts(parts).Select(Lines).ToArray();
		var fonts = new Dictionary<BandFont, Font>();
		try {
			float height = 0;
			for (int n = 0; n < cells.Max(c => c.Count); n++)
				height += LineFonts(cells, n, band).Select(f => lineHeight(Get(fonts, f, scale))).Max();
			return height;
		} finally { foreach (var f in fonts.Values) f.Dispose(); }
	}

	// A line's fonts (the band's if it has no text).
	static IEnumerable<BandFont> LineFonts(List<List<(string Text, RunFormat? Format)>>[] cells, int n, BandFont band) {
		var fonts = cells.Where(c => n < c.Count).SelectMany(c => c[n]).Select(r => Resolve(band, r.Format)).ToList();
		return fonts.Count > 0 ? fonts : new[] { band };
	}

	static Font Get(Dictionary<BandFont, Font> fonts, BandFont f, float scale) {
		if (!fonts.TryGetValue(f, out var font)) fonts[f] = font = f.Create(f.Size * scale);
		return font;
	}

	/// Draws a band's parts (fields filled in) a line below the other from top, each part's line n on the band's line n:
	/// the left part from left, the center one centred, the right one ending at left + width, on the line's baseline.
	/// Sizes times scale (the zoom); the band's own font in plain, a word's own font in formatted. Returns its height.
	public static float Draw(Graphics g, string[] parts, BandFont band, float left, float width, float top, float scale, Brush plain, Brush formatted) {
		var cells = PageBands.Parts(parts).Select(Lines).ToArray();
		var fonts = new Dictionary<BandFont, Font>();
		using var format = (StringFormat)StringFormat.GenericTypographic.Clone();
		format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces | StringFormatFlags.NoWrap;
		float Ascent(Font f) => f.GetHeight(g) * f.FontFamily.GetCellAscent(f.Style) / Math.Max(1, f.FontFamily.GetLineSpacing(f.Style));
		var clip = g.Clip;
		try {
			g.SetClip(new RectangleF(left, top - 1000, width, 100000), System.Drawing.Drawing2D.CombineMode.Intersect);
			float y = top;
			for (int n = 0; n < cells.Max(c => c.Count); n++) {
				var lineFonts = LineFonts(cells, n, band).Select(f => Get(fonts, f, scale)).ToList();
				float height = lineFonts.Max(f => f.GetHeight(g)), baseline = y + lineFonts.Max(Ascent);
				for (int c = 0; c < 3; c++) {
					if (n >= cells[c].Count || cells[c][n].Count == 0) continue;
					var runs = cells[c][n].Select(r => (r.Text, Font: Get(fonts, Resolve(band, r.Format), scale), Own: r.Format != null)).ToList();
					var widths = runs.Select(r => g.MeasureString(r.Text, r.Font, PointF.Empty, format).Width).ToList();
					float total = widths.Sum();
					float x = c == 0 ? left : c == 1 ? left + (width - total) / 2 : left + width - total;
					for (int i = 0; i < runs.Count; i++) {
						g.DrawString(runs[i].Text, runs[i].Font, runs[i].Own ? formatted : plain, x, baseline - Ascent(runs[i].Font), format);
						x += widths[i];
					}
				}
				y += height;
			}
			return y - top;
		} finally {
			g.Clip = clip;
			foreach (var f in fonts.Values) f.Dispose();
		}
	}
}

/// A document's header and footer: whether each is shown, and its left, center and right parts.
public sealed class PageBands
{
	public bool ShowHeader { get; set; }
	public string[] Header { get; set; } = { "", "", "" };
	public bool ShowFooter { get; set; }
	public string[] Footer { get; set; } = { "", "", "" };
	/// The parts as they were written into the file's real header and footer ({title} and {file} filled in), to tell
	/// whether another program changed them since.
	public string[]? WrittenHeader { get; set; }
	public string[]? WrittenFooter { get; set; }

	/// Iteration 49: the first page has a header and footer of its own (FirstHeader / FirstFooter, which may be empty:
	/// none on that page). (Iteration 55) Its own in every way: shown or not (ShowFirstHeader), font and line; null: as
	/// the other pages' (documents from before). ShowHeader and ShowFooter are then for the other pages.
	[System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingDefault)] public bool DifferentFirst { get; set; }
	public string[]? FirstHeader { get; set; }
	public string[]? FirstFooter { get; set; }
	public string[]? WrittenFirstHeader { get; set; }
	public string[]? WrittenFirstFooter { get; set; }
	public bool? ShowFirstHeader { get; set; }
	public bool? ShowFirstFooter { get; set; }
	public BandFont? FirstHeaderFont { get; set; }
	public BandFont? FirstFooterFont { get; set; }
	public bool? FirstHeaderLine { get; set; }
	public bool? FirstFooterLine { get; set; }

	bool First(int page) => DifferentFirst && page == 1;
	/// Whether a page has a header (footer), its font (null: plain) and its line.
	public bool ShowHeaderOn(int page) => First(page) ? ShowFirstHeader ?? ShowHeader : ShowHeader;
	public bool ShowFooterOn(int page) => First(page) ? ShowFirstFooter ?? ShowFooter : ShowFooter;
	public BandFont? HeaderFontOn(int page) => First(page) ? FirstHeaderFont ?? HeaderFont : HeaderFont;
	public BandFont? FooterFontOn(int page) => First(page) ? FirstFooterFont ?? FooterFont : FooterFont;
	public bool HeaderLineOn(int page) => First(page) ? FirstHeaderLine ?? HeaderLine : HeaderLine;
	public bool FooterLineOn(int page) => First(page) ? FirstFooterLine ?? FooterLine : FooterLine;
	/// Whether any page has a header (footer).
	[System.Text.Json.Serialization.JsonIgnore] public bool AnyHeader => ShowHeader || ShowHeaderOn(1);
	[System.Text.Json.Serialization.JsonIgnore] public bool AnyFooter => ShowFooter || ShowFooterOn(1);
	/// Text in the plain look: Segoe UI 8 pt in grey (no font of its own, or the first page's set plain).
	public static bool IsPlain(BandFont? font) => font == null || font == BandFont.Plain;

	/// Iteration 55: the header's and the footer's font (null: Segoe UI 8 pt in grey, as before).
	public BandFont? HeaderFont { get; set; }
	public BandFont? FooterFont { get; set; }

	/// Iteration 55: from the page's top edge to the header, and from its bottom edge to the footer (1/100 inch; null:
	/// half an inch, Word's).
	public int? HeaderDistance { get; set; }
	public int? FooterDistance { get; set; }
	public const int DefaultDistance = 50;
	/// Iteration 55: a line below the header, above the footer (Word's paragraph border), RuleGap from the text
	/// (1/100 inch).
	[System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingDefault)] public bool HeaderLine { get; set; }
	[System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingDefault)] public bool FooterLine { get; set; }
	public const float RuleGap = 4, RuleWidth = 1;
	[System.Text.Json.Serialization.JsonIgnore] public int HeaderFromEdge => HeaderDistance ?? DefaultDistance;
	[System.Text.Json.Serialization.JsonIgnore] public int FooterFromEdge => FooterDistance ?? DefaultDistance;

	/// Always three parts; line ends as "\n" (blank lines at a part's end kept: space below a header, as in Word).
	public static string[] Parts(string[]? parts) => Enumerable.Range(0, 3).Select(i => parts != null && i < parts.Length ? Lines(parts[i]) : "").ToArray();
	static string Lines(string? part) => (part ?? "").Replace("\r\n", "\n").Replace('\r', '\n');

	/// Each part's lines (at least one, maybe empty).
	public static string[][] LinesOf(string[]? parts) => Parts(parts).Select(p => p.Split('\n')).ToArray();
	/// How many lines the band has: its longest part's (0: all of it blank).
	public static int LineCount(string[]? parts) => Parts(parts).All(p => BandText.Plain(p).Trim() == "") ? 0 : LinesOf(parts).Max(l => l.Length);

	/// From the edge to the far side of a page's header (1/100 inch; 0: none; page 2: every page but a first page of its
	/// own). The page's text starts no nearer the edge than this.
	public int HeaderReach(int page = 2) => HeaderOn(page) is { } parts ? Reach(HeaderFromEdge, HeaderFontOn(page), parts, HeaderLineOn(page)) : 0;
	public int FooterReach(int page = 2) => FooterOn(page) is { } parts ? Reach(FooterFromEdge, FooterFontOn(page), parts, FooterLineOn(page)) : 0;
	static int Reach(int distance, BandFont? font, string[]? parts, bool line = false) =>
		LineCount(parts) == 0 ? 0 : distance + (int)Math.Ceiling(BandText.Height(Parts(parts), font ?? BandFont.Plain, f => f.GetHeight(100f)) + (line ? 2 * RuleGap : 0));

	public PageBands Copy() => new() {
		ShowHeader = ShowHeader, Header = Parts(Header), ShowFooter = ShowFooter, Footer = Parts(Footer),
		DifferentFirst = DifferentFirst, FirstHeader = DifferentFirst ? Parts(FirstHeader) : null, FirstFooter = DifferentFirst ? Parts(FirstFooter) : null,
		HeaderFont = HeaderFont, FooterFont = FooterFont, HeaderDistance = HeaderDistance, FooterDistance = FooterDistance,
		HeaderLine = HeaderLine, FooterLine = FooterLine,
		ShowFirstHeader = DifferentFirst ? ShowFirstHeader : null, ShowFirstFooter = DifferentFirst ? ShowFirstFooter : null,
		FirstHeaderFont = DifferentFirst ? FirstHeaderFont : null, FirstFooterFont = DifferentFirst ? FirstFooterFont : null,
		FirstHeaderLine = DifferentFirst ? FirstHeaderLine : null, FirstFooterLine = DifferentFirst ? FirstFooterLine : null
	};    // (the first page's parts only if it has its own: files without them stay as they were)

	public bool SameAs(PageBands other) => ShowHeader == other.ShowHeader && ShowFooter == other.ShowFooter
		&& Parts(Header).SequenceEqual(Parts(other.Header)) && Parts(Footer).SequenceEqual(Parts(other.Footer))
		&& DifferentFirst == other.DifferentFirst && Equals(HeaderFont, other.HeaderFont) && Equals(FooterFont, other.FooterFont)
		&& HeaderFromEdge == other.HeaderFromEdge && FooterFromEdge == other.FooterFromEdge && HeaderLine == other.HeaderLine && FooterLine == other.FooterLine
		&& (!DifferentFirst || (Parts(FirstHeader).SequenceEqual(Parts(other.FirstHeader)) && Parts(FirstFooter).SequenceEqual(Parts(other.FirstFooter))
			&& ShowHeaderOn(1) == other.ShowHeaderOn(1) && ShowFooterOn(1) == other.ShowFooterOn(1) && Equals(HeaderFontOn(1), other.HeaderFontOn(1))
			&& Equals(FooterFontOn(1), other.FooterFontOn(1)) && HeaderLineOn(1) == other.HeaderLineOn(1) && FooterLineOn(1) == other.FooterLineOn(1)));

	/// The header's parts on a page (1 is the first; null: no header there).
	public string[]? HeaderOn(int page) => !ShowHeaderOn(page) ? null : Parts(First(page) ? FirstHeader : Header);
	public string[]? FooterOn(int page) => !ShowFooterOn(page) ? null : Parts(First(page) ? FirstFooter : Footer);

	static bool Blank(string[]? parts) => Parts(parts).All(p => BandText.Plain(p).Trim() == "");
	/// The parts that go into a file's real header (null: no header there).
	[System.Text.Json.Serialization.JsonIgnore] public string[]? HeaderToWrite => !ShowHeader || Blank(Header) ? null : Parts(Header);
	[System.Text.Json.Serialization.JsonIgnore] public string[]? FooterToWrite => !ShowFooter || Blank(Footer) ? null : Parts(Footer);
	/// The first page's (with DifferentFirst only; null: none there).
	[System.Text.Json.Serialization.JsonIgnore] public string[]? FirstHeaderToWrite => !DifferentFirst || !ShowHeaderOn(1) || Blank(FirstHeader) ? null : Parts(FirstHeader);
	[System.Text.Json.Serialization.JsonIgnore] public string[]? FirstFooterToWrite => !DifferentFirst || !ShowFooterOn(1) || Blank(FirstFooter) ? null : Parts(FirstFooter);

	static readonly JsonSerializerOptions Json = new() { DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull };
	public string ToJson() => JsonSerializer.Serialize(this, Json);
	public static PageBands? FromJson(string json) {
		try { return JsonSerializer.Deserialize<PageBands>(json) is { } b ? Normalized(b) : null; }
		catch (JsonException) { return null; }
	}
	static PageBands Normalized(PageBands b) {
		b.Header = Parts(b.Header); b.Footer = Parts(b.Footer); b.FirstHeader = Parts(b.FirstHeader); b.FirstFooter = Parts(b.FirstFooter);
		return b;
	}
}

/// The page a document is laid out on in other programs (twips): Page Setup's paper and margins.
public sealed record PageLayout(int Width, int Height, int Left, int Right, int Top, int Bottom, bool Landscape, int HeaderY = 720, int FooterY = 720)
{
	public int TextWidth => Math.Max(1440, Width - Left - Right);
}

public static class HeaderFooter
{
	const string RtfDestination = "rtewriterheaderfooter";
	const string MetaName = "rte-writer-header-footer";
	const string XmlNamespace = "urn:rte-writer:header-footer";
	const string OdtProperty = "AhkWord header and footer";
	static readonly Regex Field = new(@"\{(title|file|page|pages|sectionpages|date|time)\}", RegexOptions.IgnoreCase);

	// {title} and {file} become text; the others stay fields (filled in by Word and LibreOffice).
	static string[] Filled(string[] parts, string title, string file) =>
		parts.Select(p => Regex.Replace(p, @"\{(title|file)\}", m => m.Groups[1].Value.Equals("title", StringComparison.OrdinalIgnoreCase) ? title : file, RegexOptions.IgnoreCase)).ToArray();

	/// The bands as written into a file (with WrittenHeader / WrittenFooter), for the hidden exact copy.
	static PageBands Stamped(PageBands bands, string title, string file) {
		var b = bands.Copy();
		b.WrittenHeader = bands.HeaderToWrite is { } h ? Filled(h, title, file) : null;
		b.WrittenFooter = bands.FooterToWrite is { } f ? Filled(f, title, file) : null;
		b.WrittenFirstHeader = bands.FirstHeaderToWrite is { } fh ? Filled(fh, title, file) : null;
		b.WrittenFirstFooter = bands.FirstFooterToWrite is { } ff ? Filled(ff, title, file) : null;
		return b;
	}

	/// A file's own first-page header and footer (iteration 49): whether it has them, and their parts (null: none).
	readonly record struct FirstPage(bool Different, string[]? Header, string[]? Footer);

	// What was read from a file's real header and footer (null: none), against the exact copy (null: none).
	static PageBands? Decide(PageBands? exact, string[]? header, string[]? footer, bool haveReal, FirstPage first = default) {
		if (exact != null) {
			bool same = !haveReal || (Same(exact.WrittenHeader, header) && Same(exact.WrittenFooter, footer)
				&& exact.DifferentFirst == first.Different && (!first.Different || (Same(exact.WrittenFirstHeader, first.Header) && Same(exact.WrittenFirstFooter, first.Footer))));
			if (same) { exact.WrittenHeader = exact.WrittenFooter = exact.WrittenFirstHeader = exact.WrittenFirstFooter = null; return exact; }
		}
		if (!haveReal) return null;
		var bands = new PageBands {
			ShowHeader = header != null, Header = PageBands.Parts(header),
			ShowFooter = footer != null, Footer = PageBands.Parts(footer)
		};
		if (first.Different) {
			bands.DifferentFirst = true;
			bands.FirstHeader = PageBands.Parts(first.Header);
			bands.FirstFooter = PageBands.Parts(first.Footer);
			// (iteration 55: the first page's shown on its own, the others' as the file has them)
			bands.ShowFirstHeader = first.Header != null;
			bands.ShowFirstFooter = first.Footer != null;
		}
		return bands;
	}

	static bool Same(string[]? a, string[]? b) =>
		(a == null || a.All(p => p.Trim() == "")) && (b == null || b.All(p => p.Trim() == ""))
		|| (a != null && b != null && PageBands.Parts(a).Select(Clean).SequenceEqual(PageBands.Parts(b).Select(Clean)));

	static string Clean(string s) => Regex.Replace(BandText.Plain(s).Replace('\u00A0', ' '), @"\s+", " ").Trim();    // (iteration 55: without the words' fonts)

	// A line of header text (tabs between the parts) as left, center and right; one part alone goes where its
	// paragraph is aligned.
	static string[]? PartsOf(string line, int align) {
		line = line.Trim('\r', '\n');
		if (Clean(line.Replace('\t', ' ')) == "") return null;
		var pieces = line.Split('\t');
		if (pieces.Length == 1) return align switch { 1 => new[] { "", pieces[0].Trim(), "" }, 2 => new[] { "", "", pieces[0].Trim() }, _ => new[] { pieces[0].Trim(), "", "" } };
		if (pieces.Length == 2) return new[] { pieces[0].Trim(), pieces[1].Trim(), "" };
		return new[] { pieces[0].Trim(), pieces[1].Trim(), string.Join(" ", pieces.Skip(2)).Trim() };
	}

	// (iteration 55) A band's paragraphs (each one's parts; null: a blank one) as its three parts, a line per paragraph,
	// blank ones kept (they are space in the band, as in Word; null: all blank).
	static string[]? Combined(IEnumerable<string[]?> paragraphs) {
		var rows = paragraphs.ToList();
		if (rows.All(r => r == null)) return null;
		var lines = rows.Select(r => r ?? new[] { "", "", "" }).ToList();
		// (a part's text ends at its last line with text; blank lines at the band's end go in the left part)
		var parts = Enumerable.Range(0, 3).Select(c => string.Join("\n", lines.Select(l => l[c])).TrimEnd('\n')).ToArray();
		parts[0] += new string('\n', Math.Max(0, lines.Count - 1 - parts[0].Count(ch => ch == '\n')));
		return parts;
	}

	// The paragraphs of a band's text ("\n" after each), each with the first paragraph's alignment.
	static string[]? CombinedText(string text, int align) => Combined((text.EndsWith('\n') ? text[..^1] : text).Split('\n').Select(line => PartsOf(line, align)));

	// 1/100 inch from twips, null for Word's half inch
	static int? Distance(int twips) { int d = (int)Math.Round(twips / 14.4); return d == PageBands.DefaultDistance ? null : d; }

	static string FieldOf(string instruction, string result) {
		var m = Regex.Match(instruction, @"^\s*([A-Z]+)(.*)$", RegexOptions.IgnoreCase);
		string name = m.Success ? m.Groups[1].Value.ToUpperInvariant() : "";
		return name switch {
			"PAGE" => "{page}",
			"NUMPAGES" => "{pages}",
			"SECTIONPAGES" => "{sectionpages}",    // (iteration 63)
			"DATE" or "CREATEDATE" or "SAVEDATE" or "PRINTDATE" => Regex.IsMatch(m.Groups[2].Value, @"[Hhm]{1,2}:") && !Regex.IsMatch(m.Groups[2].Value, "[dMy]") ? "{time}" : "{date}",
			"TIME" => "{time}",
			"FILENAME" => m.Groups[2].Value.Contains("\\p", StringComparison.OrdinalIgnoreCase) ? "{file}" : "{title}",
			_ => result
		};
	}

	// ================================================================ RTF

	/// RTF with the header and footer (and the page layout) added, for an .rtf file, Word or LibreOffice. Added in front
	/// of the first paragraph, after any other changes to the RTF.
	public static string AddToRtf(string rtf, PageBands bands, PageLayout layout, string title, string file, float fontSize) {
		if (BodyStart(rtf) < 0) return rtf;
		string groups = BandGroups(ref rtf, bands, layout, title, file, fontSize);
		int body = BodyStart(rtf);
		var sb = new StringBuilder();
		sb.Append($@"\paperw{layout.Width}\paperh{layout.Height}\margl{layout.Left}\margr{layout.Right}\margt{layout.Top}\margb{layout.Bottom}\headery{layout.HeaderY}\footery{layout.FooterY}");
		if (layout.Landscape) sb.Append(@"\landscape");
		sb.Append(@"\sectd");
		if (layout.Landscape) sb.Append(@"\lndscpsxn");
		if (bands.DifferentFirst) sb.Append(@"\titlepg");
		sb.Append($@"\pgwsxn{layout.Width}\pghsxn{layout.Height}\marglsxn{layout.Left}\margrsxn{layout.Right}\margtsxn{layout.Top}\margbsxn{layout.Bottom}\headery{layout.HeaderY}\footery{layout.FooterY} ");
		sb.Append(groups);
		sb.Append(@"{\*\").Append(RtfDestination).Append(' ').Append(RtfEscape(Stamped(bands, title, file).ToJson())).Append('}');
		return rtf.Insert(body, sb.ToString());
	}

	/// A header's and footer's RTF groups (\header, \footer, and the first page's \headerf and \footerf when it has its
	/// own), their fonts added to rtf's font table (iteration 63: each section's own too).
	internal static string BandGroups(ref string rtf, PageBands bands, PageLayout layout, string title, string file, float fontSize) {
		int size = (int)Math.Round(fontSize * 2);
		// (iteration 55) a header's or footer's own font: in the font table, with its size, bold, italic and underline
		string headerFont = bands.HeaderFont is { } hf ? Formatting(ref rtf, hf) : $@"\f0\fs{size}";
		string footerFont = bands.FooterFont is { } ff ? Formatting(ref rtf, ff) : $@"\f0\fs{size}";
		string firstHeaderFont = bands.HeaderFontOn(1) is { } fhf ? Formatting(ref rtf, fhf) : $@"\f0\fs{size}";
		string firstFooterFont = bands.FooterFontOn(1) is { } fff ? Formatting(ref rtf, fff) : $@"\f0\fs{size}";
		// (iteration 55) and the words' own fonts
		var numbers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
		foreach (var parts in new[] { bands.Header, bands.Footer, bands.FirstHeader, bands.FirstFooter })
			foreach (var (_, f) in PageBands.Parts(parts).SelectMany(BandText.Runs))
				if (f != null && !numbers.ContainsKey(f.Name)) numbers[f.Name] = FontNumber(ref rtf, f.Name);
		var sb = new StringBuilder();
		void Band(string destination, string[]? parts, bool empty = false) {
			if (parts == null) {
				if (empty) sb.Append("{\\").Append(destination).Append(@"\pard\plain\par}");    // (the first page's: none there)
				return;
			}
			bool isHeader = destination.StartsWith("header"), first = destination.EndsWith("f");
			string font = isHeader ? (first ? firstHeaderFont : headerFont) : (first ? firstFooterFont : footerFont);
			// (iteration 55) the line: a paragraph border, which Word draws once below (above) paragraphs that share it
			int page = first ? 1 : 2;
			string border = isHeader ? (bands.HeaderLineOn(page) ? @"\brdrb\brdrs\brdrw10\brsp40" : "") : (bands.FooterLineOn(page) ? @"\brdrt\brdrs\brdrw10\brsp40" : "");
			sb.Append("{\\").Append(destination).Append($@"\pard\plain{border}{font}\tqc\tx{layout.TextWidth / 2}\tqr\tx{layout.TextWidth} ");
			// (iteration 55) a paragraph per line
			var lines = Filled(parts, title, file).Select(BandText.Lines).ToArray();
			for (int line = 0; line < lines.Max(l => l.Count); line++)
				sb.Append(string.Join(@"\tab ", lines.Select(l => line < l.Count ? RtfRuns(l[line], numbers) : ""))).Append(@"\par ");
			sb.Length--;
			sb.Append('}');
		}
		Band("header", bands.HeaderToWrite);
		Band("footer", bands.FooterToWrite);
		if (bands.DifferentFirst) {
			Band("headerf", bands.FirstHeaderToWrite, true);
			Band("footerf", bands.FirstFooterToWrite, true);
		}
		return sb.ToString();
	}

	// A band font's RTF: \fN (the font table's entry for its name, added if it has none) \fsN \b \i \ul.
	static string Formatting(ref string rtf, BandFont font) {
		int number = FontNumber(ref rtf, font.Name);
		return $@"\f{number}\fs{(int)Math.Round(font.Size * 2)}" + (font.Bold ? @"\b" : "") + (font.Italic ? @"\i" : "") + (font.Underline ? @"\ul" : "");
	}

	static int FontNumber(ref string rtf, string name) {
		int table = FindGroup(rtf, "fonttbl");
		if (table < 0) {
			// (no font table: one after the \rtf1 header's control words)
			var header = Regex.Match(rtf, @"^\{\\rtf1(\\[a-zA-Z]+-?\d* ?)*");
			rtf = rtf.Insert(header.Length, "{\\fonttbl{\\f0\\fnil " + RtfEscape(name) + ";}}");
			return 0;
		}
		int end = GroupEnd(rtf, table);
		string fonts = rtf[table..(end + 1)];
		int highest = -1;
		// each entry: {\fN control words {\*\panose ...} name;}
		foreach (Match m in Regex.Matches(fonts, @"\{\\f(\d+)((?:\\[^\s\\{};]+ ?|\{[^{}]*\}|\s)*)([^\\{};]*);")) {
			int n = int.Parse(m.Groups[1].Value);
			highest = Math.Max(highest, n);
			if (string.Equals(m.Groups[3].Value.Trim(), name, StringComparison.OrdinalIgnoreCase)) return n;
		}
		foreach (Match m in Regex.Matches(fonts, @"\\f(\d+)")) highest = Math.Max(highest, int.Parse(m.Groups[1].Value));
		rtf = rtf.Insert(end, "{\\f" + (highest + 1) + "\\fnil " + RtfEscape(name) + ";}");
		return highest + 1;
	}

	// (iteration 55) a line's runs: those in a font of their own in a group with it
	static string RtfRuns(List<(string Text, RunFormat? Format)> runs, Dictionary<string, int> numbers) => string.Concat(runs.Select(r => r.Format is not { } f ? RtfText(r.Text)
		: $@"{{\f{numbers[f.Name]}\fs{(int)Math.Round(f.Size * 2)}{(f.Bold ? @"\b" : @"\b0")}{(f.Italic ? @"\i" : @"\i0")}{(f.Underline ? @"\ul" : @"\ulnone")} " + RtfText(r.Text) + "}"));

	// A part as RTF: text, with fields for {page} {pages} {date} {time}.
	static string RtfText(string part) {
		var sb = new StringBuilder();
		int at = 0;
		foreach (Match m in Field.Matches(part)) {
			sb.Append(RtfEscape(part[at..m.Index]));
			string instruction = m.Groups[1].Value.ToLowerInvariant() switch { "page" => "PAGE", "pages" => "NUMPAGES", "sectionpages" => "SECTIONPAGES", "date" => "DATE", _ => "TIME" };
			string shown = instruction switch { "PAGE" or "NUMPAGES" or "SECTIONPAGES" => "1", "DATE" => DateTime.Now.ToShortDateString(), _ => DateTime.Now.ToShortTimeString() };
			sb.Append(@"{\field{\*\fldinst ").Append(instruction).Append(@"}{\fldrslt ").Append(RtfEscape(shown)).Append("}}");
			at = m.Index + m.Length;
		}
		sb.Append(RtfEscape(part[at..]));
		return sb.ToString();
	}

	static string RtfEscape(string text) {
		var sb = new StringBuilder();
		foreach (char c in text) {
			if (c is '\\' or '{' or '}') sb.Append('\\').Append(c);
			else if (c > 126 || c < 32) sb.Append(@"\u").Append((short)c).Append('?');
			else sb.Append(c);
		}
		return sb.ToString();
	}

	// Where the document's text starts: its first \pard (or \sectd) outside groups other than the outermost.
	internal static int BodyStart(string rtf) {
		int depth = 0;
		for (int i = 0; i < rtf.Length; i++) {
			char c = rtf[i];
			if (c == '\\') {
				if (depth == 1 && (string.CompareOrdinal(rtf, i, @"\pard", 0, 5) == 0 || string.CompareOrdinal(rtf, i, @"\sectd", 0, 6) == 0)
					&& (i + 6 > rtf.Length || !char.IsLetter(rtf[i + (rtf[i + 1] == 'p' ? 5 : 6)]))) return i;
				i++;
			} else if (c == '{') depth++;
			else if (c == '}') depth--;
		}
		return -1;
	}

	/// The header and footer of RTF (null if it has neither): AhkWord's exact copy, or the RTF's header and footer.
	public static PageBands? FromRtf(string rtf) {
		PageBands? exact = null;
		var copy = Regex.Match(rtf, @"\{\\\*\\" + RtfDestination + @" ((?:[^\\{}]|\\[\\{}]|\\u-?\d+\?)*)\}");
		if (copy.Success) exact = PageBands.FromJson(RtfUnescape(copy.Groups[1].Value));
		string[]? header = null, footer = null;
		bool any = false, headerLine = false, footerLine = false, firstHeaderLine = false, firstFooterLine = false;
		foreach (var (name, isHeader) in new[] { ("header", true), ("headerr", true), ("footer", false), ("footerr", false) }) {
			if ((isHeader ? header : footer) != null) continue;
			int start = FindGroup(rtf, name);
			if (start < 0) continue;
			any = true;
			var (text, align) = RtfGroupText(rtf, start);
			var parts = CombinedText(text, align);
			if (isHeader) header = parts; else footer = parts;
			// (iteration 55) its line: a border below the header, above the footer
			string group = rtf[start..(GroupEnd(rtf, start) + 1)];
			if (isHeader) headerLine = Regex.IsMatch(group, @"\\brdrb(?![a-z])"); else footerLine = Regex.IsMatch(group, @"\\brdrt(?![a-z])");
		}
		// (iteration 49) the first page's: \titlepg in the first section (before any \sect), \headerf and \footerf
		var first = default(FirstPage);
		var sectionEnd = Regex.Match(rtf, @"\\sect(?![a-zA-Z])");
		if (Regex.IsMatch(sectionEnd.Success ? rtf[..sectionEnd.Index] : rtf, @"\\titlepg(?![a-zA-Z])")) {
			string[]? Read(string name) {
				int start = FindGroup(rtf, name);
				if (start < 0) return null;
				var (groupText, align) = RtfGroupText(rtf, start);
				return CombinedText(groupText, align);
			}
			first = new FirstPage(true, Read("headerf"), Read("footerf"));
			bool Bordered(string name, string word) => FindGroup(rtf, name) is int at and >= 0 && Regex.IsMatch(rtf[at..(GroupEnd(rtf, at) + 1)], @"\\" + word + "(?![a-z])");
			(firstHeaderLine, firstFooterLine) = (Bordered("headerf", "brdrb"), Bordered("footerf", "brdrt"));
			any |= first.Header != null || first.Footer != null || FindGroup(rtf, "headerf") >= 0 || FindGroup(rtf, "footerf") >= 0;
		}
		var bands = Decide(exact, header, footer, any, first);
		// (iteration 55) the distances from the edges, and the lines: the file's own
		if (bands != null) {
			if (header != null) bands.HeaderLine = headerLine;
			if (footer != null) bands.FooterLine = footerLine;
			if (first.Header != null) bands.FirstHeaderLine = firstHeaderLine;
			if (first.Footer != null) bands.FirstFooterLine = firstFooterLine;
			if (Regex.Match(rtf, @"\\headery(\d+)") is { Success: true } hy) bands.HeaderDistance = Distance(int.Parse(hy.Groups[1].Value));
			if (Regex.Match(rtf, @"\\footery(\d+)") is { Success: true } fy) bands.FooterDistance = Distance(int.Parse(fy.Groups[1].Value));
		}
		return bands;
	}

	static string RtfUnescape(string s) {
		s = Regex.Replace(s, @"\\u(-?\d+)\?", m => ((char)(short)int.Parse(m.Groups[1].Value)).ToString());
		return Regex.Replace(s, @"\\([\\{}])", "$1");
	}

	/// RTF without its header and footer groups (and AhkWord's copy), which RichEdit has no place for.
	public static string StripFromRtf(string rtf) {
		foreach (string name in new[] { "header", "headerl", "headerr", "headerf", "footer", "footerl", "footerr", "footerf", "*\\" + RtfDestination }) {
			while (true) {
				int start = FindGroup(rtf, name);
				if (start < 0) break;
				int end = GroupEnd(rtf, start);
				if (end < 0) break;
				rtf = rtf.Remove(start, end - start + 1);
			}
		}
		return rtf;
	}

	// The position of "{\name" (not a longer control word), or -1.
	static int FindGroup(string rtf, string name) {
		string open = "{\\" + name;
		for (int i = rtf.IndexOf(open, StringComparison.Ordinal); i >= 0; i = rtf.IndexOf(open, i + 1, StringComparison.Ordinal)) {
			int after = i + open.Length;
			if (after >= rtf.Length || !char.IsLetter(rtf[after])) {
				if (i > 0 && rtf[i - 1] == '\\') continue;    // an escaped brace
				return i;
			}
		}
		return -1;
	}

	static int GroupEnd(string rtf, int start) {
		int depth = 0;
		for (int i = start; i < rtf.Length; i++) {
			char c = rtf[i];
			if (c == '\\') { i++; continue; }
			if (c == '{') depth++;
			else if (c == '}' && --depth == 0) return i;
		}
		return -1;
	}

	static readonly Regex RtfToken = new(@"\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z])|([{}])|([^\\{}\r\n]+)|[\r\n]+");
	static readonly Encoding Latin = Windows1252();

	static Encoding Windows1252() {
		try {
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
			return Encoding.GetEncoding(1252);
		} catch (Exception) {
			return Encoding.Latin1;
		}
	}

	// The text of an RTF group (tabs, paragraph ends as \n, fields as {page} and so on), and the alignment of its first
	// paragraph (0 left, 1 center, 2 right).
	static (string Text, int Align) RtfGroupText(string rtf, int start) {
		int end = GroupEnd(rtf, start);
		var tokens = RtfToken.Matches(rtf[start..(end + 1)]).ToList();
		int align = -1, i = 0;
		string text = Group(tokens, ref i, ref align);
		return (text, Math.Max(0, align));
	}

	static readonly HashSet<string> Skipped = new() { "fonttbl", "colortbl", "stylesheet", "pict", "object", "info", "listtable", "listoverridetable", "shp", "nonshppict" };

	// Reads the group whose "{" is tokens[i]; returns its text and leaves i after its "}".
	static string Group(List<Match> tokens, ref int i, ref int align) {
		i++;    // "{"
		var sb = new StringBuilder();
		bool first = true;
		int skipChars = 0;
		while (i < tokens.Count) {
			var t = tokens[i];
			if (t.Groups[5].Success) {
				if (t.Value == "}") { i++; return sb.ToString(); }
				// a group inside: a field, a skipped destination, or more text
				var next = i + 1 < tokens.Count ? tokens[i + 1] : null;
				string word = next?.Groups[1].Value ?? "";
				bool starred = next?.Groups[4].Value == "*";
				if (word == "field") { sb.Append(FieldGroup(tokens, ref i, ref align)); continue; }
				if (starred || Skipped.Contains(word)) { SkipGroup(tokens, ref i); continue; }
				sb.Append(Group(tokens, ref i, ref align));
				continue;
			}
			i++;
			if (t.Groups[1].Success) {
				string word = t.Groups[1].Value;
				switch (word) {
					case "tab": sb.Append('\t'); break;
					case "par": case "sect": sb.Append('\n'); break;
					case "line": sb.Append(' '); break;
					case "u":
						sb.Append((char)(short)int.Parse(t.Groups[2].Value));
						skipChars = 1;
						continue;
					case "qc": if (align < 0) align = 1; break;
					case "qr": if (align < 0) align = 2; break;
					case "ql": case "qj": if (align < 0) align = 0; break;
				}
				if (first && Skipped.Contains(word)) { /* (handled by the caller) */ }
			} else if (t.Groups[3].Success) {
				if (skipChars > 0) { skipChars--; continue; }
				sb.Append(Latin.GetString(new[] { Convert.ToByte(t.Groups[3].Value, 16) }));
			} else if (t.Groups[4].Success) {
				string symbol = t.Groups[4].Value;
				if (symbol == "~") sb.Append('\u00A0');
				else if (symbol is "\\" or "{" or "}") sb.Append(symbol);
			} else if (t.Groups[6].Success) {
				string s = t.Value;
				if (skipChars > 0 && s.Length > 0) { s = s[1..]; skipChars = 0; }
				sb.Append(s);
			}
			first = false;
			skipChars = t.Groups[6].Success || t.Groups[3].Success ? 0 : skipChars;
		}
		return sb.ToString();
	}

	static void SkipGroup(List<Match> tokens, ref int i) {
		int depth = 0;
		for (; i < tokens.Count; i++) {
			if (tokens[i].Value == "{") depth++;
			else if (tokens[i].Value == "}" && --depth == 0) { i++; return; }
		}
	}

	// {\field{\*\fldinst PAGE}{\fldrslt 1}}: the field's {name}, or its result for other fields.
	static string FieldGroup(List<Match> tokens, ref int i, ref int align) {
		int depth = 0;
		string instruction = "", result = "";
		i++;    // "{"
		depth = 1;
		while (i < tokens.Count && depth > 0) {
			var t = tokens[i];
			if (t.Value == "{") {
				// {\*\fldinst ...} or {\fldrslt ...}
				int j = i + 1;
				if (j < tokens.Count && tokens[j].Groups[4].Value == "*") j++;
				string word = j < tokens.Count ? tokens[j].Groups[1].Value : "";
				if (word == "fldinst") {
					int a = -1;
					int k = i;
					// the instruction's text: the group read as text, without its "\*\fldinst"
					instruction = PlainGroup(tokens, ref k);
					i = k;
					continue;
				}
				if (word == "fldrslt") { result = Group(tokens, ref i, ref align); continue; }
				SkipGroup(tokens, ref i);
				continue;
			}
			if (t.Value == "}") { depth--; i++; continue; }
			i++;
		}
		return FieldOf(instruction, result);
	}

	// A group's plain text (all of it, starred or not).
	static string PlainGroup(List<Match> tokens, ref int i) {
		var sb = new StringBuilder();
		int depth = 0;
		for (; i < tokens.Count; i++) {
			var t = tokens[i];
			if (t.Value == "{") depth++;
			else if (t.Value == "}") { if (--depth == 0) { i++; break; } }
			else if (t.Groups[6].Success) sb.Append(t.Value);
			else if (t.Groups[3].Success) sb.Append(Latin.GetString(new[] { Convert.ToByte(t.Groups[3].Value, 16) }));
		}
		return sb.ToString();
	}

	// ================================================================ .docx

	/// Adds the header and footer to a .docx: header and footer parts and the page layout if it has no header or footer
	/// yet (the built-in converter's), and the exact copy in either case.
	public static void AddToDocx(string path, PageBands bands, PageLayout layout, string title, string file, bool parts) {
		using var package = WordprocessingDocument.Open(path, true);
		var main = package.MainDocumentPart!;
		var body = main.Document.Body!;
		if (parts && !main.HeaderParts.Any() && !main.FooterParts.Any()) {
			var section = body.Elements<W.SectionProperties>().LastOrDefault();
			if (section == null) { section = new W.SectionProperties(); body.Append(section); }
			foreach (var old in section.Elements<W.HeaderReference>().Cast<DocumentFormat.OpenXml.OpenXmlElement>().Concat(section.Elements<W.FooterReference>()).ToList()) old.Remove();
			DocumentFormat.OpenXml.OpenXmlElement? last = null;
			void Add(DocumentFormat.OpenXml.OpenXmlElement reference) => last = last == null ? section.PrependChild(reference) : section.InsertAfter(reference, last);
			void Header(string[]? parts, W.HeaderFooterValues type) {
				if (parts == null) return;
				int page = type == W.HeaderFooterValues.First ? 1 : 2;
				var part = main.AddNewPart<HeaderPart>();
				part.Header = new W.Header(DocxParagraphs(Filled(parts, title, file), layout, "Header", bands.HeaderFontOn(page), bands.HeaderLineOn(page)));
				part.Header.Save();
				Add(new W.HeaderReference { Type = type, Id = main.GetIdOfPart(part) });
			}
			void Footer(string[]? parts, W.HeaderFooterValues type) {
				if (parts == null) return;
				int page = type == W.HeaderFooterValues.First ? 1 : 2;
				var part = main.AddNewPart<FooterPart>();
				part.Footer = new W.Footer(DocxParagraphs(Filled(parts, title, file), layout, "Footer", bands.FooterFontOn(page), bands.FooterLineOn(page)));
				part.Footer.Save();
				Add(new W.FooterReference { Type = type, Id = main.GetIdOfPart(part) });
			}
			Header(bands.HeaderToWrite, W.HeaderFooterValues.Default);
			Footer(bands.FooterToWrite, W.HeaderFooterValues.Default);
			// (iteration 49) the first page's own (none there if it has none: a first page without one is blank)
			Header(bands.FirstHeaderToWrite, W.HeaderFooterValues.First);
			Footer(bands.FirstFooterToWrite, W.HeaderFooterValues.First);
			section.RemoveAllChildren<W.TitlePage>();
			if (bands.DifferentFirst) {
				var titlePage = new W.TitlePage();
				// (in its place in the section's properties: before these, if there are any)
				var after = section.ChildElements.FirstOrDefault(e => e is W.TextDirection or W.BiDi or W.GutterOnRight or W.DocGrid or W.PrinterSettingsReference);
				if (after != null) section.InsertBefore(titlePage, after); else section.Append(titlePage);
			}
			section.RemoveAllChildren<W.PageSize>();
			section.RemoveAllChildren<W.PageMargin>();
			var size = new W.PageSize { Width = (uint)layout.Width, Height = (uint)layout.Height };
			if (layout.Landscape) size.Orient = W.PageOrientationValues.Landscape;
			var margin = new W.PageMargin { Left = (uint)layout.Left, Right = (uint)layout.Right, Top = layout.Top, Bottom = layout.Bottom, Header = (uint)layout.HeaderY, Footer = (uint)layout.FooterY, Gutter = 0 };
			if (last == null) { section.PrependChild(margin); section.PrependChild(size); }
			else { section.InsertAfter(margin, last); section.InsertAfter(size, last); }
			main.Document.Save();
		}
		// the exact copy, in a custom XML part of its own (Word keeps those)
		foreach (var old in main.CustomXmlParts.Where(p => ReadCustomXml(p) != null).ToList()) main.DeletePart(old);
		var custom = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
		using (var stream = custom.GetStream(FileMode.Create, FileAccess.Write))
			new XDocument(new XElement(XName.Get("headerFooter", XmlNamespace), Stamped(bands, title, file).ToJson())).Save(stream);
	}

	static string? ReadCustomXml(CustomXmlPart part) {
		try {
			using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
			var root = XDocument.Load(stream).Root;
			return root?.Name == XName.Get("headerFooter", XmlNamespace) ? root.Value : null;
		} catch { return null; }
	}

	// (iteration 55) a paragraph per line
	static W.Paragraph[] DocxParagraphs(string[] parts, PageLayout layout, string style, BandFont? font, bool line = false) {
		var lines = PageBands.Parts(parts).Select(BandText.Lines).ToArray();
		var paragraphs = Enumerable.Range(0, lines.Max(l => l.Count)).Select(n => DocxParagraph(lines.Select(l => n < l.Count ? l[n] : new()).ToArray(), layout, style, font)).ToArray();
		// (iteration 55) the line below the header, above the footer
		if (line)
			foreach (var paragraph in paragraphs) {
				var border = new W.ParagraphBorders();
				if (style == "Header") border.Append(new W.BottomBorder { Val = W.BorderValues.Single, Size = 6, Space = 2 });
				else border.Append(new W.TopBorder { Val = W.BorderValues.Single, Size = 6, Space = 2 });
				paragraph.ParagraphProperties!.InsertAt(border, 0);
			}
		return paragraphs;
	}

	// A line: each part's runs (iteration 55: in the band's font, or a word's own)
	static W.Paragraph DocxParagraph(List<(string Text, RunFormat? Format)>[] cells, PageLayout layout, string style, BandFont? band) {
		BandFont? font = band;
		W.Run Run(DocumentFormat.OpenXml.OpenXmlElement content) {
			var run = new W.Run();
			if (font != null) {
				var properties = new W.RunProperties(new W.RunFonts { Ascii = font.Name, HighAnsi = font.Name, ComplexScript = font.Name, EastAsia = font.Name });
				if (font.Bold) properties.Append(new W.Bold());
				if (font.Italic) properties.Append(new W.Italic());
				if (font.Underline) properties.Append(new W.Underline { Val = W.UnderlineValues.Single });
				string halfPoints = ((int)Math.Round(font.Size * 2)).ToString(System.Globalization.CultureInfo.InvariantCulture);
				properties.Append(new W.FontSize { Val = halfPoints }, new W.FontSizeComplexScript { Val = halfPoints });
				run.Append(properties);
			}
			run.Append(content);
			return run;
		}
		var paragraph = new W.Paragraph(new W.ParagraphProperties(
			new W.Tabs(new W.TabStop { Val = W.TabStopValues.Center, Position = layout.TextWidth / 2 }, new W.TabStop { Val = W.TabStopValues.Right, Position = layout.TextWidth })));
		for (int p = 0; p < 3; p++) {
			font = band;
			if (p > 0) paragraph.Append(Run(new W.TabChar()));
			foreach (var (part, format) in cells[p]) {
			font = format == null ? band : BandText.Resolve(band ?? BandFont.Plain, format);
			int at = 0;
			foreach (Match m in Field.Matches(part)) {
				if (m.Index > at) paragraph.Append(Run(new W.Text(part[at..m.Index]) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
				string instruction = m.Groups[1].Value.ToLowerInvariant() switch { "page" => "PAGE", "pages" => "NUMPAGES", "sectionpages" => "SECTIONPAGES", "date" => "DATE", _ => "TIME" };
				paragraph.Append(new W.SimpleField(Run(new W.Text("1"))) { Instruction = " " + instruction + " " });
				at = m.Index + m.Length;
			}
			if (at < part.Length) paragraph.Append(Run(new W.Text(part[at..]) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
			}
		}
		return paragraph;
	}

	/// The header and footer of a .docx (null if it has neither).
	public static PageBands? FromDocx(string path) {
		try {
			using var package = WordprocessingDocument.Open(path, false);
			var main = package.MainDocumentPart!;
			PageBands? exact = null;
			foreach (var custom in main.CustomXmlParts)
				if (ReadCustomXml(custom) is { } json) exact = PageBands.FromJson(json);
			// (iteration 63) the first section's: the first w:sectPr (the body's own is the last section's, a paragraph's
			// those before it)
			return OfSection(main, exact, main.Document.Body?.Descendants<W.SectionProperties>().FirstOrDefault());
		} catch (Exception ex) when (ex is IOException or InvalidDataException or DocumentFormat.OpenXml.Packaging.OpenXmlPackageException) {
			return null;
		}
	}

	/// The sections after a .docx's first, in order, as section breaks: a section with no header or footer of its own keeps
	/// the one before's, as Word links them; w:pgNumType's start restarts its page numbers. Null for a section that
	/// doesn't start a new page (continuous). Empty if the file can't be read.
	public static List<SectionSettings?> SectionsFromDocx(string path) {
		var result = new List<SectionSettings?>();
		try {
			using var package = WordprocessingDocument.Open(path, false);
			var main = package.MainDocumentPart!;
			var sections = main.Document.Body?.Descendants<W.SectionProperties>().ToList() ?? new();
			if (sections.Count < 2) return result;
			var current = OfSection(main, null, sections[0]) ?? new PageBands();
			bool own = false;
			static bool Has<TRef>(W.SectionProperties s) where TRef : W.HeaderFooterReferenceType => s.Elements<TRef>().Any(r => r.Type == null || r.Type.Value == W.HeaderFooterValues.Default);
			foreach (var section in sections.Skip(1)) {
				bool header = Has<W.HeaderReference>(section), footer = Has<W.FooterReference>(section);
				if (header || footer) {
					var bands = OfSection(main, null, section) ?? new PageBands();
					if (!header) (bands.ShowHeader, bands.Header, bands.HeaderLine) = (current.ShowHeader, current.Header, current.HeaderLine);
					if (!footer) (bands.ShowFooter, bands.Footer, bands.FooterLine) = (current.ShowFooter, current.Footer, current.FooterLine);
					current = bands;
					own = true;
				}
				if (section.GetFirstChild<W.SectionType>()?.Val?.Value == W.SectionMarkValues.Continuous) { result.Add(null); continue; }
				uint? start = section.GetFirstChild<W.PageNumberType>()?.Start?.Value is int number ? (uint)number : null;
				result.Add(new SectionSettings { Same = !own, Bands = own ? Sections.CopyBands(current) : null, Restart = start != null, Start = (int)(start ?? 1) });
				own = false;
			}
		} catch (Exception ex) when (ex is IOException or InvalidDataException or DocumentFormat.OpenXml.Packaging.OpenXmlPackageException) { result.Clear(); }
		return result;
	}

	// a section's header and footer in a .docx (exact: AhkWord's own copy of them, if any)
	static PageBands? OfSection(MainDocumentPart main, PageBands? exact, W.SectionProperties? section) {
			string[]? Read<TRef>(Func<string, OpenXmlPart?> part, bool firstPage = false) where TRef : W.HeaderFooterReferenceType {
				var references = section?.Elements<TRef>().ToList() ?? new List<TRef>();
				var reference = firstPage ? references.FirstOrDefault(r => r.Type?.Value == W.HeaderFooterValues.First)
					: references.FirstOrDefault(r => r.Type == null || r.Type.Value == W.HeaderFooterValues.Default) ?? references.FirstOrDefault(r => r.Type?.Value != W.HeaderFooterValues.First);
				if (reference?.Id?.Value is not string id || part(id) is not { } found) return null;
				var root = found.RootElement;
				if (root == null) return null;
				// (iteration 55) a line per paragraph
				return Combined(root.Descendants<W.Paragraph>().Select(paragraph => { var (line, align) = DocxText(paragraph); return PartsOf(line, align); }));
			}
			bool haveReal = section != null && (section.Elements<W.HeaderReference>().Any() || section.Elements<W.FooterReference>().Any());
			var header = Read<W.HeaderReference>(id => main.GetPartById(id));
			var footer = Read<W.FooterReference>(id => main.GetPartById(id));
			// (iteration 49) the first page's own: w:titlePg on (no w:val, or a true one)
			var first = default(FirstPage);
			if (section?.GetFirstChild<W.TitlePage>() is { } titlePage && (titlePage.Val == null || titlePage.Val.Value))
				first = new FirstPage(true, Read<W.HeaderReference>(id => main.GetPartById(id), true), Read<W.FooterReference>(id => main.GetPartById(id), true));
			var bands = Decide(exact, header, footer, haveReal, first);
			// (iteration 55) the lines: a border below the header's paragraphs, above the footer's
			if (bands != null) {
				bool Bordered<TRef>(bool top) where TRef : W.HeaderFooterReferenceType =>
					section?.Elements<TRef>().FirstOrDefault(r => r.Type == null || r.Type.Value == W.HeaderFooterValues.Default)?.Id?.Value is string id
					&& main.GetPartById(id).RootElement is { } root
					&& root.Descendants<W.ParagraphBorders>().Any(b => top ? b.TopBorder?.Val?.Value is { } t && t != W.BorderValues.None && t != W.BorderValues.Nil : b.BottomBorder?.Val?.Value is { } v && v != W.BorderValues.None && v != W.BorderValues.Nil);
				if (header != null) bands.HeaderLine = Bordered<W.HeaderReference>(false);
				if (footer != null) bands.FooterLine = Bordered<W.FooterReference>(true);
			}
			// (iteration 55) the distances from the edges: the file's own
			if (bands != null && section?.GetFirstChild<W.PageMargin>() is { } margin) {
				if (margin.Header?.Value is uint h) bands.HeaderDistance = Distance((int)h);
				if (margin.Footer?.Value is uint f) bands.FooterDistance = Distance((int)f);
			}
			return bands;
	}

	// A paragraph's text, with tabs and fields ({page} ...), and its alignment.
	static (string Text, int Align) DocxText(W.Paragraph paragraph) {
		var sb = new StringBuilder();
		string? instruction = null;    // in a complex field: its instruction so far
		bool inResult = false;
		string result = "";
		foreach (var element in paragraph.Descendants()) {
			switch (element) {
				case W.SimpleField simple:
					sb.Append(FieldOf(simple.Instruction?.Value ?? "", simple.InnerText));
					break;
				case W.FieldChar fieldChar when fieldChar.FieldCharType?.Value == W.FieldCharValues.Begin:
					instruction = ""; inResult = false; result = "";
					break;
				case W.FieldChar fieldChar when fieldChar.FieldCharType?.Value == W.FieldCharValues.Separate:
					inResult = true;
					break;
				case W.FieldChar fieldChar when fieldChar.FieldCharType?.Value == W.FieldCharValues.End:
					if (instruction != null) sb.Append(FieldOf(instruction, result));
					instruction = null; inResult = false;
					break;
				case W.FieldCode code:
					if (instruction != null) instruction += code.Text;
					break;
				case W.Text text when text.Ancestors<W.SimpleField>().Any():
					break;
				case W.Text text:
					if (instruction != null) { if (inResult) result += text.Text; }
					else sb.Append(text.Text);
					break;
				case W.TabChar when element.Parent is W.Run:
					if (instruction == null) sb.Append('\t');
					break;
				case W.PositionalTab:
					sb.Append('\t');
					break;
			}
		}
		var justification = paragraph.ParagraphProperties?.Justification?.Val?.Value;
		int align = justification == W.JustificationValues.Center ? 1 : justification == W.JustificationValues.Right || justification == W.JustificationValues.End ? 2 : 0;
		return (sb.ToString(), align);
	}

	// ================================================================ .odt

	static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", Meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
		Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0", Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

	/// Adds the exact copy to an .odt (its real header and footer come from the RTF LibreOffice or Word converted).
	public static void AddToOdt(string path, PageBands bands, string title, string file) {
		using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
		var entry = zip.GetEntry("meta.xml");
		XDocument document;
		if (entry != null) { using (var stream = entry.Open()) document = XDocument.Load(stream); entry.Delete(); }
		else document = new XDocument(new XElement(Office + "document-meta", new XAttribute(XNamespace.Xmlns + "office", Office), new XAttribute(XNamespace.Xmlns + "meta", Meta),
			new XAttribute(Office + "version", "1.3"), new XElement(Office + "meta")));
		var meta = document.Root!.Element(Office + "meta");
		if (meta == null) { meta = new XElement(Office + "meta"); document.Root.Add(meta); }
		meta.Elements(Meta + "user-defined").Where(e => (string?)e.Attribute(Meta + "name") == OdtProperty).Remove();
		meta.Add(new XElement(Meta + "user-defined", new XAttribute(Meta + "name", OdtProperty), Stamped(bands, title, file).ToJson()));
		var added = zip.CreateEntry("meta.xml", CompressionLevel.Optimal);
		using (var stream = added.Open()) document.Save(stream, SaveOptions.DisableFormatting);
	}

	/// The header and footer of an .odt (null if it has neither): from its first master page.
	public static PageBands? FromOdt(string path) {
		try {
			using var zip = ZipFile.OpenRead(path);
			PageBands? exact = null;
			if (zip.GetEntry("meta.xml") is { } metaEntry) {
				using var stream = metaEntry.Open();
				var value = XDocument.Load(stream).Descendants(Meta + "user-defined").FirstOrDefault(e => (string?)e.Attribute(Meta + "name") == OdtProperty)?.Value;
				if (value != null) exact = PageBands.FromJson(value);
			}
			if (zip.GetEntry("styles.xml") is not { } stylesEntry) return Decide(exact, null, null, false);
			XDocument styles;
			using (var stream = stylesEntry.Open()) styles = XDocument.Load(stream);
			var page = styles.Descendants(Style + "master-page").FirstOrDefault(m => (string?)m.Attribute(Style + "name") == "Standard")
				?? styles.Descendants(Style + "master-page").FirstOrDefault();
			var paragraphStyles = styles.Descendants(Style + "style").Where(s => s.Attribute(Style + "name") != null)
				.GroupBy(s => (string)s.Attribute(Style + "name")!).ToDictionary(g => g.Key, g => g.First());
			string[]? Read(string band) {
				var element = page?.Element(Style + band);
				if (element == null || (string?)element.Attribute(Style + "display") == "false") return null;
				// (iteration 55) a line per paragraph
				return Combined(element.Descendants(Text + "p").Select(paragraph => {
					string styleName = (string?)paragraph.Attribute(Text + "style-name") ?? "";
					string? align = paragraphStyles.TryGetValue(styleName, out var s) ? (string?)s.Element(Style + "paragraph-properties")?.Attribute(Fo + "text-align") : null;
					return PartsOf(OdtText(paragraph), align is "center" ? 1 : align is "end" or "right" ? 2 : 0);
				}));
			}
			bool haveReal = page?.Element(Style + "header") != null || page?.Element(Style + "footer") != null;
			var (header, footer) = (Read("header"), Read("footer"));
			// (iteration 49) the first page's own: style:header-first / footer-first (LibreOffice's; without one, the first
			// page has the others' header or footer)
			var first = default(FirstPage);
			bool firstHeader = page?.Element(Style + "header-first") != null, firstFooter = page?.Element(Style + "footer-first") != null;
			if (firstHeader || firstFooter) {
				first = new FirstPage(true, firstHeader ? Read("header-first") : header, firstFooter ? Read("footer-first") : footer);
				haveReal = true;
			}
			return Decide(exact, header, footer, haveReal, first);
		} catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException) {
			return null;
		}
	}

	static string OdtText(XElement element) {
		var sb = new StringBuilder();
		foreach (var node in element.Nodes()) {
			if (node is XText text) { sb.Append(text.Value); continue; }
			if (node is not XElement e) continue;
			if (e.Name.Namespace != Text) { sb.Append(OdtText(e)); continue; }
			switch (e.Name.LocalName) {
				case "tab": sb.Append('\t'); break;
				case "s": sb.Append(' ', (int?)e.Attribute(Text + "c") ?? 1); break;
				case "line-break": sb.Append(' '); break;
				case "page-number": sb.Append("{page}"); break;
				case "page-count": sb.Append("{pages}"); break;
				case "date": sb.Append("{date}"); break;
				case "time": sb.Append("{time}"); break;
				case "file-name": sb.Append((string?)e.Attribute(Text + "display") == "full" ? "{file}" : "{title}"); break;
				default: sb.Append(OdtText(e)); break;
			}
		}
		return sb.ToString();
	}

	// ================================================================ HTML and Markdown

	public static string HtmlMeta(PageBands bands) =>
		$"<meta name=\"{MetaName}\" content=\"{System.Net.WebUtility.HtmlEncode(bands.Copy().ToJson())}\">\r\n";

	public static PageBands? FromHtml(string html) {
		var m = Regex.Match(html, "<meta\\s+name=\"" + MetaName + "\"\\s+content=\"([^\"]*)\"", RegexOptions.IgnoreCase);
		return m.Success ? PageBands.FromJson(System.Net.WebUtility.HtmlDecode(m.Groups[1].Value)) : null;
	}

	public static string MarkdownLine(PageBands bands) => $"<!-- {MetaName}: {bands.Copy().ToJson().Replace("--", "\\u002D\\u002D")} -->";

	public static PageBands? FromMarkdown(string markdown) {
		var m = Regex.Match(markdown, @"^<!-- " + MetaName + @": (.*) -->\s*$", RegexOptions.Multiline);
		return m.Success ? PageBands.FromJson(m.Groups[1].Value) : null;
	}
}
