// Footnotes, for AhkWord (AhkWord.cs, through #load): a superscript number in the text, its note at the foot of the
// page (printed, Print Preview, PDF, and on Show Pages' sheets).
//
// RichEdit has no footnotes, so a footnote is hidden text (CFE_HIDDEN, RTF \v) around its number, as cross-references
// are (rte_crossrefs.cs), the note's text in it:
//   [[fn]]1[[fn:The note's text.]]   "[[fn]]" and the note hidden, "1" in superscript
// So a note goes wherever its number goes (typing, cutting and pasting, undo, RichEdit's own RTF). The note is text of
// one or more paragraphs, with bold, italic and underline (marks: private-use characters; "]]" in it is kept as "] ]"). Footnotes are numbered 1, 2, 3 ... in the document's order
// (Update: after inserting and deleting one, before saving and printing, F9).
// Files: RTF (and so Word and LibreOffice files) gets real footnotes ({\super\chftn}{\footnote ...}), and Word's and
// LibreOffice's come back as footnotes; the built-in .docx converter gets real footnotes too (DocxFootnotes); HTML
// gets numbered links to a list of the notes at its end, Markdown its footnotes ([^1] and [^1]: note).
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET.TextObjectModel2;
using W = DocumentFormat.OpenXml.Wordprocessing;

public static class Footnotes
{
	public const string Start = "[[fn]]", NoteOpen = "[[fn:", NoteClose = "]]";
	/// A footnote in the text: its number (group 1) and its note (group 2, escaped).
	public static readonly Regex Note = new(@"\[\[fn\]\]([^\[\]\r]{0,12}?)\[\[fn:((?:[^\]\r]|\](?!\]))*)\]\]");

	/// A footnote: where it is in the text (its markers and number), its number, and its note.
	public sealed record Footnote(int Start, int End, string Number, string Text, int NumberStart, int NumberEnd);

	/// A note's paragraphs: a private-use character between them in its text (as its formatting's marks), the formats'
	/// own in each file (RTF \par, .docx paragraphs, HTML <p>, Markdown's indented paragraphs) and "\n" for scripts.
	public const char ParagraphMark = '';

	/// The note's text as kept in its marker (its line breaks, "\r\n", "\n" or "\r", paragraph marks; none at its ends;
	/// "]]" and "[[" apart), and back.
	public static string Escape(string text) {
		text = Regex.Replace(text, @"\r\n|[\r\n\v ]", ParagraphMark.ToString());
		text = Regex.Replace(text, @"[ \t]*[ \t]*", ParagraphMark.ToString());
		text = Regex.Replace(text, @"^[\s]+|[\s]+$", "");
		return text.Replace("]]", "] ]").Replace("[[", "[ [");
	}
	public static string Unescape(string text) => text;

	public static bool Any(string story) => story.Contains("[[fn]]", StringComparison.Ordinal);

	// A note's bold, italic and underline: private-use characters in its text, on and off for each (they come in no
	// text of their own), made the formats' own in each file (RTF \b, HTML <b>, Markdown **, .docx runs) and back.
	public const char BoldOn = '', BoldOff = '', ItalicOn = '', ItalicOff = '', UnderlineOn = '', UnderlineOff = '';
	static readonly Regex Marks = new("[-]");

	/// A note's text without its formatting, its paragraphs apart by "\n".
	public static string Plain(string note) => Marks.Replace(note, "").Replace(ParagraphMark, '\n');

	/// A note's paragraphs apart by "\n" (scripts), or another separator.
	public static string Paragraphs(string note, string separator = "\n") => note.Replace(ParagraphMark.ToString(), separator);

	// A note's marks nested (what was switched on last switched off first), each paragraph's its own (a format on
	// across a paragraph's end switched off at it and on again after it, as Word's runs are each paragraph's).
	static string Nested(string note) => note.IndexOfAny(new[] { BoldOn, ItalicOn, UnderlineOn }) < 0 ? note : string.Join(ParagraphMark.ToString(), SplitParagraphs(note));

	/// A note with <b> <i> <u> tags (scripts, HTML) for its formatting, and back (<strong> and <em> too).
	public static string ToTags(string note) => Nested(note).Replace(BoldOn.ToString(), "<b>").Replace(BoldOff.ToString(), "</b>")
		.Replace(ItalicOn.ToString(), "<i>").Replace(ItalicOff.ToString(), "</i>").Replace(UnderlineOn.ToString(), "<u>").Replace(UnderlineOff.ToString(), "</u>");
	public static string FromTags(string text) => Regex.Replace(text, @"<(/?)(b|strong|i|em|u)\s*>", m => {
		bool off = m.Groups[1].Value == "/";
		return m.Groups[2].Value.ToLowerInvariant() switch {
			"b" or "strong" => (off ? BoldOff : BoldOn).ToString(),
			"i" or "em" => (off ? ItalicOff : ItalicOn).ToString(),
			_ => (off ? UnderlineOff : UnderlineOn).ToString()
		};
	}, RegexOptions.IgnoreCase);

	/// A note's text as runs of one style each.
	public static List<(string Text, FontStyle Style)> Runs(string note) {
		var runs = new List<(string, FontStyle)>();
		var style = FontStyle.Regular;
		var sb = new StringBuilder();
		void Flush() { if (sb.Length > 0) { runs.Add((sb.ToString(), style)); sb.Clear(); } }
		foreach (char c in note) {
			FontStyle? now = c switch {
				BoldOn => style | FontStyle.Bold, BoldOff => style & ~FontStyle.Bold,
				ItalicOn => style | FontStyle.Italic, ItalicOff => style & ~FontStyle.Italic,
				UnderlineOn => style | FontStyle.Underline, UnderlineOff => style & ~FontStyle.Underline,
				_ => null
			};
			if (now is { } changed) { if (changed != style) { Flush(); style = changed; } continue; }
			sb.Append(c);
		}
		Flush();
		return runs;
	}

	/// Runs as a note's text with its marks (each run's style switched on and off around it), nested: what was switched
	/// on last is switched off first (as tags have to be), switched off and on again if need be.
	public static string FromRuns(IEnumerable<(string Text, FontStyle Style)> runs) {
		var sb = new StringBuilder();
		var open = new List<FontStyle>();
		var formats = new[] { (FontStyle.Bold, BoldOn, BoldOff), (FontStyle.Italic, ItalicOn, ItalicOff), (FontStyle.Underline, UnderlineOn, UnderlineOff) };
		void Switch(FontStyle to) {
			int keep = open.FindIndex(flag => (to & flag) == 0);
			if (keep >= 0)
				while (open.Count > keep) { var flag = open[^1]; sb.Append(formats.First(f => f.Item1 == flag).Item3); open.RemoveAt(open.Count - 1); }
			foreach (var (flag, on, _) in formats)
				if ((to & flag) != 0 && !open.Contains(flag)) { sb.Append(on); open.Add(flag); }
		}
		foreach (var (text, runStyle) in runs) {
			if (text.Length == 0) continue;
			Switch(runStyle & (FontStyle.Bold | FontStyle.Italic | FontStyle.Underline));
			sb.Append(text);
		}
		Switch(FontStyle.Regular);
		return sb.ToString();
	}

	/// The document's footnotes, in order (story: the document's text).
	public static List<Footnote> All(string story) {
		var list = new List<Footnote>();
		if (!Any(story)) return list;
		foreach (Match m in Note.Matches(story))
			list.Add(new Footnote(m.Index, m.Index + m.Length, m.Groups[1].Value, Unescape(m.Groups[2].Value), m.Groups[1].Index, m.Groups[1].Index + m.Groups[1].Length));
		return list;
	}

	/// The footnotes whose numbers are in the text from..to (a page's).
	public static List<Footnote> In(List<Footnote> all, int from, int to) => all.Where(f => f.NumberStart >= from && f.NumberStart < to).ToList();

	static string Story(ITextDocument2 doc) => doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";

	// Puts text at a position, hidden or (not hidden) superscript; returns where it ends.
	static int Put(ITextDocument2 doc, int at, string text, bool hidden) {
		var range = doc.Range2(at, at);
		range.Text = text;
		range.SetRange(at, at + text.Length);
		range.Font.Hidden = hidden ? tomConstants.tomTrue : tomConstants.tomFalse;
		if (!hidden) range.Font.Superscript = tomConstants.tomTrue;
		return at + text.Length;
	}

	/// Inserts a footnote at a range (in place of its text) with its note; returns where it ends. The numbers are
	/// worked out again (Update).
	public static int Insert(ITextDocument2 doc, ITextRange2 at, string note) {
		if (at.End > at.Start) at.Text = "";
		int start = at.Start;
		int end = Put(doc, start, Start, hidden: true);
		end = Put(doc, end, "0", hidden: false);
		end = Put(doc, end, NoteOpen + Escape(note) + NoteClose, hidden: true);
		var tracked = doc.Range2(end, end);
		Update(doc);
		return tracked.Start;
	}

	/// Gives a footnote another note.
	public static void SetText(ITextDocument2 doc, Footnote footnote, string note) {
		int at = footnote.NumberEnd;
		var range = doc.Range2(at, footnote.End);
		range.Text = NoteOpen + Escape(note) + NoteClose;
		range.SetRange(at, at + range.Text!.Length);
		range.Font.Hidden = tomConstants.tomTrue;
	}

	/// Numbers the footnotes 1, 2, 3 ... in order. Returns how many changed.
	public static int Update(ITextDocument2 doc) {
		var all = All(Story(doc));
		int changed = 0;
		for (int i = all.Count - 1; i >= 0; i--) {
			string wanted = (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
			var f = all[i];
			if (f.Number == wanted) continue;
			var number = doc.Range2(f.NumberStart, f.NumberEnd);
			number.Text = wanted;
			number.SetRange(f.NumberStart, f.NumberStart + wanted.Length);
			number.Font.Hidden = tomConstants.tomFalse;
			number.Font.Superscript = tomConstants.tomTrue;
			changed++;
		}
		return changed;
	}

	/// The footnote at a position (on its number or its markers), or null.
	public static Footnote? At(string story, int position) => All(story).FirstOrDefault(f => position >= f.Start && position <= f.End);

	/// After loading (HTML, Markdown, Word's RTF): the markers hidden, the numbers superscript, and numbered again.
	/// Returns how many footnotes there are.
	public static int AfterLoad(ITextDocument2 doc) {
		var all = All(Story(doc));
		foreach (var f in all) {
			var start = doc.Range2(f.Start, f.NumberStart);
			if (start.Font.Hidden != tomConstants.tomTrue) start.Font.Hidden = tomConstants.tomTrue;
			var note = doc.Range2(f.NumberEnd, f.End);
			if (note.Font.Hidden != tomConstants.tomTrue) note.Font.Hidden = tomConstants.tomTrue;
			var number = doc.Range2(f.NumberStart, f.NumberEnd);
			if (f.NumberEnd > f.NumberStart && number.Font.Superscript != tomConstants.tomTrue) number.Font.Superscript = tomConstants.tomTrue;
		}
		if (all.Count > 0) Update(doc);
		return all.Count;
	}

	/// The text without the notes and markers, the numbers kept ("text1").
	public static string PlainText(string text) => Any(text) ? Note.Replace(text, m => m.Groups[1].Value) : text;

	// ================================================================ the notes at the foot of a page

	/// The style of the document's notes (as Word's): its size, and its font ("": the text's). Without it, NoteSize.
	public const string StyleName = "Footnote Text";

	/// The notes' size for text of a size, where the document has no Footnote Text style: three quarters of it, to the
	/// half point (12 points: 9), at least 7.
	public static float NoteSize(float textSize) => Math.Max(7f, MathF.Round(textSize * 0.75f * 2f, MidpointRounding.AwayFromZero) / 2f);

	/// The font of the document's notes: the Footnote Text style's (style; its font "" or none: the text's), else its
	/// first footnote's text's (the number's paragraph) at NoteSize; the default's where there is no footnote.
	public static (string Name, float Size) NoteFontOf(ITextDocument2 doc, List<Footnote> all, string defaultName, float defaultSize, (string Font, float Size)? style = null) {
		string name = defaultName;
		float size = defaultSize;
		if (all.Count > 0) {
			int at = Math.Max(0, all[0].Start - 1);
			var font = doc.Range2(at, at + 1).Font;
			if (font.Name is { Length: > 0 } n) name = n;
			if (font.Size > 1) size = font.Size;
		}
		if (style is { Size: > 0 } own) return (own.Font is { Length: > 0 } f ? f : name, own.Size);
		return (name, NoteSize(size));
	}

	/// The notes' size a file just loaded had (its first footnote's text's, in points; null: none, or not known): set
	/// by FromWordRtf and MarkDocx, taken (and forgotten) by the editor once the file is loaded.
	public static float? LoadedSize;

	public static Font NoteFont(string name, float size) => new(name, Math.Max(1f, size), GraphicsUnit.Point);

	/// A line of the notes at the foot of a page: a note's first line has its number; Parts: its text in runs of a style
	/// (bold, italic, underline).
	public sealed record NoteLine(string? Number, string Text, List<(string Text, FontStyle Style)> Parts);

	// The note's font in a style (made once a measuring or drawing, disposed by its caller).
	static Font Styled(Dictionary<FontStyle, Font> fonts, Font font, FontStyle style) {
		if (style == FontStyle.Regular) return font;
		if (!fonts.TryGetValue(style, out var styled)) fonts[style] = styled = new Font(font.FontFamily, font.Size, font.Style | style, font.Unit);
		return styled;
	}

	/// The notes as lines of a width (in g's units), each word-wrapped after its number's hanging room; each of a note's
	/// paragraphs starts a line (an empty one is an empty line, as in Word).
	public static List<NoteLine> Lines(Graphics g, List<Footnote> notes, float width, Font font) {
		var lines = new List<NoteLine>();
		float room = Math.Max(10, width - NumberWidth(g, font));
		using var format = new StringFormat(StringFormat.GenericTypographic) { FormatFlags = StringFormatFlags.MeasureTrailingSpaces };
		var fonts = new Dictionary<FontStyle, Font>();
		try {
			float Width(string text, FontStyle style) => g.MeasureString(text, Styled(fonts, font, style), int.MaxValue, format).Width;
			float space = Width(" ", FontStyle.Regular);
			foreach (var note in notes) {
				// the paragraphs' runs
				var paragraphs = new List<List<(string Text, FontStyle Style)>> { new() };
				foreach (var (text, style) in Runs(note.Text)) {
					var pieces = text.Split(ParagraphMark);
					for (int k = 0; k < pieces.Length; k++) {
						if (k > 0) paragraphs.Add(new());
						if (pieces[k] != "") paragraphs[^1].Add((pieces[k], style));
					}
				}
				string? number = note.Number;
				foreach (var paragraph in paragraphs) {
					// the words, each in runs of a style (a word can change style in the middle)
					var words = new List<List<(string Text, FontStyle Style)>>();
					var word = new List<(string Text, FontStyle Style)>();
					foreach (var (text, style) in paragraph) {
						var pieces = text.Split(' ');
						for (int k = 0; k < pieces.Length; k++) {
							if (k > 0 && word.Count > 0) { words.Add(word); word = new(); }
							if (pieces[k] != "") word.Add((pieces[k], style));
						}
					}
					if (word.Count > 0) words.Add(word);
					var parts = new List<(string Text, FontStyle Style)>();
					float used = 0;
					void Emit() {
						var merged = new List<(string Text, FontStyle Style)>();
						foreach (var part in parts)
							if (merged.Count > 0 && merged[^1].Style == part.Style) merged[^1] = (merged[^1].Text + part.Text, part.Style);
							else merged.Add(part);
						lines.Add(new NoteLine(number, string.Concat(merged.Select(m => m.Text)), merged));
						number = null;
						parts = new();
						used = 0;
					}
					foreach (var w in words) {
						float wordWidth = w.Sum(part => Width(part.Text, part.Style));
						if (parts.Count > 0 && used + space + wordWidth > room) Emit();
						if (parts.Count > 0) { parts.Add((" ", parts[^1].Style & w[0].Style)); used += space; }
						parts.AddRange(w);
						used += wordWidth;
					}
					Emit();
				}
			}
		} finally {
			foreach (var f in fonts.Values) f.Dispose();
		}
		return lines;
	}

	/// How tall lines of notes are at the foot of a page (in g's units), with the separator line above them.
	public static float Height(Graphics g, IReadOnlyList<NoteLine> lines, Font font) {
		if (lines.Count == 0) return 0;
		float line = font.GetHeight(g);
		return line * 0.9f + lines.Count * line + lines.Skip(1).Count(l => l.Number != null) * line * 0.1f;
	}

	/// How many of the lines fit in height (in g's units), with the separator.
	public static int Fitting(Graphics g, IReadOnlyList<NoteLine> lines, Font font, float height) {
		int count = 0;
		while (count < lines.Count && Height(g, lines.Take(count + 1).ToList(), font) <= height) count++;
		return count;
	}

	static float NumberWidth(Graphics g, Font font) => g.MeasureString("000", font).Width;

	/// Draws lines of notes in area (from its top: the separator, then the lines, each note's number hanging), in g's
	/// units. A line is drawn as it is (not wrapped again: the lines were made for the printed page).
	public static void Draw(Graphics g, IReadOnlyList<NoteLine> lines, RectangleF area, Font font, Color color) {
		if (lines.Count == 0) return;
		float line = font.GetHeight(g);
		float y = area.Top + line * 0.35f;
		using (var pen = new Pen(color, Math.Max(1f, line / 22f))) g.DrawLine(pen, area.Left, y, area.Left + Math.Min(area.Width, area.Width / 3), y);
		y = area.Top + line * 0.9f;
		float hang = NumberWidth(g, font);
		using var brush = new SolidBrush(color);
		using var small = new Font(font.FontFamily, font.SizeInPoints * 0.75f, font.Style, GraphicsUnit.Point);
		using var format = new StringFormat(StringFormat.GenericTypographic) { FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.MeasureTrailingSpaces, Trimming = StringTrimming.None };
		var fonts = new Dictionary<FontStyle, Font>();
		try {
			for (int i = 0; i < lines.Count; i++) {
				if (i > 0 && lines[i].Number != null) y += line * 0.1f;
				if (lines[i].Number is { } number) g.DrawString(number, small, brush, area.Left, y);
				// each run in its style, one after another; the underline drawn here, one line under the underlined runs next
				// to each other (drawn by the font, a run of a space between two underlined words had none)
				float x = area.Left + hang;
				float? underlineFrom = null;
				void Underline(float to) {
					if (underlineFrom is not { } from) return;
					float ascent = line * font.FontFamily.GetCellAscent(font.Style) / font.FontFamily.GetLineSpacing(font.Style);
					float thickness = Math.Max(1f, line / 18f);
					using var pen = new Pen(color, thickness);
					float under = y + ascent + thickness * 1.5f;
					g.DrawLine(pen, from, under, to, under);
					underlineFrom = null;
				}
				foreach (var (text, style) in lines[i].Parts) {
					var styled = Styled(fonts, font, style & ~FontStyle.Underline);
					if ((style & FontStyle.Underline) != 0) underlineFrom ??= x;
					else Underline(x);
					g.DrawString(text, styled, brush, x, y, format);
					x += g.MeasureString(text, styled, int.MaxValue, format).Width;
				}
				Underline(x);
				y += line;
			}
		} finally {
			foreach (var f in fonts.Values) f.Dispose();
		}
	}

	// ================================================================ RTF (Word's footnotes)

	const string Words = @"((?:\\[a-z]+-?\d* ?)*)", V = @"\\v(?![a-z0-9]) ?";

	/// RTF as RichEdit wrote it, with Word's footnotes for the markers.
	/// noteSize: the document's notes' size (its Footnote Text style's; 0: NoteSize of each footnote's text's).
	public static string ToWordRtf(string rtf, string font, float size, float noteSize = 0) {
		if (!rtf.Contains("[[fn]]", StringComparison.Ordinal)) return rtf;
		int half = (int)Math.Round((noteSize > 0 ? noteSize : NoteSize(size)) * 2);
		// (footnotes, not endnotes, at the foot of the page: without \fet0 Word makes them endnotes; a document property,
		// before the text, where RichEdit writes \viewkind)
		int properties = rtf.IndexOf(@"\viewkind", StringComparison.Ordinal);
		if (properties > 0) rtf = rtf.Insert(properties, @"\fet0\ftnbj");
		return Regex.Replace(rtf, V + Words + @"\[\[fn\]\]" + Words + @"\\v0 ?([^{}]*?)" + V + Words + @"\[\[fn:((?:[^\]{}]|\\[{}]|\](?!\]))*)\]\]" + Words + @"\\v0 ?", m =>
			m.Groups[1].Value + m.Groups[2].Value + m.Groups[4].Value
			+ @"{\super\chftn}{\footnote\pard\nosupersub\fs" + (noteSize > 0 ? half : HalfPoints(rtf, m.Index, half)) + @" {\super\chftn} " + MarksToRtf(m.Groups[5].Value.Trim()) + "}"    // (no \par at its end, as Word's: one would be an empty paragraph)
			+ Delimited(string.Concat(Regex.Matches(m.Groups[3].Value, @"\\[a-z]+-?\d* ?").Select(w => w.Value)) + m.Groups[6].Value),    // (the number's formatting, as it was after it)
			RegexOptions.Singleline);
	}

	// A note's marks in its RTF (RichEdit writes them as \uN, N negative past 32767) as RTF's own: \b \i \ul and their
	// ends, \par between its paragraphs.
	static string MarksToRtf(string rtf) => Regex.Replace(rtf, @"\\u(-?\d+)(?:\?|\\'[0-9a-fA-F]{2})?", m => {
		int n = int.Parse(m.Groups[1].Value);
		return (char)(n < 0 ? n + 65536 : n) switch {
			BoldOn => @"\b ", BoldOff => @"\b0 ", ItalicOn => @"\i ", ItalicOff => @"\i0 ", UnderlineOn => @"\ul ", UnderlineOff => @"\ulnone ",
			ParagraphMark => @"\par ",
			_ => m.Value
		};
	});

	// The note's size in half points: NoteSize of the text's size where the footnote is (the last \fs before it), else
	// fallback's. (Its font is the text's too: the note goes on with the formatting it is in, not \plain's.)
	static int HalfPoints(string rtf, int at, int fallback) {
		var fs = Regex.Match(rtf[..at], @"\\fs(\d+)", RegexOptions.RightToLeft);
		return fs.Success ? (int)Math.Round(NoteSize(int.Parse(fs.Groups[1].Value) / 2f) * 2) : fallback;
	}

	/// Control words with the space that ends them (else the text's own space after them would be taken for it).
	public static string Delimited(string words) => words.Length > 0 && !words.EndsWith(' ') ? words + " " : words;

	// ================================================================ .docx (the built-in converter)

	/// The built-in converter's .docx (made from the HTML: links to a list of the notes at its end) with real footnotes:
	/// each link a footnote reference, the notes (notes: the document's, in order) in the footnotes part, the list gone.
	public static void IntoDocx(string path, List<string> notes, float noteSize = 0) {
		if (notes.Count == 0) return;
		using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, true);
		var main = package.MainDocumentPart!;
		var body = main.Document.Body!;
		var links = body.Descendants<W.Hyperlink>().Where(l => l.Anchor?.Value?.StartsWith("rte-fn-", StringComparison.Ordinal) == true).ToList();
		if (links.Count == 0) return;
		var part = main.FootnotesPart ?? main.AddNewPart<DocumentFormat.OpenXml.Packaging.FootnotesPart>();
		part.Footnotes ??= new W.Footnotes(
			new W.Footnote(new W.Paragraph(new W.Run(new W.SeparatorMark()))) { Type = W.FootnoteEndnoteValues.Separator, Id = -1 },
			new W.Footnote(new W.Paragraph(new W.Run(new W.ContinuationSeparatorMark()))) { Type = W.FootnoteEndnoteValues.ContinuationSeparator, Id = 0 });
		long next = part.Footnotes.Elements<W.Footnote>().Select(f => f.Id?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
		var sizeRun = new W.RunProperties(new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript });
		foreach (var link in links) {
			if (!int.TryParse(link.Anchor!.Value!["rte-fn-".Length..], out int number) || number < 1 || number > notes.Count) continue;
			long id = next++;
			// (the note's paragraphs, the first with the number; their runs, each with its bold, italic and underline)
			var paragraph = new W.Paragraph(new W.Run(new W.RunProperties(new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript }), new W.FootnoteReferenceMark()));
			var footnote = new W.Footnote(paragraph) { Id = id };
			foreach (var (runText, style) in Runs(" " + notes[number - 1])) {
				var pieces = runText.Split(ParagraphMark);
				for (int k = 0; k < pieces.Length; k++) {
					if (k > 0) footnote.Append(paragraph = new W.Paragraph());
					if (pieces[k] == "") continue;
					var properties = new W.RunProperties();
					if ((style & FontStyle.Bold) != 0) properties.Append(new W.Bold());
					if ((style & FontStyle.Italic) != 0) properties.Append(new W.Italic());
					if ((style & FontStyle.Underline) != 0) properties.Append(new W.Underline { Val = W.UnderlineValues.Single });
					if (noteSize > 0) properties.Append(new W.FontSize { Val = ((int)Math.Round(noteSize * 2)).ToString() });
					paragraph.Append(new W.Run(properties, new W.Text(pieces[k]) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
				}
			}
			part.Footnotes.Append(footnote);
			link.Parent!.InsertBefore(new W.Run((W.RunProperties)sizeRun.CloneNode(true), new W.FootnoteReference { Id = id }), link);
			// (the superscript run around the link, if the converter made one, has nothing left in it)
			link.Remove();
		}
		// the list of notes at the end: its paragraphs (the last ones, however the converter made the notes' paragraphs
		// into its own: taken out only when they are all the notes' text)
		var paragraphs = body.Elements<W.Paragraph>().ToList();
		static string Bare(string text) => Regex.Replace(text, @"\s+", "");
		string wanted = Bare(string.Concat(notes.Select(Plain)));
		string found = "";
		var list = new List<W.Paragraph>();
		for (int i = paragraphs.Count - 1; i >= 0 && found.Length < wanted.Length; i--) {
			string text = Bare(paragraphs[i].InnerText);
			if (text == "" && list.Count == 0) continue;
			if (!wanted.EndsWith(text + found, StringComparison.Ordinal)) break;
			found = text + found;
			list.Add(paragraphs[i]);
		}
		if (found == wanted) foreach (var p in list) p.Remove();
		part.Footnotes.Save();
		main.Document.Save();
	}

	/// A .docx's footnotes as markers in its text (for the built-in reader, which goes through HTML): each reference's
	/// run becomes "[[fn]]1[[fn:note]]", its note from the footnotes part. Returns whether there were any.
	public static bool MarkDocx(DocumentFormat.OpenXml.Packaging.WordprocessingDocument package) {
		var main = package.MainDocumentPart;
		var body = main?.Document?.Body;
		var notes = main?.FootnotesPart?.Footnotes;
		if (body == null || notes == null) return false;
		// (the notes' size: their runs' most common w:sz, else the Footnote Text style's)
		var runSizes = notes.Elements<W.Footnote>().Where(f => f.Id?.Value > 0).SelectMany(f => f.Descendants<W.Run>())
			.Where(r => r.GetFirstChild<W.FootnoteReferenceMark>() == null && r.RunProperties?.FontSize?.Val?.Value is { } v && int.TryParse(v, out _))
			.Select(r => int.Parse(r.RunProperties!.FontSize!.Val!.Value!)).ToList();
		var styleSize = main!.StyleDefinitionsPart?.Styles?.Elements<W.Style>().FirstOrDefault(s => s.StyleId?.Value == "FootnoteText")?.StyleRunProperties?.FontSize?.Val?.Value;
		LoadedSize = runSizes.Count > 0 ? runSizes.GroupBy(s => s).OrderByDescending(g => g.Count()).First().Key / 2f
			: int.TryParse(styleSize, out int fromStyle) ? fromStyle / 2f : null;
		var texts = notes.Elements<W.Footnote>().Where(f => f.Id?.Value > 0)
			.ToDictionary(f => f.Id!.Value, f => TidyMarks(FromRuns(f.Descendants<W.Paragraph>().SelectMany(p =>
				p.Descendants<W.Run>().Select(r => (string.Concat(r.Descendants().Select(RunText)), DocxStyle(r.RunProperties)))
					.Append((ParagraphMark.ToString(), FontStyle.Regular))))));
		bool any = false;
		int number = 0;
		foreach (var reference in body.Descendants<W.FootnoteReference>().ToList()) {
			if (reference.Id?.Value is not long id || !texts.TryGetValue(id, out var note)) continue;
			var run = reference.Ancestors<W.Run>().FirstOrDefault();
			if (run == null) continue;
			number++;
			run.InsertBeforeSelf(new W.Run(new W.Text(Start + number + NoteOpen + Escape(note) + NoteClose) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
			run.Remove();
			any = true;
		}
		return any;
	}

	// A .docx run's element's text: its text, a tab's space, a line break (Shift+Enter) a paragraph's mark (the nearest a
	// note has).
	static string RunText(DocumentFormat.OpenXml.OpenXmlElement e) => e switch {
		W.Text t => t.Text,
		W.TabChar => " ",
		W.Break { Type: null } or W.CarriageReturn => ParagraphMark.ToString(),
		_ => ""
	};

	// A .docx run's bold, italic and underline (an on/off property without its value is on).
	static FontStyle DocxStyle(W.RunProperties? properties) {
		var style = FontStyle.Regular;
		if (properties == null) return style;
		if (properties.Bold is { } bold && (bold.Val == null || bold.Val.Value)) style |= FontStyle.Bold;
		if (properties.Italic is { } italic && (italic.Val == null || italic.Val.Value)) style |= FontStyle.Italic;
		if (properties.Underline is { } underline && (underline.Val == null || underline.Val.Value != W.UnderlineValues.None)) style |= FontStyle.Underline;
		return style;
	}

	/// Word's (and LibreOffice's) RTF with markers for its footnotes, before RichEdit loads it (it shows a note's text in
	/// the middle of the paragraph, or drops it).
	public static string FromWordRtf(string rtf) {
		if (!rtf.Contains(@"\footnote", StringComparison.Ordinal)) return rtf;
		var sb = new StringBuilder();
		int at = 0, number = 0;
		// (the search goes on after the group found: IndexOfNote gives the group's "{", before the "\footnote" it found)
		for (int i = IndexOfNote(rtf, 0); i >= 0; i = IndexOfNote(rtf, Math.Max(at, i + 10))) {
			if (i < at) continue;
			int end = GroupEnd(rtf, i);
			if (end < 0) break;
			string note = NoteText(rtf[i..(end + 1)]);
			// (the notes' size, from the first one: its text's most common \fs)
			if (number == 0) {
				var sizes = Regex.Matches(rtf[i..(end + 1)], @"\\fs(\d+)").Select(s => int.Parse(s.Groups[1].Value)).ToList();
				LoadedSize = sizes.Count > 0 ? sizes.GroupBy(s => s).OrderByDescending(g => g.Count()).First().Key / 2f : null;
			}
			// the reference: the \chftn just before the note (Word puts the note in the reference's group)
			int reference = rtf.LastIndexOf(@"\chftn", i, Math.Max(0, i - at), StringComparison.Ordinal);
			number++;
			string mark = @"{\v [[fn]]}" + number;
			if (reference >= 0 && reference >= at) {
				sb.Append(rtf, at, reference - at);
				sb.Append(mark);
				int afterWord = reference + 6;
				if (afterWord < rtf.Length && rtf[afterWord] == ' ') afterWord++;
				sb.Append(rtf, afterWord, i - afterWord);
			} else {
				sb.Append(rtf, at, i - at);
				sb.Append(@"{\super " + mark + "}");
			}
			sb.Append(@"{\v\nosupersub [[fn:").Append(RtfEscape(Escape(note))).Append("]]}");
			at = end + 1;
		}
		if (number == 0) return rtf;
		sb.Append(rtf, at, rtf.Length - at);
		return sb.ToString();
	}

	// The next "{\footnote" or "{\*\footnote" group.
	static int IndexOfNote(string rtf, int from) {
		for (int i = rtf.IndexOf(@"\footnote", from, StringComparison.Ordinal); i >= 0; i = rtf.IndexOf(@"\footnote", i + 1, StringComparison.Ordinal)) {
			if (i + 9 < rtf.Length && char.IsLetter(rtf[i + 9])) continue;
			if (i >= 3 && rtf[i - 1] == '*' && rtf[i - 2] == '\\' && rtf[i - 3] == '{') return i - 3;    // {\*\footnote (LibreOffice's)
			if (i >= 1 && rtf[i - 1] == '{') return i - 1;    // {\footnote (Word's)
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

	static readonly Regex Token = new(@"\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z])|([{}])|([^\\{}\r\n]+)|[\r\n]+");
	static readonly HashSet<string> SkippedGroups = new() { "fonttbl", "colortbl", "stylesheet", "pict", "object", "info", "shp", "nonshppict", "fldinst" };

	// A footnote group's text (its number, \chftn, left out), with marks for its bold, italic and underline.
	static string NoteText(string group) {
		var sb = new StringBuilder();
		var tokens = Token.Matches(group).ToList();
		int depth = 0, skipUntil = -1, skipChars = 0;
		// (the formatting: each group's own, back to its outer group's at its end; written as marks before the text it is on)
		var style = FontStyle.Regular;
		var written = FontStyle.Regular;
		var outer = new Stack<FontStyle>();
		void Text(string text) {
			if (text.Length == 0) return;
			if (style != written) {
				foreach (var (flag, on, off) in new[] { (FontStyle.Bold, BoldOn, BoldOff), (FontStyle.Italic, ItalicOn, ItalicOff), (FontStyle.Underline, UnderlineOn, UnderlineOff) })
					if ((style & flag) != (written & flag)) sb.Append((style & flag) != 0 ? on : off);
				written = style;
			}
			sb.Append(text);
		}
		void Set(FontStyle flag, bool on) => style = on ? style | flag : style & ~flag;
		for (int i = 0; i < tokens.Count; i++) {
			var t = tokens[i];
			if (t.Value == "{") {
				depth++;
				outer.Push(style);
				var next = i + 1 < tokens.Count ? tokens[i + 1] : null;
				if (skipUntil < 0 && next != null && (next.Groups[4].Value == "*" && depth > 1 || SkippedGroups.Contains(next.Groups[1].Value))) skipUntil = depth;
				continue;
			}
			if (t.Value == "}") { if (depth == skipUntil) skipUntil = -1; depth--; if (outer.Count > 0) style = outer.Pop(); continue; }
			if (skipUntil >= 0) continue;
			if (t.Groups[1].Success) {
				string word = t.Groups[1].Value;
				bool zero = t.Groups[2].Value == "0";
				switch (word) {
					case "tab": Text(" "); break;
					// (a paragraph; Word's line break in a note, Shift+Enter, too: the nearest a note has)
					case "par": case "line": Text(ParagraphMark.ToString()); break;
					case "u": Text(((char)(short)int.Parse(t.Groups[2].Value)).ToString()); skipChars = 1; break;
					case "b": Set(FontStyle.Bold, !zero); break;
					case "i": Set(FontStyle.Italic, !zero); break;
					case "ulnone": Set(FontStyle.Underline, false); break;
					case "plain": style = FontStyle.Regular; break;
					default:
						// (\ul, and its kinds: \uld, \uldb, \ulw, \ulwave and the like; not \ulc, its colour)
						if (word.StartsWith("ul", StringComparison.Ordinal) && !word.StartsWith("ulc", StringComparison.Ordinal)) Set(FontStyle.Underline, !zero);
						break;
				}
			} else if (t.Groups[3].Success) {
				if (skipChars > 0) { skipChars--; continue; }
				Text(Encoding.Latin1.GetString(new[] { Convert.ToByte(t.Groups[3].Value, 16) }));
			} else if (t.Groups[4].Success) {
				if (t.Groups[4].Value is "\\" or "{" or "}") Text(t.Groups[4].Value);
				else if (t.Groups[4].Value == "~") Text("\u00A0");
			} else if (t.Groups[6].Success) {
				string s = t.Value;
				if (skipChars > 0 && s.Length > 0) { s = s[1..]; skipChars = 0; }
				Text(s);
			}
		}
		style = FontStyle.Regular;
		Text("\u0001");
		return TidyMarks(sb.ToString().TrimEnd('\u0001'));
	}

	// A note's text with its spaces made one and none at its ends or its paragraphs' (marks around them aside), no
	// paragraph marks at its ends, and marks that switch a format on and straight off again left out.
	static string TidyMarks(string note) {
		note = Regex.Replace(note, @"\s+", " ");
		note = Regex.Replace(note, " * *", "");
		note = Regex.Replace(note, "([-]+) +", "$1");
		note = Regex.Replace(note, " +([-]+)", "$1");
		note = Regex.Replace(note, "^([-]*)[ ]+", "$1");
		note = Regex.Replace(note, "[ ]+([-]*)$", "$1");
		note = Regex.Replace(note, "^([\uE021-\uE026]*) ", "$1");
		note = Regex.Replace(note, " ([\uE021-\uE026]*)$", "$1");
		note = Regex.Replace(note, "\uE021\uE022|\uE023\uE024|\uE025\uE026", "");
		return Nested(note.Trim());
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

	// ================================================================ HTML and Markdown

	static readonly Regex HtmlNote = new(@"\[\[fn\]\]((?:<[^>]*>|[^<\[\]\r])*?)\[\[fn:((?:[^\]\r]|\](?!\]))*)\]\]", RegexOptions.Singleline);

	/// HTML from the editor with each footnote a numbered link to the list of notes, which goes at the end of body.
	public static string ToHtml(string body) {
		if (!Any(body)) return body;
		var notes = new List<string>();
		// (the library can write the number's superscript as tags between the markers: those are kept, its text goes)
		body = HtmlNote.Replace(body, m => {
			notes.Add(m.Groups[2].Value);
			int n = notes.Count;
			string tags = string.Concat(Regex.Matches(m.Groups[1].Value, "<[^>]*>").Select(t => t.Value));
			return $"<sup><a href=\"#rte-fn-{n}\" id=\"rte-fnref-{n}\" class=\"rte-footnote\">{n}</a></sup>" + tags;
		});
		var list = new StringBuilder("<ol class=\"rte-footnotes\">");
		for (int i = 0; i < notes.Count; i++) {
			// (a note of more than one paragraph: each a <p>, the formatting tags closed and opened again around them)
			var paragraphs = SplitParagraphs(DecodeMarks(notes[i])).Select(ToTags).ToList();
			string note = paragraphs.Count == 1 ? paragraphs[0] : string.Concat(paragraphs.Select(p => "<p>" + p + "</p>"));
			list.Append($"<li id=\"rte-fn-{i + 1}\">{note}</li>");
		}
		return body + list.Append("</ol>");
	}

	/// A note's paragraphs, each with its own marks (a format on across a paragraph's end switched off at it and on
	/// again in the next).
	public static List<string> SplitParagraphs(string note) {
		var paragraphs = new List<List<(string Text, FontStyle Style)>> { new() };
		foreach (var (text, style) in Runs(note)) {
			var pieces = text.Split(ParagraphMark);
			for (int k = 0; k < pieces.Length; k++) {
				if (k > 0) paragraphs.Add(new());
				paragraphs[^1].Add((pieces[k], style));
			}
		}
		return paragraphs.Select(p => FromRuns(p)).ToList();
	}

	// The marks in HTML as their characters (the writer may give them as character references).
	static string DecodeMarks(string html) => Regex.Replace(html, @"&#(?:x([0-9a-fA-F]+)|(\d+));", m => {
		int n = m.Groups[1].Success ? Convert.ToInt32(m.Groups[1].Value, 16) : int.Parse(m.Groups[2].Value);
		return n is >= 0xE021 and <= 0xE027 ? ((char)n).ToString() : m.Value;
	});

	// A note's HTML with paragraph marks between its paragraphs (<p>, <div>) and for its line breaks (<br>).
	static string HtmlParagraphs(string html) {
		html = Regex.Replace(html.Trim(), @"\s+", " ");
		html = Regex.Replace(html, @"</(?:p|div)>\s*<(?:p|div)\b[^>]*>|<br\s*/?>", ParagraphMark.ToString(), RegexOptions.IgnoreCase);
		return Regex.Replace(html, @"</?(?:p|div)\b[^>]*>", "", RegexOptions.IgnoreCase);
	}

	/// HTML with markers for the footnotes ToHtml wrote (links to the list at the end, and the list taken out).
	public static string FromHtml(string html) {
		var list = Regex.Match(html, @"<ol\s+class=""rte-footnotes"">(.*?)</ol>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
		if (!list.Success) return html;
		var notes = Regex.Matches(list.Groups[1].Value, @"<li\s+id=""rte-fn-(\d+)""\s*>(.*?)</li>", RegexOptions.Singleline | RegexOptions.IgnoreCase)
			.ToDictionary(m => m.Groups[1].Value, m => TidyMarks(System.Net.WebUtility.HtmlDecode(Regex.Replace(FromTags(HtmlParagraphs(m.Groups[2].Value)), "<[^>]+>", ""))));
		html = html.Remove(list.Index, list.Length);
		return Regex.Replace(html, @"<a\s+href=""#rte-fn-(\d+)""[^>]*>(.*?)</a>", m =>
			notes.TryGetValue(m.Groups[1].Value, out var note) ? $"[[fn]]<sup>{m.Groups[1].Value}</sup>[[fn:{System.Net.WebUtility.HtmlEncode(Escape(note))}]]" : m.Value,
			RegexOptions.Singleline | RegexOptions.IgnoreCase);
	}

	// Markdown escapes the markers' brackets: "\[\[fn\]\]".
	// (an escape, or another character, or a bracket that doesn't close the marker: one way each, so that a near miss
	// fails at once rather than trying every way to split the backslashes)
	static readonly Regex MarkdownNote = new(@"\\?\[\\?\[fn\\?\]\\?\]((?:<[^>\r\n]{0,20}>|[^\[\]\\\r\n<]){0,40}?)\\?\[\\?\[fn:((?:\\[^\r\n]|[^\\\]\r\n]|\](?!\\?\]))*)\\?\]\\?\]");

	/// Markdown from the editor with its footnotes: [^1] in the text, the notes at the end ([^1]: note).
	public static string ToMarkdown(string markdown) {
		if (!markdown.Contains("fn", StringComparison.Ordinal)) return markdown;
		var notes = new List<string>();
		markdown = MarkdownNote.Replace(markdown, m => {
			notes.Add(m.Groups[2].Value);
			return $"[^{notes.Count}]";
		});
		if (notes.Count == 0) return markdown;
		var sb = new StringBuilder(markdown.TrimEnd('\n')).Append("\n\n");
		// (a note's paragraphs after its first: indented four spaces after an empty line, Markdown's footnote paragraphs)
		for (int i = 0; i < notes.Count; i++) sb.Append($"[^{i + 1}]: {string.Join("\n\n    ", SplitParagraphs(notes[i]).Select(MarksToMarkdown))}\n");
		return sb.ToString();
	}

	// A note's marks as Markdown's ** * and <u> (spaces just inside them moved out, as Markdown needs), and back.
	static string MarksToMarkdown(string note) {
		note = Regex.Replace(Nested(note), "([]+)(\\s+)", "$2$1");
		note = Regex.Replace(note, "(\\s+)([]+)", "$2$1");
		return note.Replace(BoldOn.ToString(), "**").Replace(BoldOff.ToString(), "**").Replace(ItalicOn.ToString(), "*").Replace(ItalicOff.ToString(), "*")
			.Replace(UnderlineOn.ToString(), "<u>").Replace(UnderlineOff.ToString(), "</u>");
	}
	static string MarkdownToMarks(string note) {
		note = Regex.Replace(note, "<u>", UnderlineOn.ToString(), RegexOptions.IgnoreCase);
		note = Regex.Replace(note, "</u>", UnderlineOff.ToString(), RegexOptions.IgnoreCase);
		note = Regex.Replace(note, @"(?<!\\)(\*\*|__)(?=\S)(.+?)(?<=[^\s\\])\1", BoldOn + "$2" + BoldOff);
		note = Regex.Replace(note, @"(?<![\\*])\*(?=[^\s*])(.+?)(?<=[^\s\\*])\*(?!\*)", ItalicOn + "$1" + ItalicOff);
		note = Regex.Replace(note, @"(?<![\\\w])_(?=[^\s_])(.+?)(?<=[^\s\\_])_(?!\w)", ItalicOn + "$1" + ItalicOff);
		return note;
	}

	/// Markdown with markers for its footnotes ([^1] and [^1]: note), before it is loaded (the notes' lines taken out).
	public static string FromMarkdown(string markdown) {
		if (!markdown.Contains("[^", StringComparison.Ordinal)) return markdown;
		var notes = new Dictionary<string, string>();
		// (a note's paragraphs after its first: its lines indented four spaces or a tab, after an empty line; an indented
		// line straight after a line goes on with its paragraph)
		markdown = Regex.Replace(markdown, @"^\[\^([^\]\s]+)\]:[ \t]*(.*)((?:\r?\n(?:[ \t]*\r?\n)*(?: {4}|\t)[ \t]*\S.*)*)(?:\r?\n|$)", m => {
			var paragraphs = new List<string>();
			var lines = new List<string> { m.Groups[2].Value.Trim() };
			foreach (var line in m.Groups[3].Value.Split('\n').Skip(1)) {
				if (line.Trim() == "") { if (lines.Count > 0) paragraphs.Add(string.Join(" ", lines)); lines.Clear(); }
				else lines.Add(line.Trim());
			}
			if (lines.Count > 0) paragraphs.Add(string.Join(" ", lines));
			notes[m.Groups[1].Value] = string.Join(ParagraphMark.ToString(), paragraphs.Select(MarkdownToMarks));
			return "";
		}, RegexOptions.Multiline);
		if (notes.Count == 0) return markdown;
		int number = 0;
		return Regex.Replace(markdown, @"\[\^([^\]\s]+)\]", m => notes.TryGetValue(m.Groups[1].Value, out var note)
			? $@"\[\[fn\]\]{++number}\[\[fn:{Escape(note).Replace("[", "\\[").Replace("]", "\\]").Replace("_", "\\_").Replace("*", "\\*")}\]\]" : m.Value);
	}
}
