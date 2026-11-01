// Sections, for AhkWord (AhkWord.cs, through #load). Iteration 63.
//
// A section break is a page break whose paragraph also holds the new section's settings, as a hidden marker in front of
// the form feed: [[sect:{...}]] (hidden text, CFE_HIDDEN, RTF \v, which RichEdit keeps through editing, undo, copying and
// its RTF, as the cross-references' and footnotes' markers are kept; rte_crossrefs.cs). So a section's header, footer
// and page numbering travel with its break. The text before the first break is the first section, with the document's
// own header and footer (WriterForm.Bands), as before.
// A continuous section break (Word's, iteration 67) starts no page: it is a paragraph of its own, all hidden (the marker
// and the paragraph mark, so it takes no room and the caret doesn't stop in it), and its section starts after it.
//   - SectionSettings: a section's own: its header and footer (PageBands; Same: as the section before's, Word's "Link to
//     Previous", until one is given), whether its page numbers start again (Restart, on unless turned off) and at what
//     number (Start), and (Page) the page setup it has of its own.
//   - PageSection: a section as laid out: where its text starts, its header and footer, its numbering.
//   - Sections: the markers read and written; a section break put in; each page's section, number shown ({page}), the
//     section's page count ({sectionpages}) and whether it is the section's first page (its own first-page header).
using System.Text.Json;
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

/// A section's own page setup: what its file gives it that differs from the first section's (twips; null: the document's).
/// Top and Bottom are shown and printed; the rest is only kept, and written back when the file is saved.
public sealed record SectionPage(int? Width = null, int? Height = null, int? Left = null, int? Right = null, int? Top = null, int? Bottom = null, bool? Landscape = null, int? HeaderY = null, int? FooterY = null)
{
	/// What this section has that first hasn't (null: nothing).
	public SectionPage? Against(SectionPage first) {
		static int? Own(int? a, int? b) => a == b ? null : a;
		var own = new SectionPage(Own(Width, first.Width), Own(Height, first.Height), Own(Left, first.Left), Own(Right, first.Right), Own(Top, first.Top), Own(Bottom, first.Bottom),
			(Landscape ?? false) == (first.Landscape ?? false) ? null : Landscape ?? false, Own(HeaderY, first.HeaderY), Own(FooterY, first.FooterY));
		return own == new SectionPage() ? null : own;
	}

	/// With other top and bottom margins (twips; null: the document's); null if that leaves it nothing of its own.
	public static SectionPage? WithMargins(SectionPage? page, int? top, int? bottom) {
		var own = (page ?? new SectionPage()) with { Top = top, Bottom = bottom };
		return own == new SectionPage() ? null : own;
	}

	/// Whether it has more than margins top and bottom and the header's and footer's distances: things kept but not shown.
	[System.Text.Json.Serialization.JsonIgnore] public bool KeptOnly => Width != null || Height != null || Left != null || Right != null || Landscape != null;
}

/// A section's columns: how many, the gap between them (twips), whether a line is drawn in the gap; of equal width, or
/// each of its own (Widths, Gaps: twips; Gaps as Word's files may have them, one after each column but the last; on the
/// page, in proportion to the text area's width). They are the printed page's: RichEdit shows a document at one width, so
/// the text is edited in one column.
public sealed record SectionColumns(int Count = 2, int Gap = 720, bool Line = false, int[]? Widths = null, int[]? Gaps = null)
{
	/// Columns narrower than this (twips) aren't used: the section prints in one.
	const int Narrowest = 1080;

	/// Null (one column) or the columns asked for.
	public static SectionColumns? Of(int count, int gap, bool line, int[]? widths = null, int[]? gaps = null) {
		if (count < 2) return null;
		gap = Math.Max(0, gap);
		if (widths?.Length != count || widths.Any(w => w <= 0) || (widths.Distinct().Count() == 1 && (gaps == null || gaps.Distinct().Count() == 1))) return new SectionColumns(count, gaps?.Length == count - 1 ? Math.Max(0, gaps[0]) : gap, line);
		return new SectionColumns(count, gap, line, widths, gaps?.Length == count - 1 ? gaps : null);
	}

	int GapAfter(int index) => Math.Max(0, Gaps?[index] ?? Gap);

	/// The narrowest column's width in a text area this wide.
	public int NarrowestWidth(int width) => Enumerable.Range(0, Count).Min(i => { var (left, right) = Slot(0, width, i); return right - left; });

	public bool FitIn(int width) => Enumerable.Range(0, Count).All(i => { var (left, right) = Slot(0, width, i); return right - left >= Narrowest; });

	/// Column index's left and right in a text area from left to right.
	public (int Left, int Right) Slot(int left, int right, int index) {
		int area = right - left;
		if (Widths == null) {
			int width = (area - (Count - 1) * Gap) / Count, start = left + index * (width + Gap);
			return (start, start + width);
		}
		double scale = (double)area / (Widths.Sum() + Enumerable.Range(0, Count - 1).Sum(GapAfter));
		double at = left;
		for (int i = 0; i < index; i++) at += (Widths[i] + GapAfter(i)) * scale;
		return ((int)at, (int)(at + Widths[index] * scale));
	}

	/// A section's RTF properties' columns (null: one).
	public static SectionColumns? FromRtf(string properties) {
		var count = Regex.Match(properties, @"\\cols(\d+)");
		if (!count.Success) return null;
		var gap = Regex.Match(properties, @"\\colsx(\d+)");
		var own = Regex.Matches(properties, @"\\colno\d+\s*\\colw(\d+)(?:\s*\\colsr(\d+))?");
		return Of(int.Parse(count.Groups[1].Value), gap.Success ? int.Parse(gap.Groups[1].Value) : 720, Regex.IsMatch(properties, @"\\linebetcol(?![a-zA-Z])"),
			own.Count == 0 ? null : own.Select(m => int.Parse(m.Groups[1].Value)).ToArray(),
			own.Count == 0 ? null : own.Take(own.Count - 1).Select(m => m.Groups[2].Success ? int.Parse(m.Groups[2].Value) : gap.Success ? int.Parse(gap.Groups[1].Value) : 720).ToArray());
	}

	public string ToRtf() => $@"\cols{Count}\colsx{Gap}" + (Line ? @"\linebetcol" : "")
		+ (Widths == null ? "" : string.Concat(Widths.Select((w, i) => $@"\colno{i + 1}\colw{w}" + (i + 1 < Count ? $@"\colsr{GapAfter(i)}" : ""))));
}

/// A section's own settings, as its break's marker holds them.
public sealed class SectionSettings
{
	/// The section's header and footer (null, or Same: the section before's).
	public PageBands? Bands { get; set; }
	/// Word's "Link to Previous": the header and footer of the section before.
	public bool Same { get; set; } = true;
	/// Whether page numbers start again in this section, and at what number.
	public bool Restart { get; set; } = true;
	public int Start { get; set; } = 1;
	/// The section's own page setup (null: the document's).
	public SectionPage? Page { get; set; }
	/// The section's columns (null: one).
	public SectionColumns? Columns { get; set; }
	/// Whether the section starts on the same page as the text before it.
	public bool Continuous { get; set; }

	public SectionSettings Copy() => new() { Bands = Bands == null ? null : Sections.CopyBands(Bands), Same = Same, Restart = Restart, Start = Start, Page = Page, Columns = Columns, Continuous = Continuous };
}

/// A section as laid out: where its text starts (the first section at 0), its header and footer, its numbering, its
/// own top and bottom margins (1/100 inch; null: the document's), its columns (null: one), whether it starts on the same
/// page as the text before it, and where the text before it stops (a continuous break's hidden paragraph).
public sealed record PageSection(int Start, PageBands Bands, bool Restart, int StartAt, int? Top = null, int? Bottom = null, SectionColumns? Columns = null, bool Continuous = false, int Stop = 0);

public static class Sections
{
	/// A section break's marker: [[sect:JSON]] (no "]]" in the JSON: a "]" in its text escaped, two that close lists
	/// apart).
	public static readonly Regex Mark = new(@"\[\[sect:(\{(?:[^\]\r]|\](?!\]))*\})\]\]");

	static readonly JsonSerializerOptions Json = new() {
		DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
		Encoder = System.Text.Encodings.Web.JavaScriptEncoder.Create(MarkerSafe()),
	};

	// any text but "]" (escaped in strings: a footer's "]]" can't end the marker early)
	static System.Text.Encodings.Web.TextEncoderSettings MarkerSafe() {
		var settings = new System.Text.Encodings.Web.TextEncoderSettings(System.Text.Unicode.UnicodeRanges.All);
		settings.ForbidCharacter(']');
		return settings;
	}

	public static string MarkerFor(SectionSettings settings) => "[[sect:" + JsonSerializer.Serialize(settings, Json).Replace("]]", "] ]") + "]]";

	public static SectionSettings? Parse(string json) {
		try { return JsonSerializer.Deserialize<SectionSettings>(json, Json); }
		catch (JsonException) { return null; }
	}

	public static PageBands CopyBands(PageBands bands) => JsonSerializer.Deserialize<PageBands>(JsonSerializer.Serialize(bands, Json), Json) ?? new PageBands();

	static string Story(ITextDocument2 doc) => doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";

	/// A section break in a document: its marker (where, how long), where its section's text starts (after the break's
	/// paragraph), its settings.
	public sealed record Break(int MarkStart, int MarkLength, int SectionStart, SectionSettings Settings);

	/// The document's section breaks, in order: markers right in front of a page break's form feed, or a continuous one's
	/// paragraph mark (others are ignored).
	public static List<Break> Breaks(ITextDocument2 doc) => Breaks(Story(doc));

	public static List<Break> Breaks(string story) {
		var breaks = new List<Break>();
		foreach (Match m in Mark.Matches(story)) {
			int after = m.Index + m.Length;
			if (after >= story.Length || Parse(m.Groups[1].Value) is not { } settings) continue;
			if (settings.Continuous) {
				if (story[after] == '\r') breaks.Add(new Break(m.Index, m.Length, after + 1, settings));
				continue;
			}
			if (story[after] != PageBreaks.Break) continue;
			int start = after + 1 < story.Length && story[after + 1] == '\r' ? after + 2 : after + 1;
			breaks.Add(new Break(m.Index, m.Length, start, settings));
		}
		return breaks;
	}

	/// Whether the text in front of a form feed (from its paragraph's start) is a section break's marker: the page break
	/// stays one paragraph with it (PageBreaks.Normalize).
	public static bool IsMarkerBefore(string story, int formFeed) {
		int start = story.LastIndexOf('\r', Math.Max(0, formFeed - 1)) + 1;
		if (formFeed <= 0 || start >= formFeed) return false;
		var m = Mark.Match(story, start, formFeed - start);
		return m.Success && m.Index == start && m.Index + m.Length == formFeed;
	}

	/// The sections as laid out: the first with the document's header and footer (numbered from 1), then one after each
	/// break (its own header and footer, or the section before's). firstColumns: the first section's, the document's.
	public static List<PageSection> Resolve(ITextDocument2 doc, PageBands first, SectionColumns? firstColumns = null) => Resolve(Breaks(doc), first, firstColumns);

	public static List<PageSection> Resolve(List<Break> breaks, PageBands first, SectionColumns? firstColumns = null) {
		var sections = new List<PageSection> { new(0, first, true, 1, Columns: firstColumns) };
		foreach (var b in breaks) {
			var bands = b.Settings.Same || b.Settings.Bands == null ? sections[^1].Bands : b.Settings.Bands;
			static int? Hundredths(int? twips) => twips is int t ? (int)Math.Round(Math.Abs(t) / 14.4) : null;
			sections.Add(new PageSection(b.SectionStart, bands, b.Settings.Restart, Math.Max(0, b.Settings.Start), Hundredths(b.Settings.Page?.Top), Hundredths(b.Settings.Page?.Bottom), b.Settings.Columns, b.Settings.Continuous, b.MarkStart));
		}
		return sections;
	}

	/// A page's top and bottom margins (1/100 inch): its section's own, else the document's (top, bottom), with its header
	/// and footer pushing the text in; on: 1 for the section's first page, else 2. A section's own that would leave less
	/// than an inch of text on a page pageHeight tall aren't used.
	public static (int Top, int Bottom) MarginsOn(PageSection section, int on, int top, int bottom, int pageHeight) {
		int ownTop = section.Top ?? top, ownBottom = section.Bottom ?? bottom;
		if (pageHeight - ownTop - ownBottom < 100) (ownTop, ownBottom) = (top, bottom);
		return (Math.Max(ownTop, section.Bands.HeaderReach(on)), Math.Max(ownBottom, section.Bands.FooterReach(on)));
	}

	/// The sections (numbered from 1) with a page setup that is kept but not shown: paper, orientation, left and right
	/// margins of their own.
	public static List<int> KeptOnly(List<Break> breaks) => breaks.Select((b, i) => (b, i)).Where(x => x.b.Settings.Page?.KeptOnly == true).Select(x => x.i + 2).ToList();

	/// The section a position is in.
	public static int IndexAt(IReadOnlyList<PageSection> sections, int position) {
		int found = 0;
		for (int i = 1; i < sections.Count; i++) if (sections[i].Start <= position) found = i;
		return found;
	}

	/// Each page's section, its number as shown ({page}), its section's page count ({sectionpages}) and its place in the
	/// section (1: the section's first page, which has its own first-page header if the section has one); pages from
	/// where each starts.
	public sealed record PageNumber(int Section, int Shown, int SectionPages, int InSection);

	public static List<PageNumber> Number(IReadOnlyList<int> pageStarts, IReadOnlyList<PageSection> sections) {
		var result = new List<PageNumber>();
		int shown = 0, inSection = 0, previous = -1;
		var sectionOf = pageStarts.Select(p => IndexAt(sections, p)).ToList();
		for (int i = 0; i < pageStarts.Count; i++) {
			int s = sectionOf[i];
			bool first = s != previous;
			inSection = first ? 1 : inSection + 1;
			shown = !first ? shown + 1 : i == 0 ? sections[s].StartAt : sections[s].Restart ? sections[s].StartAt : shown + 1;
			result.Add(new PageNumber(s, shown, sectionOf.Count(x => x == s), inSection));
			previous = s;
		}
		return result;
	}

	// ================================================================ editing

	/// Puts a section break at the selection (a page break with the new section's settings: the section before's header
	/// and footer, Same, page numbers starting again at 1), and the caret after it. Returns the break's marker position.
	/// The caller groups this into one undo step.
	public static int Insert(RichEditBox editor, SectionSettings? settings = null) {
		int at = PageBreaks.Insert(editor);    // (the form feed's position)
		string marker = MarkerFor(settings ?? new SectionSettings());
		var doc = editor.TextDocument;
		doc.Range2(at, at).Text = marker;
		var range = doc.Range2(at, at + marker.Length);
		range.Font.Hidden = tomConstants.tomTrue;
		doc.Selection2.SetRange(at + marker.Length + 2, at + marker.Length + 2);
		return at;
	}

	/// Puts a continuous section break at a position (between paragraphs: in a paragraph, it splits it), with settings, or
	/// those of the section it splits, the header and footer the same, page numbers going on; the caret after it. Returns
	/// where the new section starts. The caller groups this into one undo step.
	public static int InsertContinuous(RichEditBox editor, int at, SectionSettings? settings = null, SectionColumns? firstColumns = null) {
		var doc = editor.TextDocument;
		if (RtfTables.At(editor, at) != null) throw new InvalidOperationException("A section break cannot go in a table.");
		var paragraph = doc.Range2(at, at);
		paragraph.Expand(tomConstants.tomParagraph);
		if (at > paragraph.Start) {
			if (at >= paragraph.End - 1 && paragraph.End < doc.Range2(0, 0).StoryLength) at = paragraph.End;
			else { doc.Range2(at, at).Text = "\r"; at++; }
		}
		if (at <= 0) throw new InvalidOperationException("A section break can't go before the document's first paragraph.");
		var before = BreakFor(doc, at);
		var own = settings?.Copy() ?? new SectionSettings { Restart = false, Page = before?.Settings.Page, Columns = before == null ? firstColumns : before.Settings.Columns };
		own.Continuous = true;
		string marker = MarkerFor(own);
		doc.Range2(at, at).Text = marker + "\r";
		var range = doc.Range2(at, at + marker.Length + 1);
		if ((range.Para.ListType & 0xFFFF) != tomConstants.tomListNone) range.Para.ListType = tomConstants.tomListNone;
		range.Font.Hidden = tomConstants.tomTrue;
		doc.Selection2.SetRange(range.End, range.End);
		return range.End;
	}

	/// A section break's settings changed (its marker written again, hidden).
	public static void Write(ITextDocument2 doc, Break b, SectionSettings settings) {
		string marker = MarkerFor(settings);
		var range = doc.Range2(b.MarkStart, b.MarkStart + b.MarkLength);
		range.Text = marker;
		range = doc.Range2(b.MarkStart, b.MarkStart + marker.Length);
		range.Font.Hidden = tomConstants.tomTrue;
	}

	static readonly Regex DocxMark = new(@"\[\[docsect=(\d+)\]\]");

	/// The built-in reader's marks of where a .docx's sections end (CrossRefs.MarkedDocx) made section breaks, with the
	/// settings of the sections after them (HeaderFooter.SectionsFromDocx; null: only the mark goes).
	public static void FromDocxMarkers(RichEditBox editor, List<SectionSettings?> settings) {
		var doc = editor.TextDocument;
		foreach (var m in DocxMark.Matches(Story(doc)).Cast<Match>().Reverse()) {
			doc.Range2(m.Index, m.Index + m.Length).Text = "";
			int k = int.Parse(m.Groups[1].Value);
			if (k >= settings.Count || settings[k] is not { } section) continue;
			var paragraph = doc.Range2(m.Index, m.Index);
			paragraph.Expand(tomConstants.tomParagraph);
			int at = paragraph.End;
			if (section.Continuous) {
				if (at < doc.Range2(0, 0).StoryLength)
					try { InsertContinuous(editor, at, section); }
					catch (InvalidOperationException) { }    // (in a table: left out)
				continue;
			}
			if ((paragraph.Text ?? "").Trim('\r', '\n') == "") {    // (Word's break in a paragraph of its own: this break takes its place)
				at = paragraph.Start;
				doc.Range2(at, at + 1).Text = "";
			}
			doc.Selection2.SetRange(at, at);
			try { Insert(editor, section); }
			catch (InvalidOperationException) { }    // (in a table: left out)
		}
		doc.Selection2.SetRange(0, 0);
	}

	/// The section break that starts the section a position is in (null: the first section, the document's own).
	public static Break? BreakFor(ITextDocument2 doc, int position) => Breaks(doc).LastOrDefault(b => b.SectionStart <= position);

	// ================================================================ RTF: Word's sections

	// groups whose \sect isn't the document's: headers, footers and the other destinations
	static readonly Regex Destination = new(@"\G\{\s*\\(\*|header[lrf]?|footer[lrf]?|fonttbl|colortbl|stylesheet|info|pict|listtable|listoverridetable|object|shp[a-z]*|nonshppict|themedata|datastore|latentstyles|xmlnstbl|rsidtbl|generator)(?![a-zA-Z])");

	// a section break's marker as RTF writes it (braces escaped; letters outside ASCII as \uN? or \'hh; Word puts font
	// words such as \hich\af44 in the middle of it)
	static readonly Regex RtfMark = new(@"\[\[sect:((?:[^\]\\]|\\[\\{}]|\\'[0-9a-fA-F]{2}|\\u-?\d+\??|\\[a-zA-Z]+-?\d* ?|\](?!\]))*)\]\]");

	// only RTF's own between two places: control words, braces, spaces
	static readonly Regex OnlyControls = new(@"^(?:\s|\\[a-zA-Z]+-?\d* ?|[{}])*$");

	// a paragraph mark with only RTF's own after it
	static readonly Regex EndsParagraph = new(@"\\par(?![a-zA-Z]) ?(?:\s|\\[a-zA-Z]+-?\d* ?|[{}])*$");

	// the break paragraph's mark after its \page (RichEdit writes "\page\par")
	static readonly Regex BreakMark = new(@"\G(?:\s|\\(?!par(?![a-zA-Z]))[a-zA-Z]+-?\d* ?)*?(\\par(?![a-zA-Z]) ?)");

	/// The \sect control words of an RTF document's text (not its headers', footers' or other destinations'): where each
	/// is and how long (with the space after it).
	static List<(int Start, int Length)> SectsIn(string rtf) {
		var found = new List<(int, int)>();
		int depth = 0, skipFrom = -1;
		for (int i = 0; i < rtf.Length; i++) {
			char c = rtf[i];
			if (c == '{') { depth++; if (skipFrom < 0 && Destination.IsMatch(rtf, i)) skipFrom = depth; continue; }
			if (c == '}') { if (skipFrom == depth) skipFrom = -1; depth--; continue; }
			if (c != '\\' || i + 1 >= rtf.Length) continue;
			if (!char.IsLetter(rtf[i + 1])) { i++; continue; }
			int j = i + 1;
			while (j < rtf.Length && char.IsLetter(rtf[j])) j++;
			if (skipFrom < 0 && j - i - 1 == 4 && string.CompareOrdinal(rtf, i + 1, "sect", 0, 4) == 0 && (j >= rtf.Length || (!char.IsDigit(rtf[j]) && rtf[j] != '-')))
				found.Add((i, (j < rtf.Length && rtf[j] == ' ' ? j + 1 : j) - i));
			i = j - 1;
		}
		return found;
	}

	// the header and footer groups in rtf from start (before end), as (where, how long)
	static List<(int Start, int Length)> BandGroupsIn(string rtf, int start, int end) {
		var found = new List<(int, int)>();
		end = Math.Min(end, rtf.Length);
		var group = new Regex(@"\{\s*\\(header|footer)[lrf]?(?![a-zA-Z])");
		int from = start;
		while (group.Match(rtf, from) is { Success: true } m && m.Index < end) {
			int depth = 0, i = m.Index;
			for (; i < rtf.Length; i++) {
				if (rtf[i] == '\\') { i++; continue; }
				if (rtf[i] == '{') depth++;
				else if (rtf[i] == '}' && --depth == 0) break;
			}
			found.Add((m.Index, Math.Min(rtf.Length, i + 1) - m.Index));
			from = i + 1;
		}
		return found;
	}

	static string RtfUnescape(string s) {
		s = Regex.Replace(s, @"\\\\|\\(?!u-?\d)[a-zA-Z]+-?\d* ?", m => m.Value == @"\\" ? m.Value : "");
		s = Regex.Replace(s, @"\\u(-?\d+)\??", m => ((char)(short)int.Parse(m.Groups[1].Value)).ToString());
		s = Regex.Replace(s, @"\\'([0-9a-fA-F]{2})", m => ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString());
		return Regex.Replace(s, @"\\([\\{}])", "$1");
	}

	static string RtfEscape(string text) {
		var sb = new System.Text.StringBuilder();
		foreach (char c in text) {
			if (c is '\\' or '{' or '}') sb.Append('\\').Append(c);
			else if (c > 126 || c < 32) sb.Append(@"\u").Append((short)c).Append('?');
			else sb.Append(c);
		}
		return sb.ToString();
	}

	/// RTF with Word's sections as AhkWord's: each \sect of its text a section break (a page break, its marker holding the
	/// section after it: the header and footer the file gives that section, or the section before's when it gives none, as
	/// Word links them; its page numbers starting again, \pgnrestart, at \pgnstarts). A marker AhkWord wrote in front of a
	/// \sect is its exact settings, but for its page setup, which the section after the \sect gives: another program may
	/// have changed it. What differs from the first section's is the section's own (SectionPage). A continuous section
	/// (\sbknone) starts no page: a continuous break, its hidden paragraph after the paragraph its \sect ends. (The first
	/// section's header and footer are the document's: HeaderFooter.FromRtf.)
	public static string ImportRtf(string rtf) {
		var sects = SectsIn(rtf);
		int body = HeaderFooter.BodyStart(rtf);
		if (sects.Count == 0 || body < 0) return rtf;
		string head = rtf[..body];
		var starts = new List<int> { body };
		starts.AddRange(sects.Select(s => s.Start + s.Length));
		var ends = sects.Select(s => s.Start).ToList();
		ends.Add(rtf.Length);
		// a section's header and footer, as its own groups give them (null: none)
		PageBands? BandsOf(int k) {
			string segment = rtf[starts[k]..Math.Max(starts[k], ends[k])];
			if (!Regex.IsMatch(segment, @"\{\s*\\(header|footer)[lrf]?(?![a-zA-Z])")) return null;
			return HeaderFooter.FromRtf(head + segment + "}");
		}
		// a section's own properties, before its header and footer and its text
		string PropertiesOf(int k) {
			string segment = rtf[starts[k]..Math.Max(starts[k], ends[k])];
			var cut = Regex.Match(segment, @"\{\s*\\(header|footer)|\\pard(?![a-zA-Z])");
			return cut.Success ? segment[..cut.Index] : segment[..Math.Min(segment.Length, 2000)];
		}
		bool Continuous(int k) => Regex.IsMatch(PropertiesOf(k), @"\\sbknone(?![a-zA-Z])");
		// a section's page setup: its own words, else the document's, else RTF's defaults
		SectionPage PageOf(int k) {
			string properties = PropertiesOf(k);
			int Twips(string own, string document, int fallback) {
				var m = Regex.Match(properties, @"\\" + own + @"(-?\d+)");
				if (!m.Success) m = Regex.Match(head, @"\\" + document + @"(-?\d+)");
				return m.Success ? int.Parse(m.Groups[1].Value) : fallback;
			}
			bool landscape = Regex.IsMatch(properties, @"\\lndscpsxn(?![a-zA-Z])") || (!properties.Contains(@"\pgwsxn") && Regex.IsMatch(head, @"\\landscape(?![a-zA-Z])"));
			return new SectionPage(Twips("pgwsxn", "paperw", 12240), Twips("pghsxn", "paperh", 15840), Twips("marglsxn", "margl", 1800), Twips("margrsxn", "margr", 1800),
				Twips("margtsxn", "margt", 1440), Twips("margbsxn", "margb", 1440), landscape, Twips("headery", "headery", 720), Twips("footery", "footery", 720));
		}
		var firstPage = PageOf(0);
		var result = new System.Text.StringBuilder(rtf);
		for (int k = sects.Count - 1; k >= 0; k--) {
			var (start, length) = sects[k];
			string segment = rtf[starts[k + 1]..Math.Max(starts[k + 1], ends[k + 1])];
			string properties = PropertiesOf(k + 1);
			var restart = Regex.IsMatch(properties, @"\\pgnrestart(?![a-zA-Z])");
			var at = Regex.Match(properties, @"\\pgnstarts(\d+)");
			var bands = BandsOf(k + 1);
			bool continuous = Continuous(k + 1);
			var page = PageOf(k + 1).Against(firstPage);
			var columns = SectionColumns.FromRtf(properties);
			var settings = new SectionSettings { Bands = bands, Same = bands == null, Restart = restart, Start = at.Success ? int.Parse(at.Groups[1].Value) : 1, Page = page, Columns = columns, Continuous = continuous };
			// (the section's own header and footer groups out of its text: they are its marker's now)
			if (BandsOf(k + 1) != null)
				foreach (var (groupStart, groupLength) in BandGroupsIn(rtf, starts[k + 1], starts[k + 1] + (Regex.Match(segment, @"\\pard(?![a-zA-Z])") is { Success: true } pard ? pard.Index : segment.Length)).AsEnumerable().Reverse())
					result.Remove(groupStart, groupLength);
			// AhkWord's own marker in front of the \sect: its exact settings (and taken out, to go in again as the break's)
			int from = Math.Max(0, start - 6000);
			var mark = RtfMark.Matches(rtf[from..start]).Cast<Match>().LastOrDefault();
			if (mark != null && OnlyControls.IsMatch(rtf[(from + mark.Index + mark.Length)..start]) && Parse(RtfUnescape(mark.Groups[1].Value)) is { } exact) {
				settings = exact;
				settings.Page = page;
				settings.Columns = columns;
				settings.Continuous = continuous;
				if (continuous) {
					result.Remove(start, length).Insert(start, @"\par " + HiddenParagraph(settings));
					result.Remove(from + mark.Index, mark.Length);
					continue;
				}
				result.Remove(start, length).Insert(start, @"\page\par ");
				result.Remove(from + mark.Index, mark.Length).Insert(from + mark.Index, RtfEscape(MarkerFor(settings)));
				continue;
			}
			// (Word's \sect ends its paragraph too: the break a paragraph of its own, as AhkWord's are)
			bool ended = EndsParagraph.IsMatch(rtf[Math.Max(0, start - 400)..start]);
			result.Remove(start, length).Insert(start, (ended ? "" : @"\par ") + (continuous ? HiddenParagraph(settings) : @"{\v " + RtfEscape(MarkerFor(settings)) + @"}\page\par "));
		}
		return result.ToString();
	}

	// a continuous break's paragraph
	static string HiddenParagraph(SectionSettings settings) => @"{\pard\v " + RtfEscape(MarkerFor(settings)) + @"\par }";

	/// RTF (AhkWord's own, with its header and footer written: HeaderFooter.AddToRtf) with its section breaks as Word's:
	/// each break's \page a \sect, the section after it given its settings (its own header and footer as \header and
	/// \footer groups, none when it is the same as the section before's; \pgnrestart \pgnstarts; the page layout, with what
	/// it has of its own). A continuous break's hidden paragraph goes: the \sect (\sbknone) ends the paragraph before it,
	/// with the marker at its end. The markers stay (hidden text) for AhkWord to read its own settings back.
	public static string ExportRtf(string rtf, PageLayout layout, string title, string file, float fontSize) {
		if (!rtf.Contains("[[sect:")) return rtf;
		// the sections' groups first (their fonts go into the font table, which moves what follows)
		var groups = new List<string>();
		foreach (Match m in RtfMark.Matches(rtf).Cast<Match>().ToList()) {
			var settings = Parse(RtfUnescape(m.Groups[1].Value));
			groups.Add(settings is { Same: false, Bands: { } bands } ? HeaderFooter.BandGroups(ref rtf, bands, layout, title, file, fontSize) : "");
		}
		var marks = RtfMark.Matches(rtf).Cast<Match>().ToList();
		for (int k = marks.Count - 1; k >= 0; k--) {
			var m = marks[k];
			if (Parse(RtfUnescape(m.Groups[1].Value)) is not { } settings) continue;
			int after = m.Index + m.Length;
			var page = Regex.Match(rtf[after..Math.Min(rtf.Length, after + 400)], settings.Continuous ? @"\\par(?![a-zA-Z]) ?" : @"\\page(?![a-zA-Z]) ?");
			if (!page.Success || !OnlyControls.IsMatch(rtf[after..(after + page.Index)])) continue;
			var sb = new System.Text.StringBuilder(@"\sect\sectd");
			if (settings.Continuous) sb.Append(@"\sbknone");
			var own = settings.Page ?? new SectionPage();
			if (own.Landscape ?? layout.Landscape) sb.Append(@"\lndscpsxn");
			if (settings.Restart) sb.Append(@"\pgnrestart\pgnstarts").Append(Math.Max(0, settings.Start));
			if (settings is { Same: false, Bands.DifferentFirst: true }) sb.Append(@"\titlepg");
			if (settings.Columns != null) sb.Append(settings.Columns.ToRtf());
			sb.Append($@"\pgwsxn{own.Width ?? layout.Width}\pghsxn{own.Height ?? layout.Height}\marglsxn{own.Left ?? layout.Left}\margrsxn{own.Right ?? layout.Right}\margtsxn{own.Top ?? layout.Top}\margbsxn{own.Bottom ?? layout.Bottom}\headery{own.HeaderY ?? layout.HeaderY}\footery{own.FooterY ?? layout.FooterY} ");
			sb.Append(k < groups.Count ? groups[k] : "");
			// (\sect ends the paragraph: the break paragraph's own mark would be an empty line in Word)
			int pageAt = after + page.Index;
			if (settings.Continuous) {
				int before = Math.Max(0, m.Index - 400);
				var ending = Regex.Matches(rtf[before..m.Index], @"\\par(?![a-zA-Z]) ?").Cast<Match>().LastOrDefault();
				if (ending == null || !OnlyControls.IsMatch(rtf[(before + ending.Index + ending.Length)..m.Index])) {    // (after a table: the \sect ends its own paragraph)
					rtf = rtf.Remove(pageAt, page.Length).Insert(pageAt, sb.ToString());
					continue;
				}
				rtf = rtf.Remove(pageAt, page.Length).Remove(m.Index, m.Length).Remove(before + ending.Index, ending.Length).Insert(before + ending.Index, @"{\v " + m.Value + "}" + sb);
				continue;
			}
			if (BreakMark.Match(rtf, pageAt + page.Length) is { Success: true } mark) rtf = rtf.Remove(mark.Groups[1].Index, mark.Groups[1].Length);
			rtf = rtf.Remove(pageAt, page.Length).Insert(pageAt, sb.ToString());
		}
		return rtf;
	}

	/// The first section's columns as RTF gives them (null: one).
	public static SectionColumns? FirstColumnsFromRtf(string rtf) {
		int body = HeaderFooter.BodyStart(rtf);
		if (body < 0) return null;
		string rest = rtf[body..];
		var cut = Regex.Match(rest, @"\{\s*\\(header|footer)|\\pard(?![a-zA-Z])");
		return SectionColumns.FromRtf(cut.Success ? rest[..cut.Index] : rest[..Math.Min(rest.Length, 2000)]);
	}

	/// A document's page setup as its RTF gives it (the first section's, else the document's): margins (1/100 inch: left,
	/// right, top, bottom) and whether its pages are landscape; null when it gives no margins.
	public static (int[] Margins, bool Landscape)? PageSetupFromRtf(string rtf) {
		int sect = SectsIn(rtf) is { Count: > 0 } s ? s[0].Start : rtf.Length;
		string firstSection = rtf[..sect];
		int? Twips(string word) {
			var m = Regex.Match(firstSection, @"\\" + word + @"sxn(-?\d+)");
			if (!m.Success) m = Regex.Match(firstSection, @"\\" + word + @"(-?\d+)");
			return m.Success ? int.Parse(m.Groups[1].Value) : null;
		}
		var (left, right, top, bottom) = (Twips("margl"), Twips("margr"), Twips("margt"), Twips("margb"));
		if (left == null && right == null && top == null && bottom == null) return null;
		static int Hundredths(int? twips, int fallback) => twips is int t ? (int)Math.Round(Math.Abs(t) / 14.4) : fallback;
		bool landscape = Regex.IsMatch(firstSection, @"\\(landscape|lndscpsxn)(?![a-zA-Z])");
		// (RTF's own default where one isn't given: 1.25 inch left and right, 1 inch top and bottom)
		return (new[] { Hundredths(left, 125), Hundredths(right, 125), Hundredths(top, 100), Hundredths(bottom, 100) }, landscape);
	}

	/// A .docx's page setup (its first section's: the first paragraph's w:sectPr, else the body's), as PageSetupFromRtf;
	/// null when it gives no margins or can't be read.
	public static (int[] Margins, bool Landscape)? PageSetupFromDocx(string path) {
		try {
			using var zip = System.IO.Compression.ZipFile.OpenRead(path);
			if (zip.GetEntry("word/document.xml") is not { } entry) return null;
			System.Xml.Linq.XDocument document;
			using (var stream = entry.Open()) document = System.Xml.Linq.XDocument.Load(stream);
			System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
			var section = document.Descendants(w + "sectPr").FirstOrDefault();
			if (section?.Element(w + "pgMar") is not { } margins) return null;
			int Hundredths(string side, int fallback) => int.TryParse((string?)margins.Attribute(w + side), out int twips) ? (int)Math.Round(Math.Abs(twips) / 14.4) : fallback;
			bool landscape = (string?)section.Element(w + "pgSz")?.Attribute(w + "orient") == "landscape";
			return (new[] { Hundredths("left", 100), Hundredths("right", 100), Hundredths("top", 100), Hundredths("bottom", 100) }, landscape);
		} catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException) {
			return null;
		}
	}
}
