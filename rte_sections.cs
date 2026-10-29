// Sections, for AhkWord (AhkWord.cs, through #load). Iteration 63.
//
// A section break is a page break whose paragraph also holds the new section's settings, as a hidden marker in front of
// the form feed: [[sect:{...}]] (hidden text, CFE_HIDDEN, RTF \v, which RichEdit keeps through editing, undo, copying and
// its RTF, as the cross-references' and footnotes' markers are kept; rte_crossrefs.cs). So a section's header, footer
// and page numbering travel with its break. The text before the first break is the first section, with the document's
// own header and footer (WriterForm.Bands), as before.
//   - SectionSettings: a section's own: its header and footer (PageBands; Same: as the section before's, Word's "Link to
//     Previous", until one is given), whether its page numbers start again (Restart, on unless turned off) and at what
//     number (Start).
//   - PageSection: a section as laid out: where its text starts, its header and footer, its numbering.
//   - Sections: the markers read and written; a section break put in; each page's section, number shown ({page}), the
//     section's page count ({sectionpages}) and whether it is the section's first page (its own first-page header).
using System.Text.Json;
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

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

	public SectionSettings Copy() => new() { Bands = Bands == null ? null : Sections.CopyBands(Bands), Same = Same, Restart = Restart, Start = Start };
}

/// A section as laid out: where its text starts (the first section at 0), its header and footer, its numbering.
public sealed record PageSection(int Start, PageBands Bands, bool Restart, int StartAt);

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

	/// The document's section breaks, in order: markers right in front of a page break's form feed (others are ignored).
	public static List<Break> Breaks(ITextDocument2 doc) => Breaks(Story(doc));

	public static List<Break> Breaks(string story) {
		var breaks = new List<Break>();
		foreach (Match m in Mark.Matches(story)) {
			int after = m.Index + m.Length;
			if (after >= story.Length || story[after] != PageBreaks.Break) continue;
			if (Parse(m.Groups[1].Value) is not { } settings) continue;
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
	/// break (its own header and footer, or the section before's).
	public static List<PageSection> Resolve(ITextDocument2 doc, PageBands first) => Resolve(Breaks(doc), first);

	public static List<PageSection> Resolve(List<Break> breaks, PageBands first) {
		var sections = new List<PageSection> { new(0, first, true, 1) };
		foreach (var b in breaks) {
			var bands = b.Settings.Same || b.Settings.Bands == null ? sections[^1].Bands : b.Settings.Bands;
			sections.Add(new PageSection(b.SectionStart, bands, b.Settings.Restart, Math.Max(0, b.Settings.Start)));
		}
		return sections;
	}

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

	/// A section break's settings changed (its marker written again, hidden).
	public static void Write(ITextDocument2 doc, Break b, SectionSettings settings) {
		string marker = MarkerFor(settings);
		var range = doc.Range2(b.MarkStart, b.MarkStart + b.MarkLength);
		range.Text = marker;
		range = doc.Range2(b.MarkStart, b.MarkStart + marker.Length);
		range.Font.Hidden = tomConstants.tomTrue;
	}

	/// The section break that starts the section a position is in (null: the first section, the document's own).
	public static Break? BreakFor(ITextDocument2 doc, int position) => Breaks(doc).LastOrDefault(b => b.SectionStart <= position);

	// ================================================================ RTF: Word's sections

	// groups whose \sect isn't the document's: headers, footers and the other destinations
	static readonly Regex Destination = new(@"\G\{\s*\\(\*|header[lrf]?|footer[lrf]?|fonttbl|colortbl|stylesheet|info|pict|listtable|listoverridetable|object|shp[a-z]*|nonshppict|themedata|datastore|latentstyles|xmlnstbl|rsidtbl|generator)(?![a-zA-Z])");

	// a section break's marker as RTF writes it (braces escaped; letters outside ASCII as \uN? or \'hh)
	static readonly Regex RtfMark = new(@"\[\[sect:((?:[^\]\\]|\\[\\{}]|\\'[0-9a-fA-F]{2}|\\u-?\d+\??|\](?!\]))*)\]\]");

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
	/// \sect is its exact settings. (The first section's header and footer are the document's: HeaderFooter.FromRtf.)
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
		var result = new System.Text.StringBuilder(rtf);
		for (int k = sects.Count - 1; k >= 0; k--) {
			var (start, length) = sects[k];
			string segment = rtf[starts[k + 1]..Math.Max(starts[k + 1], ends[k + 1])];
			var cut = Regex.Match(segment, @"\{\s*\\(header|footer)|\\pard(?![a-zA-Z])");
			string properties = cut.Success ? segment[..cut.Index] : segment[..Math.Min(segment.Length, 2000)];
			var restart = Regex.IsMatch(properties, @"\\pgnrestart(?![a-zA-Z])");
			var at = Regex.Match(properties, @"\\pgnstarts(\d+)");
			var bands = BandsOf(k + 1);
			var settings = new SectionSettings { Bands = bands, Same = bands == null, Restart = restart, Start = at.Success ? int.Parse(at.Groups[1].Value) : 1 };
			// (the section's own header and footer groups out of its text: they are its marker's now)
			if (bands != null)
				foreach (var (groupStart, groupLength) in BandGroupsIn(rtf, starts[k + 1], starts[k + 1] + (Regex.Match(segment, @"\\pard(?![a-zA-Z])") is { Success: true } pard ? pard.Index : segment.Length)).AsEnumerable().Reverse())
					result.Remove(groupStart, groupLength);
			// AhkWord's own marker in front of the \sect: its exact settings (and taken out, to go in again as the break's)
			int from = Math.Max(0, start - 6000);
			var mark = RtfMark.Matches(rtf[from..start]).Cast<Match>().LastOrDefault();
			if (mark != null && OnlyControls.IsMatch(rtf[(from + mark.Index + mark.Length)..start]) && Parse(RtfUnescape(mark.Groups[1].Value)) is { } exact) {
				settings = exact;
				result.Remove(start, length).Insert(start, @"\page\par ");
				result.Remove(from + mark.Index, mark.Length).Insert(from + mark.Index, RtfEscape(MarkerFor(settings)));
				continue;
			}
			// (Word's \sect ends its paragraph too: the break a paragraph of its own, as AhkWord's are)
			bool ended = EndsParagraph.IsMatch(rtf[Math.Max(0, start - 400)..start]);
			result.Remove(start, length).Insert(start, (ended ? "" : @"\par ") + @"{\v " + RtfEscape(MarkerFor(settings)) + @"}\page\par ");
		}
		return result.ToString();
	}

	/// RTF (AhkWord's own, with its header and footer written: HeaderFooter.AddToRtf) with its section breaks as Word's:
	/// each break's \page a \sect, the section after it given its settings (its own header and footer as \header and
	/// \footer groups, none when it is the same as the section before's; \pgnrestart \pgnstarts; the page layout). The
	/// markers stay (hidden text) for AhkWord to read its own settings back.
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
			var page = Regex.Match(rtf[after..Math.Min(rtf.Length, after + 400)], @"\\page(?![a-zA-Z]) ?");
			if (!page.Success || !OnlyControls.IsMatch(rtf[after..(after + page.Index)])) continue;
			var sb = new System.Text.StringBuilder(@"\sect\sectd");
			if (layout.Landscape) sb.Append(@"\lndscpsxn");
			if (settings.Restart) sb.Append(@"\pgnrestart\pgnstarts").Append(Math.Max(0, settings.Start));
			if (settings is { Same: false, Bands.DifferentFirst: true }) sb.Append(@"\titlepg");
			sb.Append($@"\pgwsxn{layout.Width}\pghsxn{layout.Height}\marglsxn{layout.Left}\margrsxn{layout.Right}\margtsxn{layout.Top}\margbsxn{layout.Bottom}\headery{layout.HeaderY}\footery{layout.FooterY} ");
			sb.Append(k < groups.Count ? groups[k] : "");
			// (\sect ends the paragraph: the break paragraph's own mark would be an empty line in Word)
			int pageAt = after + page.Index;
			if (BreakMark.Match(rtf, pageAt + page.Length) is { Success: true } mark) rtf = rtf.Remove(mark.Groups[1].Index, mark.Groups[1].Length);
			rtf = rtf.Remove(pageAt, page.Length).Insert(pageAt, sb.ToString());
		}
		return rtf;
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
