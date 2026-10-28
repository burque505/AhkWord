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
}
