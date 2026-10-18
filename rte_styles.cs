// Paragraph styles for RTE Writer (rte_writer32.cs, through #load).
//   - ParagraphStyle: a named set of formatting: the font (name, size, bold, italic, underline, colour) and the
//     paragraph's (alignment, indents, space before and after, line spacing, keep with next), a heading level (0 for
//     body text; 1 to 6 for headings: the outline, Markdown's #, HTML's h1-h6) and the style that follows it after Enter.
//   - StyleSheet: a document's styles; StyleSets: the three that come with the editor (Classic, the default; Plain, the
//     editor's look before styles; Modern), none of them Word's.
//   - StyleEngine: each paragraph carries its style in RichEdit's own paragraph style (TOM's ITextPara.Style: Normal -1,
//     Heading 1 to 6 -2 to -7, which RichEdit keeps in RTF itself; other styles numbers of their own). Applying a style
//     sets the paragraph's formatting, and the character formatting that is the style's (formatting given to part of the
//     paragraph, such as a bold word, stays); changing a style changes the paragraphs that have it in the same way.
//   - RtfStyles: the styles in RTF, as Word and LibreOffice read them (a style sheet, and each paragraph's \sN), with the
//     styles' own definitions kept for RTE Writer in {\*\rtewriterstyles ...}.
//   - StyleHtml: headings as h1-h6 in HTML, other styles as data-rte-style; read back (markers, as rte_linespacing2.cs).
//   - StyleDialog: Modify Style and New Style.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public sealed class ParagraphStyle
{
	public string Name { get; set; } = "Normal";
	public int Level { get; set; }                     // 0: body text; 1-6: a heading
	public string Next { get; set; } = "";             // the style after Enter at its end ("": the same)
	public string Font { get; set; } = "Times New Roman";
	public float Size { get; set; } = 12;
	public bool Bold { get; set; }
	public bool Italic { get; set; }
	public bool Underline { get; set; }
	public string Color { get; set; } = "";            // #RRGGBB ("": automatic)
	public string Align { get; set; } = "left";        // left, center, right, justify
	public float LeftIndent { get; set; }              // points
	public float RightIndent { get; set; }
	public float FirstLine { get; set; }               // points from the left indent (negative: hanging)
	public float Before { get; set; }                  // space before and after, points
	public float After { get; set; }
	public string LineSpacing { get; set; } = "1";     // as Rte.SetLineSpacing: "1.15", "exactly 14pt" ...
	public bool KeepWithNext { get; set; }

	public ParagraphStyle Copy() => (ParagraphStyle)MemberwiseClone();

	/// The styles whose names are fixed (and their levels): Normal, Heading 1 to 6.
	public static bool BuiltIn(string name) => string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase) || HeadingNumber(name) > 0;

	/// 1-6 for "Heading 1" to "Heading 6" (also Word's "heading 1"), else 0.
	public static int HeadingNumber(string name) {
		var m = Regex.Match(name.Trim(), @"^heading\s*([1-6])$", RegexOptions.IgnoreCase);
		return m.Success ? int.Parse(m.Groups[1].Value) : 0;
	}
}

public sealed class StyleSheet
{
	public string Set { get; set; } = "Classic";
	public List<ParagraphStyle> Styles { get; set; } = new();

	public StyleSheet Copy() => new() { Set = Set, Styles = Styles.Select(s => s.Copy()).ToList() };

	public ParagraphStyle? Find(string name) => Styles.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));

	public ParagraphStyle Normal => Find("Normal") ?? StyleSets.Get("Classic").Find("Normal")!;

	/// The style a paragraph's TOM style number stands for (null: one this sheet hasn't got).
	public ParagraphStyle? ById(int id) => StyleEngine.NameOf(id) is { } name ? Find(name) : null;

	/// A style by name, put in the sheet if it isn't (from the set's own, or a copy of Normal).
	public ParagraphStyle Ensure(string name) {
		if (Find(name) is { } found) return found;
		var style = StyleSets.Get(Set).Find(name)?.Copy() ?? StyleSets.Get("Classic").Find(name)?.Copy() ?? Normal.Copy();
		style.Name = ParagraphStyle.HeadingNumber(name) is int n and > 0 ? "Heading " + n : name;
		style.Level = ParagraphStyle.HeadingNumber(style.Name);
		if (style.Level == 0 && !string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase) && StyleSets.Get("Classic").Find(name) == null) style.Next = "";
		Styles.Add(style);
		return style;
	}
}

public static class StyleSets
{
	public static readonly string[] Names = { "Classic", "Plain", "Modern" };

	/// A set by name (Classic if unknown), new each time.
	public static StyleSheet Get(string name) => name.Trim().ToLowerInvariant() switch {
		"plain" => Plain(),
		"modern" => Modern(),
		_ => Classic()
	};

	static ParagraphStyle S(string name, string font, float size, int level = 0, bool bold = false, bool italic = false, bool underline = false,
		string align = "left", float before = 0, float after = 0, float left = 0, float right = 0, float first = 0, string line = "1",
		bool keep = false, string next = "", string color = "") =>
		new() { Name = name, Font = font, Size = size, Level = level, Bold = bold, Italic = italic, Underline = underline, Align = align,
			Before = before, After = after, LeftIndent = left, RightIndent = right, FirstLine = first, LineSpacing = line, KeepWithNext = keep, Next = next, Color = color };

	// Classic: Times New Roman 12, black headings in the same face, the conventions of court papers and contracts.
	static StyleSheet Classic() {
		const string F = "Times New Roman";
		return new StyleSheet { Set = "Classic", Styles = {
			S("Normal", F, 12),
			S("Title", F, 16, bold: true, align: "center", after: 12, keep: true, next: "Normal"),
			S("Subtitle", F, 12, italic: true, align: "center", after: 12, keep: true, next: "Normal"),
			S("Heading 1", F, 14, 1, bold: true, before: 12, after: 6, keep: true, next: "Normal"),
			S("Heading 2", F, 12, 2, bold: true, before: 12, after: 6, keep: true, next: "Normal"),
			S("Heading 3", F, 12, 3, bold: true, italic: true, before: 6, after: 3, keep: true, next: "Normal"),
			S("Heading 4", F, 12, 4, italic: true, before: 6, after: 3, keep: true, next: "Normal"),
			S("Heading 5", F, 12, 5, underline: true, before: 6, after: 3, keep: true, next: "Normal"),
			S("Heading 6", F, 12, 6, italic: true, underline: true, before: 6, after: 3, keep: true, next: "Normal"),
			S("Quote", F, 12, left: 36, right: 36, after: 12),
			S("Code", "Courier New", 10),
		} };
	}

	// Plain: the editor's look before styles: Segoe UI 11, bold headings of 24 to 8 points.
	static StyleSheet Plain() {
		const string F = "Segoe UI";
		var sizes = new[] { 0f, 24f, 18f, 14f, 12f, 10f, 8f };
		var sheet = new StyleSheet { Set = "Plain", Styles = { S("Normal", F, 11), S("Title", F, 28, align: "center", after: 12, next: "Normal"), S("Subtitle", F, 14, align: "center", after: 12, next: "Normal") } };
		for (int level = 1; level <= 6; level++) sheet.Styles.Add(S("Heading " + level, F, sizes[level], level, bold: true, next: "Normal"));
		sheet.Styles.Add(S("Quote", F, 11, italic: true, left: 36));
		sheet.Styles.Add(S("Code", "Courier New", 10));
		return sheet;
	}

	// Modern: Georgia text with Segoe UI headings in dark grey, a little more air between lines.
	static StyleSheet Modern() {
		const string T = "Georgia", H = "Segoe UI Semibold";
		return new StyleSheet { Set = "Modern", Styles = {
			S("Normal", T, 11, line: "1.15", after: 8),
			S("Title", "Segoe UI Light", 26, after: 12, keep: true, next: "Normal", color: "#262626"),
			S("Subtitle", "Segoe UI", 13, after: 12, keep: true, next: "Normal", color: "#595959"),
			S("Heading 1", H, 18, 1, before: 18, after: 6, keep: true, next: "Normal", color: "#262626"),
			S("Heading 2", H, 15, 2, before: 14, after: 4, keep: true, next: "Normal", color: "#262626"),
			S("Heading 3", H, 13, 3, before: 12, after: 4, keep: true, next: "Normal", color: "#262626"),
			S("Heading 4", H, 11, 4, before: 10, after: 2, keep: true, next: "Normal", color: "#262626"),
			S("Heading 5", "Segoe UI", 11, 5, italic: true, before: 10, after: 2, keep: true, next: "Normal", color: "#404040"),
			S("Heading 6", "Segoe UI", 10, 6, italic: true, before: 8, after: 2, keep: true, next: "Normal", color: "#404040"),
			S("Quote", T, 11, italic: true, left: 36, right: 36, line: "1.15", after: 8, color: "#404040"),
			S("Code", "Consolas", 10),
		} };
	}
}

public static class StyleEngine
{
	const int CharFormat = 13, Paragraph = 4, AutoColor = -9999997;
	static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

	// The TOM style numbers of the styles that aren't RichEdit's own (Normal, Heading 1-9), by name, for the session.
	static readonly Dictionary<string, int> custom = new(StringComparer.OrdinalIgnoreCase);

	/// The TOM paragraph style number of a style: Normal -1, Heading 1 to 6 -2 to -7, others 100 and on.
	public static int IdOf(string name) {
		if (string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase)) return -1;
		if (ParagraphStyle.HeadingNumber(name) is int n and > 0) return -(n + 1);
		lock (custom) {
			if (!custom.TryGetValue(name, out int id)) custom[name] = id = 100 + custom.Count;
			return id;
		}
	}

	/// The style name a TOM paragraph style number stands for (null: none RTE Writer gave).
	public static string? NameOf(int id) {
		if (id == -1 || id == 0) return "Normal";
		if (id is <= -2 and >= -7) return "Heading " + (-id - 1);
		lock (custom) return custom.FirstOrDefault(c => c.Value == id).Key;
	}

	/// Gives a paragraph its style (a TOM style number) and nothing else. RichEdit, given one of its own styles (Normal,
	/// the headings), also gives the paragraph that style's paragraph formatting (indents, spacing, alignment ... back
	/// to its defaults), even given the style it has already: the paragraph's formatting is put back.
	public static void Tag(ITextPara para, int id) {
		if (para.Style == id) return;
		int align = para.Alignment, rule = para.LineSpacingRule, keepWith = para.KeepWithNext, keepTogether = para.KeepTogether,
			widow = para.WidowControl, breakBefore = para.PageBreakBefore, noNumber = para.NoLineNumber;
		float first = para.FirstLineIndent, left = para.LeftIndent, right = para.RightIndent, before = para.SpaceBefore, after = para.SpaceAfter, line = para.LineSpacing;
		int listType = para.ListType, listLevel = para.ListLevelIndex, listStart = para.ListStart, listAlign = para.ListAlignment;
		float listTab = para.ListTab;
		var tabs = new List<(float, int, int)>();
		for (int i = 0; i < para.TabCount; i++) { para.GetTab(i, out float position, out int alignment, out int leader); tabs.Add((position, alignment, leader)); }
		para.Style = id;
		if (para.Alignment != align) para.Alignment = align;
		para.SetIndents(first, left, right);
		para.SpaceBefore = before;
		para.SpaceAfter = after;
		if (rule >= 0) para.SetLineSpacing(rule, line);
		para.KeepWithNext = keepWith;
		para.KeepTogether = keepTogether;
		para.WidowControl = widow;
		para.PageBreakBefore = breakBefore;
		para.NoLineNumber = noNumber;
		if (para.ListType != listType) para.ListType = listType;
		if (listType != tomConstants.tomListNone) {
			para.ListLevelIndex = listLevel;
			para.ListStart = listStart;
			para.ListAlignment = listAlign;
			para.ListTab = listTab;
		}
		para.ClearAllTabs();
		foreach (var (position, alignment, leader) in tabs) para.AddTab(position, alignment, leader);
	}

	/// The paragraphs (start, end, with the mark) that a range touches.
	public static List<(int Start, int End)> Paragraphs(ITextDocument2 doc, int start, int end) {
		var list = new List<(int, int)>();
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var range = doc.Range2(start, start);
		range.Expand(Paragraph);
		while (true) {
			list.Add((range.Start, range.End));
			if (range.End >= storyEnd || range.End >= Math.Max(end, start + 1) || range.End <= range.Start) break;
			range.SetRange(range.End, range.End);
			range.Expand(Paragraph);
		}
		return list;
	}

	/// The style of the paragraph at a position (Normal if it has none of the sheet's).
	public static ParagraphStyle StyleAt(ITextDocument2 doc, StyleSheet sheet, int position) =>
		sheet.ById(doc.Range2(position, position).Para.Style) ?? sheet.Normal;

	/// The heading level of the paragraph at a position (0: body text).
	public static int LevelAt(ITextDocument2 doc, StyleSheet sheet, int position) {
		int id = doc.Range2(position, position).Para.Style;
		if (id is <= -2 and >= -7) return -id - 1;
		return sheet.ById(id)?.Level ?? 0;
	}

	// The character runs (same formatting) from start to end.
	static List<ITextRange2> Runs(ITextDocument2 doc, int start, int end) {
		var runs = new List<ITextRange2>();
		int position = start;
		while (position < end) {
			var run = doc.Range2(position, position);
			run.MoveEnd(CharFormat, 1);
			int stop = Math.Min(run.End, end);
			if (stop <= position) break;
			runs.Add(doc.Range2(position, stop));
			position = stop;
		}
		return runs;
	}

	static bool Same(float a, float b) => Math.Abs(a - b) < 0.05f;
	static bool Is(int tom) => tom == tomConstants.tomTrue;
	static int Tom(bool on) => on ? tomConstants.tomTrue : tomConstants.tomFalse;
	static int ColorOf(string color) => color is { Length: 7 } && color[0] == '#' && int.TryParse(color[1..], NumberStyles.HexNumber, Invariant, out int rgb)
		? ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF) : AutoColor;    // (TOM's colour: 0x00BBGGRR)
	static string ColorText(int tom) => tom == AutoColor || tom < 0 ? "" : $"#{tom & 0xFF:X2}{(tom >> 8) & 0xFF:X2}{(tom >> 16) & 0xFF:X2}";
	static int AlignOf(string align) => align switch { "center" => 1, "right" => 2, "justify" => 3, _ => 0 };
	static string AlignText(int tom) => tom switch { 1 => "center", 2 => "right", 3 => "justify", _ => "left" };

	/// Applies a style to the paragraphs from start to end: their paragraph formatting becomes the style's, and so does
	/// the character formatting that was their style's or covers a whole paragraph (part of a paragraph formatted
	/// otherwise, such as a bold word, stays as it is).
	public static void Apply(ITextDocument2 doc, StyleSheet sheet, ParagraphStyle style, int start, int end) {
		foreach (var (pStart, pEnd) in Paragraphs(doc, start, end)) {
			var para = doc.Range2(pStart, pStart).Para;
			var old = sheet.ById(para.Style) ?? sheet.Normal;
			Tag(para, IdOf(style.Name));
			SetParagraph(para, style, keepIndents: (para.ListType & 0xFFFF) != tomConstants.tomListNone);    // (a list item keeps its list's indents)
			SetCharacters(doc, pStart, pEnd, old, style, whole: true);
		}
	}

	/// Gives the paragraph at a position a style, as Apply, but only its character formatting (its indents and spacing
	/// stay: a list item's, a Markdown quote's).
	public static void ApplyCharacters(ITextDocument2 doc, StyleSheet sheet, ParagraphStyle style, int position) {
		var (start, end) = Paragraphs(doc, position, position)[0];
		var para = doc.Range2(start, start).Para;
		var old = sheet.ById(para.Style) ?? sheet.Normal;
		Tag(para, IdOf(style.Name));
		SetCharacters(doc, start, end, old, style, whole: true);
	}

	/// After a style changed (old to updated), the paragraphs that have it: what was the style's becomes the new style's;
	/// formatting of their own stays.
	public static int Restyle(ITextDocument2 doc, StyleSheet sheet, ParagraphStyle old, ParagraphStyle updated) {
		int id = IdOf(old.Name), count = 0;
		foreach (var (pStart, pEnd) in Paragraphs(doc, 0, doc.Range2(0, 0).StoryLength)) {
			var para = doc.Range2(pStart, pStart).Para;
			if (para.Style != id) continue;
			if (!string.Equals(old.Name, updated.Name, StringComparison.OrdinalIgnoreCase)) Tag(para, IdOf(updated.Name));
			if ((AlignText(para.Alignment)) == old.Align) para.Alignment = AlignOf(updated.Align);
			if (Same(para.SpaceBefore, old.Before)) para.SpaceBefore = updated.Before;
			if (Same(para.SpaceAfter, old.After)) para.SpaceAfter = updated.After;
			if (Same(para.LeftIndent, old.LeftIndent) && Same(para.FirstLineIndent, old.FirstLine) && Same(para.RightIndent, old.RightIndent))
				para.SetIndents(updated.FirstLine, updated.LeftIndent, updated.RightIndent);
			if (global::LineSpacing.Of(doc.Range2(pStart, pStart)) == Normalize(old.LineSpacing)) SetLine(para, updated.LineSpacing);
			if (Is(para.KeepWithNext) == old.KeepWithNext) para.KeepWithNext = Tom(updated.KeepWithNext);
			SetCharacters(doc, pStart, pEnd, old, updated, whole: false);
			count++;
		}
		return count;
	}

	static string Normalize(string spacing) {
		try { var (rule, value) = global::LineSpacing.Parse(spacing); return global::LineSpacing.Describe(rule, value); } catch (ArgumentException) { return "1"; }
	}

	static void SetLine(ITextPara para, string spacing) {
		var (rule, value) = global::LineSpacing.Parse(spacing.Trim() == "" ? "1" : spacing);
		para.SetLineSpacing(rule, rule == global::LineSpacing.Single ? 1 : value);
	}

	static void SetParagraph(ITextPara para, ParagraphStyle style, bool keepIndents = false) {
		para.Alignment = AlignOf(style.Align);
		if (!keepIndents) para.SetIndents(style.FirstLine, style.LeftIndent, style.RightIndent);
		para.SpaceBefore = style.Before;
		para.SpaceAfter = style.After;
		SetLine(para, style.LineSpacing);
		para.KeepWithNext = Tom(style.KeepWithNext);
	}

	// The paragraph's characters from the old style's formatting to the new's: what was the old style's becomes the new
	// style's; formatting of their own stays, unless (whole: a style applied, as Word does) it covers more than half the
	// paragraph's text, a pasted line's look, say (a bold word stays bold).
	static void SetCharacters(ITextDocument2 doc, int start, int end, ParagraphStyle old, ParagraphStyle style, bool whole) {
		var runs = Runs(doc, start, end);
		if (runs.Count == 0) return;
		var fonts = runs.Select(r => (Font: r.Font, Length: Math.Max(1, r.End - r.Start))).ToList();
		int total = fonts.Sum(f => f.Length);
		int oldColor = ColorOf(old.Color), newColor = ColorOf(style.Color);
		bool Over(Func<ITextFont, bool> own) => whole && fonts.Where(f => own(f.Font)).Sum(f => f.Length) * 2 > total;
		bool OwnName(ITextFont f) => !string.Equals(f.Name, old.Font, StringComparison.OrdinalIgnoreCase);
		bool OwnSize(ITextFont f) => !Same(f.Size, old.Size);
		bool OwnBold(ITextFont f) => Is(f.Bold) != old.Bold;
		bool OwnItalic(ITextFont f) => Is(f.Italic) != old.Italic;
		bool OwnUnderline(ITextFont f) => (f.Underline != 0) != old.Underline;
		bool OwnColor(ITextFont f) => !(f.ForeColor == oldColor || (oldColor == AutoColor && f.ForeColor < 0));
		bool name = Over(OwnName), size = Over(OwnSize), bold = Over(OwnBold), italic = Over(OwnItalic), underline = Over(OwnUnderline), color = Over(OwnColor);
		foreach (var (font, _) in fonts) {
			if (name || !OwnName(font)) font.Name = style.Font;
			if (size || !OwnSize(font)) font.Size = style.Size;
			if (bold || !OwnBold(font)) font.Bold = Tom(style.Bold);
			if (italic || !OwnItalic(font)) font.Italic = Tom(style.Italic);
			if (underline || !OwnUnderline(font)) font.Underline = style.Underline ? tomConstants.tomSingle : tomConstants.tomNone;
			if (color || !OwnColor(font)) font.ForeColor = newColor;
		}
	}

	/// Clear Formatting: the text from start to end back to its paragraphs' styles (all of its formatting of its own
	/// goes); with paragraphs (the caret, no selection), also their paragraph formatting.
	public static void ResetFormatting(ITextDocument2 doc, StyleSheet sheet, int start, int end, bool paragraphs) {
		foreach (var (pStart, pEnd) in Paragraphs(doc, start, end)) {
			var para = doc.Range2(pStart, pStart).Para;
			var style = sheet.ById(para.Style) ?? sheet.Normal;
			if (paragraphs) SetParagraph(para, style, keepIndents: (para.ListType & 0xFFFF) != tomConstants.tomListNone);
			int from = paragraphs ? pStart : Math.Max(pStart, start), to = paragraphs ? pEnd : Math.Min(pEnd, Math.Max(end, start + 1));
			if (to > from) SetTyping(doc.Range2(from, to).Font, style);
		}
	}
	/// What is typed next at the caret (in an empty paragraph) gets the style's character formatting.
	public static void SetTyping(ITextFont font, ParagraphStyle style) {
		font.Name = style.Font;
		font.Size = style.Size;
		font.Bold = Tom(style.Bold);
		font.Italic = Tom(style.Italic);
		font.Underline = style.Underline ? tomConstants.tomSingle : tomConstants.tomNone;
		font.ForeColor = ColorOf(style.Color);
	}

	/// The formatting of the paragraph at a position, as a style named name (with level and next as given).
	public static ParagraphStyle Capture(ITextDocument2 doc, int position, string name, int level, string next) {
		var (start, end) = Paragraphs(doc, position, position)[0];
		var para = doc.Range2(start, start).Para;
		int at = start;
		string text = doc.Range2(start, end).Text ?? "";
		while (at < end - 1 && text[at - start] is ' ' or '\t' or '￼') at++;    // (the first character with text's formatting)
		var font = doc.Range2(at, at + 1).Font;
		return new ParagraphStyle {
			Name = name, Level = level, Next = next, Font = font.Name ?? "Times New Roman", Size = (float)Math.Round(font.Size, 1),
			Bold = Is(font.Bold), Italic = Is(font.Italic), Underline = font.Underline > 0, Color = ColorText(font.ForeColor),
			Align = AlignText(para.Alignment), LeftIndent = para.LeftIndent, RightIndent = para.RightIndent, FirstLine = para.FirstLineIndent,
			Before = Math.Max(0, para.SpaceBefore), After = Math.Max(0, para.SpaceAfter),
			LineSpacing = global::LineSpacing.Of(doc.Range2(start, start)) is var line && line != "mixed" ? line : "1",
			KeepWithNext = Is(para.KeepWithNext)
		};
	}

	/// The paragraphs' style names, in order as RtfStyles.Inject counts them (paragraph marks and cell ends, without the
	/// marks that start and end table rows).
	public static List<string> Names(ITextDocument2 doc, StyleSheet sheet) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var names = new List<string>();
		int start = 0;
		for (int i = 0; i < story.Length; i++) {
			if (story[i] is not ('\r' or '\u0007')) continue;
			string piece = story[start..i];
			if (!(story[i] == '\r' && piece.Length > 0 && piece.Trim('￹', '￻') == ""))
				names.Add(sheet.ById(doc.Range2(start, start).Para.Style)?.Name ?? "Normal");
			start = i + 1;
		}
		return names;
	}

	/// Gives the paragraphs (counted as Names counts them) the styles named; returns false if the counts don't match.
	public static bool SetNames(ITextDocument2 doc, IReadOnlyList<string> names) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var starts = new List<int>();
		int start = 0;
		for (int i = 0; i < story.Length; i++) {
			if (story[i] is not ('\r' or '\u0007')) continue;
			string piece = story[start..i];
			if (!(story[i] == '\r' && piece.Length > 0 && piece.Trim('￹', '￻') == "")) starts.Add(start);
			start = i + 1;
		}
		if (starts.Count < names.Count || starts.Count > names.Count + 1) return false;    // (RichEdit's last empty paragraph)
		for (int i = 0; i < names.Count; i++) Tag(doc.Range2(starts[i], starts[i]).Para, IdOf(names[i]));
		return true;
	}
}

// Styles in RTF. Word, LibreOffice and RichEdit read a style sheet ({\stylesheet{\s0 Normal;}{\s1 heading 1;}...}) and
// each paragraph's style (\sN); RichEdit keeps only its own (Normal, the headings) and drops the others, so RTE Writer
// writes them itself, and keeps the styles' definitions for itself in {\*\rtewriterstyles base64-JSON} (others ignore it).
public static class RtfStyles
{
	// Destinations whose text is not the document's (their paragraphs aren't counted)
	static readonly HashSet<string> NotText = new() {
		"fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "nonshppict", "header", "headerl", "headerr", "headerf",
		"footer", "footerl", "footerr", "footerf", "footnote", "pntext", "pntxta", "pntxtb", "listtext", "listtable",
		"listoverridetable", "xe", "tc", "rsidtbl", "generator", "fldinst", "themedata", "colorschememapping", "latentstyles"
	};

	// The style sheet's numbers: Normal 0, Heading 1-6 1-6, the others from 10 on, in the sheet's order.
	static Dictionary<string, int> Numbers(StyleSheet sheet) {
		var numbers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { ["Normal"] = 0 };
		for (int n = 1; n <= 6; n++) numbers["Heading " + n] = n;
		int next = 10;
		foreach (var style in sheet.Styles) if (!numbers.ContainsKey(style.Name)) numbers[style.Name] = next++;
		return numbers;
	}

	// (Word's own names for its built-in styles are "heading 1" ...)
	static string RtfName(string name) => ParagraphStyle.HeadingNumber(name) is int n and > 0 ? "heading " + n : name;

	static string Escape(string text) {
		var sb = new StringBuilder();
		foreach (char c in text) {
			if (c is '\\' or '{' or '}') sb.Append('\\').Append(c);
			else if (c > 127) sb.Append("\\u").Append((short)c).Append('?');
			else sb.Append(c);
		}
		return sb.ToString();
	}

	/// The RTF (RichEdit's) with the styles: the style sheet (each style's whole definition, so that Word and LibreOffice
	/// have them as RTE Writer has them, and RTE Writer's own copy), and each paragraph's \sN at its start, before its
	/// own formatting (which a reader that applies the style's then puts back). names: the paragraphs' styles in order
	/// (StyleEngine.Names). If the paragraphs don't match in number, only the style sheet is written.
	public static string Inject(string rtf, IReadOnlyList<string> names, StyleSheet sheet) {
		var numbers = Numbers(sheet);
		// where each paragraph starts and ends in the body, and its \pard (its own formatting begins after the last one)
		var marks = new List<(string Word, int At, int After)>();
		Scan(rtf, (word, at, _, after) => { if (word is "par" or "cell" or "pard") marks.Add((word, at, after)); });
		var inserts = new List<(int At, string Text)>();
		int paragraph = 0, start = -1, lastPard = -1;
		foreach (var (word, at, after) in marks) {
			if (word == "pard") { lastPard = after; continue; }
			// a paragraph's end: its style after its last \pard, or at its start
			if (paragraph < names.Count) {
				int place = lastPard > start ? lastPard : start >= 0 ? start : -1;
				if (place >= 0) inserts.Add((place, "\\s" + (numbers.TryGetValue(names[paragraph], out int n) ? n : 0) + " "));
			}
			paragraph++;
			start = after;
			lastPard = -1;
		}
		string result = rtf;
		if ((paragraph == names.Count || paragraph == names.Count + 1) && inserts.Count >= names.Count - 1)
			foreach (var (at, text) in inserts.OrderByDescending(i => i.At)) result = result.Insert(at, text);
		// the style sheet: RichEdit's out, this one in, after the font and colour tables (with the styles' fonts and
		// colours in them)
		result = RemoveGroup(result, "{\\stylesheet");
		result = RemoveGroup(result, "{\\*\\rtewriterstyles");
		var sheetRtf = new StringBuilder("{\\stylesheet");
		foreach (var (name, number) in numbers.OrderBy(p => p.Value)) {
			var style = sheet.Find(name);
			if (style == null && number > 6) continue;
			string next = style?.Next is { Length: > 0 } nx && numbers.TryGetValue(nx, out int nn) ? $"\\snext{nn}" : $"\\snext{number}";
			sheetRtf.Append("{\\s").Append(number).Append(number == 0 ? "" : "\\sbasedon0").Append(next);
			if (style != null) sheetRtf.Append(Definition(ref result, style));
			sheetRtf.Append(' ').Append(Escape(RtfName(name))).Append(";}");
		}
		sheetRtf.Append('}');
		string json = JsonSerializer.Serialize(sheet);
		sheetRtf.Append("{\\*\\rtewriterstyles ").Append(Convert.ToBase64String(Encoding.UTF8.GetBytes(json))).Append('}');
		int place2 = AfterTables(result);
		return result.Insert(place2, sheetRtf.ToString());
	}

	// A style's formatting as RTF control words (its font and colour put into the RTF's tables if they aren't there).
	static string Definition(ref string rtf, ParagraphStyle style) {
		static int Tw(float points) => (int)Math.Round(points * 20);
		var sb = new StringBuilder();
		sb.Append(style.Align switch { "center" => "\\qc", "right" => "\\qr", "justify" => "\\qj", _ => "\\ql" });
		sb.Append($"\\li{Tw(style.LeftIndent)}\\ri{Tw(style.RightIndent)}\\fi{Tw(style.FirstLine)}\\sb{Tw(style.Before)}\\sa{Tw(style.After)}");
		try {
			var (rule, value) = global::LineSpacing.Parse(style.LineSpacing.Trim() == "" ? "1" : style.LineSpacing);
			sb.Append(rule switch {
				global::LineSpacing.Exactly => $"\\sl-{Tw(value)}\\slmult0",
				global::LineSpacing.AtLeast => $"\\sl{Tw(value)}\\slmult0",
				global::LineSpacing.Multiple => $"\\sl{(int)Math.Round(value * 240)}\\slmult1",
				_ => ""
			});
		} catch (ArgumentException) { }
		if (style.KeepWithNext) sb.Append("\\keepn");
		sb.Append("\\f").Append(FontNumber(ref rtf, style.Font)).Append("\\fs").Append((int)Math.Round(style.Size * 2));
		if (style.Bold) sb.Append("\\b");
		if (style.Italic) sb.Append("\\i");
		if (style.Underline) sb.Append("\\ul");
		if (style.Color is { Length: 7 } color) sb.Append("\\cf").Append(ColorNumber(ref rtf, color));
		return sb.ToString();
	}

	// The number of a font in the RTF's font table, added to it if it isn't there.
	static int FontNumber(ref string rtf, string font) {
		int table = rtf.IndexOf("{\\fonttbl", StringComparison.Ordinal);
		if (table < 0) return 0;
		int end = GroupEnd(rtf, table);
		var fonts = Regex.Matches(rtf[table..end], @"\{\\f(\d+)[^{};]*?(?:\{[^{}]*\}[^{};]*?)*\s?([^\\{};]+);\}");
		foreach (Match m in fonts)
			if (string.Equals(m.Groups[2].Value.Trim(), font, StringComparison.OrdinalIgnoreCase)) return int.Parse(m.Groups[1].Value);
		int number = fonts.Count == 0 ? 0 : fonts.Max(m => int.Parse(m.Groups[1].Value)) + 1;
		rtf = rtf.Insert(end - 1, "{\\f" + number + "\\fnil\\fcharset0 " + Escape(font) + ";}");
		return number;
	}

	// The number of a colour (#RRGGBB) in the RTF's colour table, added to it (or the table made) if it isn't there.
	static int ColorNumber(ref string rtf, string color) {
		int rgb = int.Parse(color[1..], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
		string entry = $"\\red{(rgb >> 16) & 0xFF}\\green{(rgb >> 8) & 0xFF}\\blue{rgb & 0xFF};";
		int table = rtf.IndexOf("{\\colortbl", StringComparison.Ordinal);
		if (table < 0) {
			int fonts = rtf.IndexOf("{\\fonttbl", StringComparison.Ordinal);
			int at = fonts >= 0 ? GroupEnd(rtf, fonts) : AfterTables(rtf);
			rtf = rtf.Insert(at, "{\\colortbl ;" + entry + "}");
			return 1;
		}
		int end = GroupEnd(rtf, table);
		string body = rtf[(table + "{\\colortbl".Length)..(end - 1)];
		var entries = body.Split(';');    // (the first, empty: the automatic colour; the last after the final ;)
		for (int i = 0; i < entries.Length - 1; i++)
			if (entries[i].Replace(" ", "") == entry.TrimEnd(';')) return i;
		rtf = rtf.Insert(end - 1, entry);
		return entries.Length - 1;
	}
	/// What an RTF says about styles: RTE Writer's definitions (null if it has none), the style sheet's names by number,
	/// and each paragraph's style number (in order, as Inject counts them). Any: whether it has styles at all.
	public static (StyleSheet? Sheet, Dictionary<int, string> Names, List<int> Paragraphs, bool Any) Read(string rtf) {
		StyleSheet? sheet = null;
		var own = Regex.Match(rtf, @"\{\\\*\\rtewriterstyles\s+([A-Za-z0-9+/=]+)\}");
		if (own.Success) try { sheet = JsonSerializer.Deserialize<StyleSheet>(Encoding.UTF8.GetString(Convert.FromBase64String(own.Groups[1].Value))); } catch (Exception ex) when (ex is FormatException or JsonException) { }
		var names = new Dictionary<int, string>();
		int sheetAt = rtf.IndexOf("{\\stylesheet", StringComparison.Ordinal);
		if (sheetAt >= 0) {
			int end = GroupEnd(rtf, sheetAt);
			foreach (Match m in Regex.Matches(rtf[sheetAt..end], @"\{(?:\\\*)?\\s(\d+)\b[^{};]*?(?:\{[^{}]*\}[^{};]*?)*\s([^\\{};][^{};]*?);\}"))
				names[int.Parse(m.Groups[1].Value)] = Unescape(m.Groups[2].Value.Trim());
		}
		var paragraphs = ParagraphStyles(rtf);
		return (sheet, names, paragraphs, sheet != null || names.Count > 1 || paragraphs.Any(p => p > 0));
	}

	static string Unescape(string text) => Regex.Replace(text, @"\\'([0-9a-fA-F]{2})|\\u(-?\d+)\??", m =>
		m.Groups[1].Success ? ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString() : ((char)(short)int.Parse(m.Groups[2].Value)).ToString());

	// The body's paragraph ends (\par and \cell outside other destinations): their positions.
	static List<int> ParagraphEnds(string rtf) {
		var ends = new List<int>();
		Scan(rtf, (word, at, _, _) => { if (word is "par" or "cell") ends.Add(at); });
		return ends;
	}

	// Each body paragraph's \s number (0 if none), in order.
	static List<int> ParagraphStyles(string rtf) {
		var styles = new List<int>();
		int current = 0;
		Scan(rtf, (word, at, value, _) => {
			if (word == "pard") current = 0;
			else if (word == "s" && value is int n) current = n;
			else if (word is "par" or "cell") styles.Add(current);
		});
		return styles;
	}

	// Walks the body's control words (word, position of its backslash, its number, the position after it), outside the
	// destinations whose text isn't the document's, and outside {\* ...} groups.
	static void Scan(string rtf, Action<string, int, int?, int> seen) {
		var skip = new Stack<bool>();
		bool skipping = false;
		int i = 0;
		while (i < rtf.Length) {
			char c = rtf[i];
			if (c == '{') {
				skip.Push(skipping);
				// the group's first control word decides
				int j = i + 1;
				bool starred = j + 1 < rtf.Length && rtf[j] == '\\' && rtf[j + 1] == '*';
				if (starred) j += 2;
				var word = Regex.Match(rtf.Substring(j, Math.Min(40, rtf.Length - j)), @"^\\([a-z]+)");
				if (starred || (word.Success && NotText.Contains(word.Groups[1].Value))) skipping = true;
				i++;
				continue;
			}
			if (c == '}') { skipping = skip.Count > 0 && skip.Pop(); i++; continue; }
			if (c != '\\') { i++; continue; }
			if (i + 1 >= rtf.Length) break;
			char d = rtf[i + 1];
			if (d == '\'') { i += 4; continue; }
			if (!char.IsLetter(d)) { i += 2; continue; }
			int k = i + 1;
			while (k < rtf.Length && char.IsLetter(rtf[k])) k++;
			string name = rtf[(i + 1)..k];
			int numStart = k;
			if (k < rtf.Length && rtf[k] == '-') k++;
			while (k < rtf.Length && char.IsDigit(rtf[k])) k++;
			int? value = k > numStart && int.TryParse(rtf[numStart..k], out int v) ? v : null;
			if (k < rtf.Length && rtf[k] == ' ') k++;
			if (!skipping && skip.Count >= 1) seen(name, i, value, k);
			if (name is "bin" && value is int bytes) k += bytes;
			i = k;
		}
	}

	static int GroupEnd(string rtf, int start) {
		int depth = 0;
		for (int i = start; i < rtf.Length; i++) {
			if (rtf[i] == '\\') { i++; continue; }
			if (rtf[i] == '{') depth++;
			else if (rtf[i] == '}' && --depth == 0) return i + 1;
		}
		return rtf.Length;
	}

	static string RemoveGroup(string rtf, string start) {
		int at = rtf.IndexOf(start, StringComparison.Ordinal);
		return at < 0 ? rtf : rtf.Remove(at, GroupEnd(rtf, at) - at);
	}

	// Where the style sheet goes: after the font table and the colour table.
	static int AfterTables(string rtf) {
		int at = 0;
		foreach (var table in new[] { "{\\fonttbl", "{\\colortbl" }) {
			int found = rtf.IndexOf(table, StringComparison.Ordinal);
			if (found >= 0) at = Math.Max(at, GroupEnd(rtf, found));
		}
		if (at == 0) {
			var header = Regex.Match(rtf, @"^\{\\rtf1[^{\\]*(\\[a-z]+-?\d*\s?)*");
			at = header.Success ? header.Length : Math.Min(rtf.Length, 6);
		}
		return at;
	}
}

// Styles through the clipboard: RichEdit's own copy and paste keep only the headings' styles, so Copy and Cut also put
// the copied paragraphs' styles (and their definitions) on the clipboard, in a format of RTE Writer's own, and Paste gives
// the pasted paragraphs their styles again (a style the document hasn't got comes with its definition, from another
// window's document too).
public static class StyleClipboard
{
	public const string Format = "RTE Writer paragraph styles";

	/// The clipboard (tests use one of their own, so as to leave the user's as it is).
	public static Func<IDataObject?> Read = Clipboard.GetDataObject;
	public static Action<IDataObject> Write = data => Clipboard.SetDataObject(data, true);

	public sealed class Data
	{
		public List<string> Names { get; set; } = new();
		public List<ParagraphStyle> Styles { get; set; } = new();
	}

	// The paragraphs whose marks (paragraph ends, cell ends; not the marks of table rows) are from start to end: their starts.
	static List<int> Marked(ITextDocument2 doc, int start, int end) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var starts = new List<int>();
		int paragraph = 0;
		for (int i = 0; i < story.Length && i < end; i++) {
			if (story[i] is not ('\r' or '\u0007')) continue;
			string piece = story[paragraph..i];
			bool row = story[i] == '\r' && piece.Length > 0 && piece.Trim('\uFFF9', '\uFFFB') == "";
			if (i >= start && !row) starts.Add(paragraph);
			paragraph = i + 1;
		}
		return starts;
	}

	/// copy: RichEdit's own Copy or Cut; the selection's paragraph styles go on the clipboard with what it put there.
	public static void Copy(RichEditBox editor, StyleSheet sheet, Action copy) {
		var doc = editor.TextDocument;
		var selection = doc.Selection2;
		var data = new Data();
		foreach (int start in Marked(doc, selection.Start, selection.End)) {
			var style = StyleEngine.StyleAt(doc, sheet, start);
			data.Names.Add(style.Name);
			if (!data.Styles.Any(s => s.Name == style.Name)) data.Styles.Add(style.Copy());
		}
		copy();
		if (data.Names.Count == 0) return;
		try {
			var put = Read();
			if (put == null) return;
			var both = new DataObject();
			foreach (var format in new[] { DataFormats.Rtf, DataFormats.UnicodeText, DataFormats.Text })
				if (put.GetDataPresent(format, false) && put.GetData(format, false) is { } value) both.SetData(format, false, value);
			both.SetData(Format, false, JsonSerializer.Serialize(data));
			Write(both);
		} catch (System.Runtime.InteropServices.ExternalException) { }
	}

	/// paste: RichEdit's own Paste; then the pasted paragraphs get their styles back (if the clipboard has RTE Writer's).
	/// Returns whether the style sheet got styles it hadn't.
	public static bool Paste(RichEditBox editor, StyleSheet sheet, Action paste) {
		Data? data = null;
		try {
			if (Read()?.GetData(Format) is string json) data = JsonSerializer.Deserialize<Data>(json);
		} catch (Exception ex) when (ex is System.Runtime.InteropServices.ExternalException or JsonException) { }
		var doc = editor.TextDocument;
		int start = doc.Selection2.Start, before = doc.Range2(0, 0).StoryLength - (doc.Selection2.End - doc.Selection2.Start);
		paste();
		if (data == null || data.Names.Count == 0) return false;
		int end = start + (doc.Range2(0, 0).StoryLength - before);
		var marked = Marked(doc, start, end);
		if (marked.Count != data.Names.Count) return false;
		bool added = false;
		for (int i = 0; i < marked.Count; i++) {
			string name = data.Names[i];
			if (sheet.Find(name) == null) {
				sheet.Styles.Add(data.Styles.FirstOrDefault(s => s.Name == name)?.Copy() ?? sheet.Ensure(name));
				added = true;
			}
			StyleEngine.Tag(doc.Range2(marked[i], marked[i]).Para, StyleEngine.IdOf(name));
		}
		return added;
	}
}
// Styles in HTML: headings as h1-h6 (from their style, not their font), other styles than Normal as data-rte-style; read
// back by markers put in before the HTML is loaded (StyleHtml.MarkHtml) and turned into the paragraphs' styles once it is
// (StyleHtml.FromMarkers).
public static class StyleHtml
{
	static readonly Regex BlockStart = new(@"<(p|h[1-6]|li|pre)\b([^>]*)>", RegexOptions.IgnoreCase);
	const string MarkerStart = "@@RTE-STYLE:", MarkerEnd = "@@";
	static readonly Regex Marker = new(Regex.Escape(MarkerStart) + @"([^@]+)" + Regex.Escape(MarkerEnd));

	/// RichEditorNET's HTML (one block per paragraph) with each paragraph's tag from its style: h1-h6 for headings, p
	/// for the others (with data-rte-style for styles other than Normal); lists and preformatted text are left alone.
	public static string Export(RichEditBox editor, StyleSheet sheet, string html) {
		var doc = editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var starts = new List<int>();
		int start = 0;
		for (int i = 0; i < story.Length; i++)
			if (story[i] is '\r' or '\u0007' or '\f') { starts.Add(start); start = i + 1; }
		var blocks = BlockStart.Matches(html).ToList();
		if (blocks.Count != starts.Count) return html;
		var sb = new StringBuilder(html.Length + 256);
		int position = 0;
		for (int k = 0; k < blocks.Count; k++) {
			var block = blocks[k];
			string tag = block.Groups[1].Value.ToLowerInvariant();
			if (tag is "li" or "pre" || block.Index < position) continue;
			int at = starts[k];
			if (at < story.Length && story[at] is '￹' or '￻') continue;
			var style = StyleEngine.StyleAt(doc, sheet, at);
			int level = StyleEngine.LevelAt(doc, sheet, at);
			string want = level > 0 ? "h" + Math.Min(6, level) : "p";
			string attributes = Regex.Replace(block.Groups[2].Value, @"\s+data-rte-style=""[^""]*""", "");
			bool normal = string.Equals(style.Name, "Normal", StringComparison.OrdinalIgnoreCase);
			if (level == 0 && !normal) attributes += $" data-rte-style=\"{System.Net.WebUtility.HtmlEncode(style.Name)}\"";
			// (a style's look in the paragraph's own CSS: browsers, and the built-in Word converter, show it as the editor does)
			if (!normal) {
				var css = Regex.Match(attributes, @"\sstyle=""([^""]*)""", RegexOptions.IgnoreCase);
				string look = Css(style);
				attributes = css.Success
					? attributes.Remove(css.Index, css.Length).Insert(css.Index, $" style=\"{look}; {css.Groups[1].Value}\"")
					: attributes + $" style=\"{look}\"";
			}
			// (the library takes a paragraph starting in bold at 24, 18, 14, 12, 10 or 8 points for a heading, and writes
			// its text without bold and without that size: a paragraph that isn't one has its text written again)
			string? inner = level == 0 && tag.Length == 2 && tag[0] == 'h' ? Inline(editor, at, story) : null;
			if (want == tag && attributes == block.Groups[2].Value && inner == null) continue;
			// the block's end tag
			var close = new Regex("</" + tag + @"\s*>", RegexOptions.IgnoreCase).Match(html, block.Index + block.Length);
			if (!close.Success) continue;
			sb.Append(html, position, block.Index - position);
			sb.Append('<').Append(want).Append(attributes).Append('>');
			if (inner != null) sb.Append(inner);
			else sb.Append(html, block.Index + block.Length, close.Index - block.Index - block.Length);
			sb.Append("</").Append(want).Append('>');
			position = close.Index + close.Length;
		}
		if (position == 0) return html;
		sb.Append(html, position, html.Length - position);
		return sb.ToString();
	}

	// A paragraph's text as the library writes a paragraph's (its inline writer, which isn't public; null: not
	// available, or the paragraph has a picture, which it writes only with the editor's OLE interface).
	static System.Reflection.MethodInfo? inlineWriter;
	static bool inlineLooked;
	static string? Inline(RichEditBox editor, int start, string story) {
		int end = story.IndexOfAny(new[] { '\r', '\u0007', '\f' }, start);
		if (end < 0) end = story.Length;
		if (story.IndexOf('\uFFFC', start, end - start) >= 0) return null;
		if (!inlineLooked) {
			inlineLooked = true;
			inlineWriter = typeof(RichEditBox).Assembly.GetType("EllipticBit.RichEditorNET.Formatting.HtmlFormatter")
				?.GetMethod("EmitHtmlInlineContent", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
			if (inlineWriter?.GetParameters().Length != 10) inlineWriter = null;
		}
		if (inlineWriter == null) return null;
		var sb = new StringBuilder();
		try {
			inlineWriter.Invoke(null, new object?[] { editor.TextDocument, null, sb, start, end, !editor.EnableHtmlFontSizing, false,
				editor.Font.Name, editor.Font.SizeInPoints, editor.HtmlThumbnailImageFormat });
		} catch (Exception ex) when (ex is System.Reflection.TargetInvocationException or ArgumentException) {
			return null;
		}
		return sb.ToString();
	}

	/// A style's look as CSS (its font and alignment; the spacing and indents are written with the paragraph's own).
	public static string Css(ParagraphStyle style) {
		var parts = new List<string> {
			"font-family: '" + System.Net.WebUtility.HtmlEncode(style.Font).Replace("'", "") + "'",
			"font-size: " + style.Size.ToString("0.#", CultureInfo.InvariantCulture) + "pt",
			"font-weight: " + (style.Bold ? "bold" : "normal"),
			"font-style: " + (style.Italic ? "italic" : "normal"),
			"text-decoration: " + (style.Underline ? "underline" : "none")
		};
		if (style.Color is { Length: 7 }) parts.Add("color: " + style.Color);
		if (style.Align != "left") parts.Add("text-align: " + style.Align);
		return string.Join("; ", parts);
	}

	const string MetaName = "rte-writer-styles";

	/// The document's styles for RTE Writer, in the HTML file's head (as the header and footer are).
	public static string Meta(StyleSheet sheet) =>
		$"<meta name=\"{MetaName}\" content=\"{Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(sheet)))}\">";

	/// The styles an HTML file has for RTE Writer (null: none).
	public static StyleSheet? FromMeta(string html) {
		var m = Regex.Match(html, "<meta\\s+name=\"" + MetaName + "\"\\s+content=\"([A-Za-z0-9+/=]+)\"", RegexOptions.IgnoreCase);
		if (!m.Success) return null;
		try { return JsonSerializer.Deserialize<StyleSheet>(Encoding.UTF8.GetString(Convert.FromBase64String(m.Groups[1].Value))); }
		catch (Exception ex) when (ex is FormatException or JsonException) { return null; }
	}

	/// HTML to load, with a marker after each heading's and styled paragraph's start tag: its style.
	public static string MarkHtml(string html) => Regex.Replace(html, @"<(h[1-6]|p)\b([^>]*)>", m => {
		string tag = m.Groups[1].Value.ToLowerInvariant();
		string? name = tag[0] == 'h' ? "Heading " + tag[1]
			: Regex.Match(m.Groups[2].Value, @"\sdata-rte-style=""([^""]+)""", RegexOptions.IgnoreCase) is { Success: true } s ? System.Net.WebUtility.HtmlDecode(s.Groups[1].Value) : null;
		return name == null ? m.Value : m.Value + MarkerStart + System.Net.WebUtility.HtmlEncode(name).Replace("@", "&#64;") + MarkerEnd;
	}, RegexOptions.IgnoreCase);

	/// The markers MarkHtml left as the paragraphs' styles (removed); returns the styles' names that were used.
	public static HashSet<string> FromMarkers(RichEditBox editor) {
		var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var doc = editor.TextDocument;
		while (true) {
			string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
			var m = Marker.Matches(story).LastOrDefault();
			if (m == null) break;
			doc.Range2(m.Index, m.Index + m.Length).Text = "";
			string name = System.Net.WebUtility.HtmlDecode(m.Groups[1].Value);
			StyleEngine.Tag(doc.Range2(m.Index, m.Index).Para, StyleEngine.IdOf(name));
			used.Add(name);
		}
		return used;
	}
}

// Modify Style and New Style: a style's settings, with a preview.
sealed class StyleDialog : Form
{
	readonly TextBox name = new() { Width = 220 };
	readonly ComboBox level = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
	readonly ComboBox next = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
	readonly ComboBox font = new() { DropDownStyle = ComboBoxStyle.DropDown, Width = 220, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems };
	readonly NumericUpDown size = new() { Minimum = 1, Maximum = 400, DecimalPlaces = 1, Increment = 0.5m, Width = 70 };
	readonly CheckBox bold = new() { Text = "Bold", AutoSize = true }, italic = new() { Text = "Italic", AutoSize = true }, underline = new() { Text = "Underline", AutoSize = true };
	readonly CheckBox automatic = new() { Text = "Automatic colour", AutoSize = true };
	readonly Button colorButton = new() { Text = "Colour...", AutoSize = true };
	readonly ComboBox align = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 110 };
	readonly NumericUpDown left = Points(), right = Points(), first = Points(-1584), before = Points(), after = Points();
	readonly ComboBox line = new() { DropDownStyle = ComboBoxStyle.DropDown, Width = 140 };
	readonly CheckBox keep = new() { Text = "Keep with the next paragraph", AutoSize = true };
	readonly Panel preview = new() { Height = 120, Dock = DockStyle.Fill, BackColor = SystemColors.Window, BorderStyle = BorderStyle.FixedSingle };
	readonly Label error = new() { AutoSize = true, ForeColor = Color.Firebrick };
	Color color = Color.Black;
	readonly ParagraphStyle original;
	readonly StyleSheet sheet;
	readonly bool isNew;

	static NumericUpDown Points(int minimum = 0) => new() { Minimum = minimum, Maximum = 1584, DecimalPlaces = 1, Increment = 1, Width = 70 };
	static readonly string[] Aligns = { "left", "center", "right", "justify" };

	/// style: the style to change (or the new one's first settings); isNew: a new style (its name to give).
	public StyleDialog(ParagraphStyle style, StyleSheet sheet, bool isNew) {
		original = style;
		this.sheet = sheet;
		this.isNew = isNew;
		Text = isNew ? "New Style" : "Modify Style: " + style.Name;
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MaximizeBox = MinimizeBox = false;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoScaleMode = AutoScaleMode.Dpi;
		Font = SystemFonts.MessageBoxFont ?? Font;

		bool builtIn = !isNew && ParagraphStyle.BuiltIn(style.Name);
		name.Text = style.Name;
		name.Enabled = !builtIn;
		level.Items.Add("Body text");
		for (int n = 1; n <= 6; n++) level.Items.Add("Heading level " + n);
		level.SelectedIndex = Math.Clamp(style.Level, 0, 6);
		level.Enabled = !builtIn;
		next.Items.Add("(the same style)");
		foreach (var s in sheet.Styles) next.Items.Add(s.Name);
		next.SelectedIndex = style.Next == "" ? 0 : Math.Max(0, next.Items.IndexOf(sheet.Find(style.Next)?.Name ?? "(the same style)"));
		using (var fonts = new InstalledFontCollection()) foreach (var family in fonts.Families) font.Items.Add(family.Name);
		font.Text = style.Font;
		size.Value = (decimal)Math.Clamp(style.Size, 1, 400);
		bold.Checked = style.Bold; italic.Checked = style.Italic; underline.Checked = style.Underline;
		automatic.Checked = style.Color == "";
		if (style.Color != "") try { color = ColorTranslator.FromHtml(style.Color); } catch (Exception) { }
		colorButton.Enabled = !automatic.Checked;
		colorButton.Click += (s, e) => { using var dlg = new ColorDialog { Color = color, FullOpen = true }; if (dlg.ShowDialog(this) == DialogResult.OK) { color = dlg.Color; preview.Invalidate(); } };
		automatic.CheckedChanged += (s, e) => { colorButton.Enabled = !automatic.Checked; preview.Invalidate(); };
		align.Items.AddRange(new object[] { "Left", "Centered", "Right", "Justified" });
		align.SelectedIndex = Math.Max(0, Array.IndexOf(Aligns, style.Align));
		left.Value = (decimal)Math.Clamp(style.LeftIndent, 0, 1584);
		right.Value = (decimal)Math.Clamp(style.RightIndent, 0, 1584);
		first.Value = (decimal)Math.Clamp(style.FirstLine, -1584, 1584);
		before.Value = (decimal)Math.Clamp(style.Before, 0, 1584);
		after.Value = (decimal)Math.Clamp(style.After, 0, 1584);
		line.Items.AddRange(new object[] { "1", "1.15", "1.5", "2", "exactly 12pt", "exactly 24pt", "at least 12pt" });
		line.Text = style.LineSpacing;
		keep.Checked = style.KeepWithNext;
		foreach (Control c in new Control[] { font, size, bold, italic, underline, align, left, right, first, before, after, line })
			switch (c) {
				case ComboBox box: box.TextChanged += (s, e) => preview.Invalidate(); box.SelectedIndexChanged += (s, e) => preview.Invalidate(); break;
				case NumericUpDown number: number.ValueChanged += (s, e) => preview.Invalidate(); break;
				case CheckBox check: check.CheckedChanged += (s, e) => preview.Invalidate(); break;
			}
		preview.Paint += PaintPreview;

		var table = new TableLayoutPanel { ColumnCount = 4, AutoSize = true, Padding = new Padding(12), Dock = DockStyle.Fill };
		void Row(string label, Control a, string? label2 = null, Control? b = null) {
			int r = table.RowCount++;
			table.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) }, 0, r);
			table.Controls.Add(a, 1, r);
			if (label2 != null) table.Controls.Add(new Label { Text = label2, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 6, 8, 0) }, 2, r);
			if (b != null) table.Controls.Add(b, 3, r);
		}
		FlowLayoutPanel Flow(params Control[] controls) { var f = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) }; f.Controls.AddRange(controls); return f; }
		Row("Name:", name, "Heading:", level);
		Row("After Enter:", next);
		Row("Font:", font, "Size (pt):", size);
		Row("", Flow(bold, italic, underline), "", Flow(automatic, colorButton));
		Row("Alignment:", align, "Line spacing:", line);
		Row("Indent left (pt):", left, "Right (pt):", right);
		Row("First line (pt):", first, "", new Label { Text = "(negative: hanging)", AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(0, 6, 0, 0) });
		Row("Space before (pt):", before, "After (pt):", after);
		Row("", keep);
		int pr = table.RowCount++;
		table.Controls.Add(preview, 0, pr);
		table.SetColumnSpan(preview, 4);
		int er = table.RowCount++;
		table.Controls.Add(error, 0, er);
		table.SetColumnSpan(error, 4);
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var ok = new Button { Text = "OK", AutoSize = true };
		ok.Click += (s, e) => { if (Validate(out string message)) { DialogResult = DialogResult.OK; Close(); } else error.Text = message; };
		buttons.Controls.Add(cancel);
		buttons.Controls.Add(ok);
		int br = table.RowCount++;
		table.Controls.Add(buttons, 0, br);
		table.SetColumnSpan(buttons, 4);
		Controls.Add(table);
		AcceptButton = ok;
		CancelButton = cancel;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
	}

	bool Validate(out string message) {
		message = "";
		string n = name.Text.Trim();
		if (n == "" || n.Contains('@') || n.Contains(';') || n.Contains('{') || n.Contains('}')) { message = "Give the style a name (without @ ; { })."; return false; }
		if ((isNew || !string.Equals(n, original.Name, StringComparison.OrdinalIgnoreCase)) && sheet.Find(n) != null) { message = $"There is a style \"{n}\" already."; return false; }
		if (isNew && ParagraphStyle.BuiltIn(n)) { message = $"\"{n}\" is one of the styles there always are."; return false; }
		try { global::LineSpacing.Parse(line.Text); } catch (ArgumentException ex) { message = ex.Message; return false; }
		if (font.Text.Trim() == "") { message = "Choose a font."; return false; }
		return true;
	}

	/// The style as set in the dialog.
	public ParagraphStyle Result => new() {
		Name = ParagraphStyle.BuiltIn(original.Name) && !isNew ? original.Name : name.Text.Trim(),
		Level = ParagraphStyle.BuiltIn(original.Name) && !isNew ? original.Level : level.SelectedIndex,
		Next = next.SelectedIndex <= 0 ? "" : next.Text,
		Font = font.Text.Trim(), Size = (float)size.Value, Bold = bold.Checked, Italic = italic.Checked, Underline = underline.Checked,
		Color = automatic.Checked ? "" : $"#{color.R:X2}{color.G:X2}{color.B:X2}",
		Align = Aligns[Math.Max(0, align.SelectedIndex)], LeftIndent = (float)left.Value, RightIndent = (float)right.Value, FirstLine = (float)first.Value,
		Before = (float)before.Value, After = (float)after.Value, LineSpacing = line.Text.Trim(), KeepWithNext = keep.Checked
	};

	// A sample paragraph in the style, with grey text before and after it (the space and indents to scale).
	void PaintPreview(object? sender, PaintEventArgs e) {
		var g = e.Graphics;
		g.Clear(SystemColors.Window);
		float scale = DeviceDpi / 72f * 0.5f;
		var area = preview.ClientRectangle;
		area.Inflate(-8, -6);
		using var grey = new SolidBrush(Color.FromArgb(0xB0, 0xB0, 0xB0));
		using var context = new Font("Segoe UI", 7f);
		g.DrawString("The paragraph before, in grey.", context, grey, area.Left, area.Top);
		float y = area.Top + context.Height + (float)before.Value * scale;
		var styleFont = FontStyle.Regular | (bold.Checked ? FontStyle.Bold : 0) | (italic.Checked ? FontStyle.Italic : 0) | (underline.Checked ? FontStyle.Underline : 0);
		Font sample;
		try { sample = new Font(font.Text.Trim() == "" ? "Segoe UI" : font.Text.Trim(), Math.Max(4f, (float)size.Value * 0.75f), styleFont); } catch (ArgumentException) { sample = new Font("Segoe UI", 9f, styleFont); }
		using (sample)
		using (var ink = new SolidBrush(automatic.Checked ? SystemColors.WindowText : color)) {
			float x = area.Left + (float)left.Value * scale, width = Math.Max(40, area.Width - (float)(left.Value + right.Value) * scale);
			var format = new StringFormat { Alignment = align.SelectedIndex switch { 1 => StringAlignment.Center, 2 => StringAlignment.Far, _ => StringAlignment.Near } };
			string text = "A sample of the style: the quick brown fox jumps over the lazy dog, and goes on to a second line.";
			float firstLine = (float)first.Value * scale;
			g.DrawString(text, sample, ink, new RectangleF(x + Math.Max(0, firstLine), y, width - Math.Max(0, firstLine), area.Bottom - y), format);
			y += sample.Height * 2 + (float)after.Value * scale;
		}
		if (y < area.Bottom) g.DrawString("The paragraph after, in grey.", context, grey, area.Left, y);
	}
}
