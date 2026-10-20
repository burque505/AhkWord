// Legal numbering (1., 1.1, 1.1.1) for RichEditBox, used by AhkWord (AhkWord.cs, through #load).
//
// RichEdit numbers each list level on its own (1. / a. / i.) and cannot show compound numbers; given an RTF list table
// with them, it drops the list. So a legal list is made of list paragraphs whose own number is hidden (TOM list type
// flag 0x40000, "no number": the paragraph keeps its list level and indent) with the number written in front of the
// text, "1.2.1" and a tab, and kept up to date by Renumber.
// The numbers are text, so commands that change a legal list renumber it in the same undo step (BeginEditCollection),
// and Ctrl+Z undoes both together.
//
// Numbering goes on through the whole document, over the paragraphs between legal items (body text between clauses),
// until it is restarted. Each item's ListStart says which series it belongs to: the top level restarts where a
// top-level item's ListStart differs from the one before it, at that number (ListStart mod 1000: a restart at the same
// number as the series before gets 1000 added, to tell the two apart). A deeper level starts at 1 under each item of
// the level above.
// Loaders (HTML, Markdown, typed numbers) mark items as Normalize expects: 0 for "goes on", Loaded + n for "a list
// starting at n" (it goes on if n is the next number, and restarts otherwise).
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static class LegalLists
{
	public const int NoNumber = 0x40000;    // TOM ListType flag: a list paragraph without its number
	public const int Loaded = 30000;        // ListStart from a loader: Loaded + the list's first number

	/// How the numbers look: "1." on the top level and "1.1" below (Top), "1." and "1.1." (All), or "1" and "1.1" (None).
	public enum Periods { Top, All, None }
	public static Periods Format = Periods.Top;

	/// A number already written in front of an item, then a tab: "1.2.1", or a numbered style's "3.", "(b)", "iv)" ...
	public static readonly Regex Prefix = new(@"^(?:\d+(?:\.\d+)*\.?|\((?:\d{1,4}|[a-z]{1,3}|[A-Z]{1,3}|[ivxlcdm]{1,7}|[IVXLCDM]{1,7})\)|(?:\d{1,4}|[a-z]{1,3}|[A-Z]{1,3}|[ivxlcdm]{1,7}|[IVXLCDM]{1,7})[.)])\t");

	/// Numbered styles (rte_styles.cs): the numbering of a paragraph style, by its TOM style number: its number's format
	/// (NumberFormats) and level. null: none, and the item is a legal item ("1.2.1", indents of its own). The editor sets
	/// it from its style sheet.
	public static Func<int, (string Format, int Level)?> StyleNumbering = _ => null;

	/// A numbered style's format at a level (its own level's, or a sub-level's below it: ParagraphStyle.FormatAt), by
	/// its TOM style number and the level (0-based); null: a legal item's.
	public static Func<int, int, string?> StyleFormat = (_, _) => null;

	/// A number typed in front of a paragraph ("1.2.1 " or "1.2.1<tab>"), which Make takes as the item's level.
	static readonly Regex Typed = new(@"^(\d+(?:\.\d+)*)\.?(?:\t| +)");

	// (a range over paragraphs with different list types has tomUndefined, which is negative)
	public static bool IsLegal(int listType) => listType > 0 && (listType & NoNumber) != 0 && (listType & 0xFFFF) != tomConstants.tomListNone;
	public static bool IsLegal(ITextPara para) => IsLegal(para.ListType);

	/// The first number of a series from its ListStart.
	public static int StartOf(int listStart) => listStart <= 0 ? 1 : (listStart - 1) % 1000 + 1;

	/// An item's number: a numbered style's format (NumberFormats: "1.", "(a)" ... of its own level's count), or the
	/// legal number of the counts down to its level ("1.2.1"; format null or NumberFormats.Legal).
	public static string Number(IReadOnlyList<int> counters, int level, string? format = null) {
		if (!string.IsNullOrEmpty(format) && format != NumberFormats.Legal) return NumberFormats.Render(format, counters[level]);
		string number = string.Join(".", counters.Take(level + 1));
		return Format switch { Periods.All => number + ".", Periods.Top => level == 0 ? number + "." : number, _ => number };
	}

	public sealed record Paragraph(int Start, int End, int ListType, int Level, int ListStart, string Text, int Style = 0)
	{
		public bool Legal => IsLegal(ListType);
		public int PrefixLength => Legal ? Prefix.Match(Text).Length : 0;
		/// Numbered by its style (whose indents it keeps), not a legal item of its own.
		public bool Styled => StyleNumbering(Style) != null;
	}

	public static List<Paragraph> Paragraphs(ITextDocument2 doc) {
		var list = new List<Paragraph>();
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var range = doc.Range2(0, 0);
		range.Expand(tomConstants.tomParagraph);
		while (true) {
			var para = range.Para;
			list.Add(new Paragraph(range.Start, range.End, para.ListType, Math.Max(0, para.ListLevelIndex), para.ListStart, (range.Text ?? "").TrimEnd('\r'), para.Style));
			if (range.End >= storyEnd || range.End <= range.Start) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
		return list;
	}

	/// Whether the document has legal items (reads only the paragraphs' list types, so it is quick).
	public static bool AnyLegal(ITextDocument2 doc) {
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var range = doc.Range2(0, 0);
		range.Expand(tomConstants.tomParagraph);
		while (true) {
			if (IsLegal(range.Para.ListType)) return true;
			if (range.End >= storyEnd || range.End <= range.Start) return false;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
	}

	/// The number each legal paragraph should have, and the series (ListStart) it belongs to, by paragraph start.
	public static Dictionary<int, (string Number, int Series, int[] Counters)> Numbers(List<Paragraph> paragraphs) {
		var numbers = new Dictionary<int, (string, int, int[])>();
		var counters = new int[10];
		int? series = null;
		foreach (var p in paragraphs) {
			if (!p.Legal) continue;
			int level = Math.Min(p.Level, 9);
			if (level == 0 && p.ListStart != series) {
				counters = new int[10];
				counters[0] = StartOf(p.ListStart) - 1;
				series = p.ListStart;
			}
			series ??= p.ListStart;
			counters[level]++;
			for (int k = level + 1; k < counters.Length; k++) counters[k] = 0;
			for (int k = 0; k < level; k++) if (counters[k] == 0) counters[k] = 1;    // a skipped level counts as 1
			numbers[p.Start] = (Number(counters, level, StyleFormat(p.Style, level)), series.Value, (int[])counters.Clone());
		}
		return numbers;
	}

	/// A numbered item as the editor numbers it: its level, its numbered style's format (null: a legal item's), its
	/// number's text and its series (LegalLists.Numbers).
	public sealed record Item(int Level, string? Format, string Number, int Series);

	/// The document's numbered items, in order.
	public static List<Item> Items(ITextDocument2 doc) {
		var paragraphs = Paragraphs(doc);
		var numbers = Numbers(paragraphs);
		return paragraphs.Where(p => numbers.ContainsKey(p.Start))
			.Select(p => new Item(Math.Min(p.Level, 9), StyleFormat(p.Style, Math.Min(p.Level, 9)), numbers[p.Start].Number, numbers[p.Start].Series)).ToList();
	}

	/// Writes the right number in front of every legal item, and gives nested items their series' ListStart (Shift+Tab
	/// can make them top-level items). Keeps the selection where it was in the text. Returns true if anything changed.
	public static bool Renumber(RichEditBox editor) => Renumber(editor.TextDocument);

	public static bool Renumber(ITextDocument2 doc) {
		var paragraphs = Paragraphs(doc);
		var numbers = Numbers(paragraphs);
		var edits = new List<(int Start, int Length, string Text)>();
		bool changed = false;
		foreach (var p in paragraphs) {
			if (!numbers.TryGetValue(p.Start, out var number)) continue;
			if (p.Level > 0 && p.ListStart != number.Series) {
				doc.Range2(p.Start, p.Start).Para.ListStart = number.Series;
				changed = true;
			}
			string wanted = number.Number + "\t";
			var m = Prefix.Match(p.Text);
			if (m.Success && m.Value == wanted) continue;
			edits.Add((p.Start, m.Success ? m.Length : 0, wanted));
		}
		if (edits.Count > 0) {
			var selection = doc.Selection2;
			int selStart = selection.Start, selEnd = selection.End;
			// positions are worked out against the original text; the edits are applied from the end
			int newStart = ShiftAll(selStart, edits);
			int newEnd = selEnd == selStart ? newStart : ShiftAll(selEnd, edits);
			foreach (var e in Enumerable.Reverse(edits)) doc.Range2(e.Start, e.Start + e.Length).Text = e.Text;
			selection.SetRange(newStart, newEnd);
			changed = true;
		}
		return FitIndents(doc) | changed;
	}

	/// The hanging indent of each level (points): wide enough for that level's widest number in its font and size, so
	/// that the tab after it goes to the text (a number wider than the indent pushed its text to the next tab stop).
	/// Renumber measures it; Indents, and the exports' list definitions, use it.
	public static readonly float[] Hang = Enumerable.Range(0, 10).Select(level => 22f + 7 * level).ToArray();

	// Measures the documents' legal numbers and gives every legal item its level's indents. Returns true if any changed.
	static bool FitIndents(ITextDocument2 doc) {
		var items = Paragraphs(doc).Where(p => p.Legal && p.PrefixLength > 0 && !p.Styled).ToList();    // (a numbered style's items: its indents)
		if (items.Count == 0) return false;
		var needed = new float[10];
		var sizes = new Dictionary<int, float>();
		foreach (var p in items) {
			var font = doc.Range2(p.Start, p.Start + 1).Font;
			float size = font.Size > 0 ? font.Size : 11;
			sizes[p.Start] = size;
			int level = Math.Min(p.Level, 9);
			float width = TextWidth(p.Text[..p.PrefixLength].TrimEnd('\t'), font.Name ?? "Segoe UI", size) + size * 0.5f;    // and a gap
			needed[level] = Math.Max(needed[level], width);
		}
		for (int level = 0; level < 10; level++) Hang[level] = needed[level] > 0 ? MathF.Ceiling(Math.Max(needed[level], 18)) : 22f + 7 * level;
		bool changed = false;
		foreach (var p in items) {
			var para = doc.Range2(p.Start, p.Start).Para;
			var (first, left) = Indents(p.Level);
			if (Math.Abs(para.FirstLineIndent - first) < 0.5f && Math.Abs(para.LeftIndent - left) < 0.5f) continue;
			para.SetIndents(first, left, para.RightIndent);
			changed = true;
		}
		return changed;
	}

	static readonly Dictionary<(string, float, string), float> widths = new();
	static float? screenDpi;

	// The width of text in a font, in points.
	public static float TextWidth(string text, string fontName, float size) {
		if (widths.TryGetValue((fontName, size, text), out float known)) return known;
		if (screenDpi == null) {
			using var screen = System.Drawing.Graphics.FromHwnd(IntPtr.Zero);
			screenDpi = screen.DpiX;
		}
		float points;
		try {
			using var font = new System.Drawing.Font(fontName, size, System.Drawing.GraphicsUnit.Point);
			var pixels = System.Windows.Forms.TextRenderer.MeasureText(text, font, System.Drawing.Size.Empty, System.Windows.Forms.TextFormatFlags.NoPadding);
			points = pixels.Width * 72f / screenDpi.Value;
		} catch (ArgumentException) {
			points = text.Length * size * 0.55f;
		}
		return widths[(fontName, size, text)] = points;
	}

	// Where a position ends up after the edits (in document order, not overlapping): moved by the edits before it; one in
	// (or at) a number that is replaced goes after the new number.
	static int ShiftAll(int position, List<(int Start, int Length, string Text)> edits) {
		int result = position;
		foreach (var e in edits) {
			if (position >= e.Start + e.Length) result += e.Text.Length - e.Length;
			else if (position >= e.Start) result += e.Start + e.Text.Length - position;
		}
		return result;
	}

	/// Runs action and the renumbering after it as one undo step.
	public static void Grouped(RichEditBox editor, Action action) {
		var doc = editor.TextDocument;
		doc.BeginEditCollection();
		try {
			action();
			Renumber(editor);
		} finally {
			doc.EndEditCollection();
		}
	}

	/// Turns the marks loaders leave in ListStart (0: goes on; Loaded + n: a list starting at n) into series. A list
	/// starting at the next number goes on; one starting elsewhere restarts there. Values that are already series are
	/// kept, so it can run on a document that is part loaded, part not.
	public static void Normalize(ITextDocument2 doc) {
		int? series = null;
		int count = 0;    // the top level's last number
		foreach (var p in Paragraphs(doc)) {
			if (!p.Legal) continue;
			var para = doc.Range2(p.Start, p.Start).Para;
			int value = p.ListStart;
			if (p.Level > 0) {
				series ??= 1;
				if (value != series) para.ListStart = series.Value;
				continue;
			}
			if (series != null && (value <= 0 || value == series)) {    // goes on
				if (value != series) para.ListStart = series.Value;
				count++;
				continue;
			}
			if (series == null && value <= 0) {    // the document's first top-level item, marked "goes on": from 1
				series = 1;
				para.ListStart = 1;
				count = 1;
				continue;
			}
			int start = value >= Loaded ? value - Loaded : StartOf(value);
			if (value >= Loaded && series != null && start == count + 1) {    // a list that carries on
				para.ListStart = series.Value;
				count++;
				continue;
			}
			// a restart at start: a value that differs from the series before it
			int restart = start;
			while (restart == series) restart += 1000;
			if (restart != value) para.ListStart = restart;
			series = restart;
			count = start;
		}
	}

	// ---------------------------------------------------------------- making and removing legal lists

	/// Indents for a legal item at a level, in points: (first line, left). The hanging part holds the number.
	/// The number starts where an ordinary list's number at the same level does (IndentsFor).
	public static (float First, float Left) Indents(int level) => IndentsFor(level, Hang[Math.Clamp(level, 0, 9)]);

	/// The document's list levels' own indents (Paragraph > List levels: Number at, Text at, points from the margin;
	/// null: automatic), set by the editor; null: none set.
	public static Func<int, (float? NumberAt, float? TextAt)>? LevelIndents;

	/// A list item's indents at a level, in points: (first line, left). The number at the level's Number at (automatic:
	/// 18 points a level, from 18), the text at its Text at (automatic: after the number, hang: the level's widest
	/// number and a gap). Lists, legal numbering and the Word and OpenDocument lists all take them from here.
	public static (float First, float Left) IndentsFor(int level, float hang) {
		level = Math.Clamp(level, 0, 8);
		var (numberAt, textAt) = LevelIndents?.Invoke(level) ?? (null, null);
		float number = numberAt ?? 18 * (level + 1);
		float text = textAt ?? number + hang;
		return (number - text, text);
	}

	/// The paragraphs from start to end (positions) that the selection touches.
	public static List<Paragraph> Touched(ITextDocument2 doc, int start, int end) =>
		Paragraphs(doc).Where(p => p.End > start && p.Start <= Math.Max(start, end - 1) || p.Start == start).ToList();

	/// Makes the paragraphs from start to end legal items. List items keep their level (a list that started at n starts
	/// the legal numbering at n, unless that carries on from the legal items before it); other paragraphs become
	/// top-level items, or take the level of a number typed in front of them ("2.1 Text" becomes a level-2 item).
	public static void Make(RichEditBox editor, int start, int end) => Grouped(editor, () => {
		var doc = editor.TextDocument;
		var touched = Touched(doc, start, end);
		bool first = true;
		foreach (var p in touched) {
			if (p.Legal) { first = false; continue; }
			var para = doc.Range2(p.Start, p.Start).Para;
			bool list = (p.ListType & 0xFFFF) != tomConstants.tomListNone;
			int level = list ? p.Level : 0;
			int mark = first && list && level == 0 ? Loaded + Math.Max(1, p.ListStart) : 0;
			SetLegal(para, level, mark);
			first = false;
		}
		// typed numbers go (from the end, as they change positions), giving their levels
		foreach (var p in Enumerable.Reverse(touched)) {
			if (p.Legal || (p.ListType & 0xFFFF) != tomConstants.tomListNone || Typed.Match(p.Text) is not { Success: true } typed) continue;
			int level = Math.Min(8, typed.Groups[1].Value.Count(c => c == '.'));
			doc.Range2(p.Start, p.Start + typed.Length).Text = "";
			var para = doc.Range2(p.Start, p.Start).Para;
			para.ListLevelIndex = level;
			var (firstIndent, left) = Indents(level);
			para.SetIndents(firstIndent, left, para.RightIndent);
		}
		Normalize(doc);
	});

	/// Makes a paragraph a legal item at a level (0-based), with a ListStart (see Normalize).
	public static void SetLegal(ITextPara para, int level, int listStart = 0) {
		para.ListType = tomConstants.tomListNumberAsArabic | NoNumber;
		para.ListLevelIndex = level;
		para.ListStart = Math.Max(0, listStart);
		var (first, left) = Indents(level);
		para.SetIndents(first, left, para.RightIndent);
	}

	/// Makes a paragraph an item of its numbered style at a level (0-based): numbered as a legal item is, its indents
	/// left as its style has them.
	public static void SetStyled(ITextPara para, int level, int listStart = 0) {
		// (RichEdit gives a paragraph that becomes a list item a list's indents of its own: the paragraph's are put back)
		float first = para.FirstLineIndent, left = para.LeftIndent, right = para.RightIndent;
		para.ListType = tomConstants.tomListNumberAsArabic | NoNumber;
		para.ListLevelIndex = Math.Clamp(level, 0, 8);
		para.ListStart = Math.Max(0, listStart);
		if (Math.Abs(para.FirstLineIndent - first) > 0.01f || Math.Abs(para.LeftIndent - left) > 0.01f || Math.Abs(para.RightIndent - right) > 0.01f)
			para.SetIndents(first, left, right);
		NoNumberSpace(para);
	}

	/// A numbered style's item: no room for a number of RichEdit's own (it has none: its number is text). RichEdit
	/// draws a list item's first line that far (ListTab, 18 points) to the right of its first-line indent, so a
	/// hanging "-21" showed at "-3", and Numbered Paragraph's 36 at 54; legal items' indents count on it, so keep it.
	public static void NoNumberSpace(ITextPara para) {
		if (Math.Abs(para.ListTab) > 0.01f) para.ListTab = 0;
	}

	/// Ends legal numbering for the paragraphs from start to end: their numbers go and they become normal paragraphs.
	public static void Remove(RichEditBox editor, int start, int end) => Grouped(editor, () => {
		var doc = editor.TextDocument;
		foreach (var p in Enumerable.Reverse(Touched(doc, start, end)).Where(p => p.Legal)) {
			if (p.PrefixLength > 0) doc.Range2(p.Start, p.Start + p.PrefixLength).Text = "";
			var para = doc.Range2(p.Start, p.Start).Para;
			para.ListType = tomConstants.tomListNone;
			para.SetIndents(0, 0, para.RightIndent);
		}
	});

	/// After a legal item's level changed: the indent of its new level.
	public static void Reindent(ITextDocument2 doc, int start, int end) {
		foreach (var p in Touched(doc, start, end).Where(p => p.Legal && !p.Styled)) {
			var para = doc.Range2(p.Start, p.Start).Para;
			var (first, left) = Indents(p.Level);
			para.SetIndents(first, left, para.RightIndent);
		}
	}

	// ---------------------------------------------------------------- restart and continue

	static int IndexAt(List<Paragraph> paragraphs, int position) {
		int index = paragraphs.FindIndex(p => position >= p.Start && position < Math.Max(p.End, p.Start + 1));
		if (index < 0 || !paragraphs[index].Legal) throw new InvalidOperationException("Put the caret in a legal numbered item first.");
		return index;
	}

	/// Restarts the top level at start from the top-level item with the caret, up to the next restart.
	public static void Restart(RichEditBox editor, int position, int start) {
		if (start < 0 || start > 999) throw new ArgumentException("Legal numbering restarts at 0 to 999.");
		var paragraphs = Paragraphs(editor.TextDocument);
		int index = IndexAt(paragraphs, position);
		if (paragraphs[index].Level != 0) throw new InvalidOperationException("Legal numbering restarts at the top level: put the caret in a top-level item (1., 2., ...).");
		int old = paragraphs[index].ListStart;
		int? before = paragraphs.Take(index).LastOrDefault(p => p.Legal && p.Level == 0)?.ListStart;
		int? after = paragraphs.Skip(index + 1).FirstOrDefault(p => p.Legal && p.Level == 0 && p.ListStart != old)?.ListStart;
		int value = start == 0 ? 1000 : start;    // (0 would mean "goes on" to a loader; 1000 shows 0 too)
		while (value == before || value == after) value += 1000;
		Grouped(editor, () => {
			for (int i = index; i < paragraphs.Count; i++) {
				var p = paragraphs[i];
				if (!p.Legal) continue;
				if (i > index && p.Level == 0 && p.ListStart != old) break;
				editor.TextDocument.Range2(p.Start, p.Start).Para.ListStart = value;
			}
		});
	}

	/// Undoes a restart: the numbering of the series with the caret carries on from the series before it.
	public static void Continue(RichEditBox editor, int position) {
		var paragraphs = Paragraphs(editor.TextDocument);
		int index = IndexAt(paragraphs, position);
		var numbers = Numbers(paragraphs);
		int series = numbers[paragraphs[index].Start].Series;
		int first = paragraphs.FindIndex(p => p.Legal && numbers[p.Start].Series == series);
		var previous = paragraphs.Take(first).LastOrDefault(p => p.Legal);
		if (previous == null) throw new InvalidOperationException("The numbering already starts here: there is no legal list before this one to continue.");
		int previousSeries = numbers[previous.Start].Series;
		Grouped(editor, () => {
			for (int i = first; i < paragraphs.Count; i++) {
				var p = paragraphs[i];
				if (!p.Legal) continue;
				if (numbers[p.Start].Series != series) break;
				editor.TextDocument.Range2(p.Start, p.Start).Para.ListStart = previousSeries;
			}
		});
	}

	// RichEdit's RTF doesn't keep a list paragraph's ListStart either where it has no number of its own, and with it goes
	// where a numbering starts again, nor its level (a numbered style's sub-levels): AhkWord's RTF keeps both, in
	// order, in {\*\rtewriterlists 1:0,1:1,1001:0} (start:level; files from before levels were kept have starts alone).
	static readonly Regex StartsGroupRx = new(@"\{\\\*\\rtewriterlists ([0-9,:\-]*)\}");

	/// The group with the items' ListStart values and levels ("": no items).
	public static string StartsGroup(ITextDocument2 doc) {
		var items = Paragraphs(doc).Where(p => p.Legal).Select(p => $"{p.ListStart}:{p.Level}").ToList();
		return items.Count == 0 ? "" : @"{\*\rtewriterlists " + string.Join(",", items) + "}";
	}

	/// The values an RTF has kept (null: none; a level of -1: not kept).
	public static List<(int Start, int Level)>? StartsFrom(string rtf) {
		var m = StartsGroupRx.Match(rtf);
		if (!m.Success) return null;
		try {
			return m.Groups[1].Value.Split(',', StringSplitOptions.RemoveEmptyEntries).Select(v => v.Split(':') is var parts && parts.Length > 1
				? (int.Parse(parts[0]), Math.Clamp(int.Parse(parts[1]), 0, 8)) : (int.Parse(parts[0]), -1)).ToList();
		} catch (FormatException) { return null; }
	}

	/// Gives the items their kept values (and levels) back, if there are as many items as values. Returns whether any
	/// changed.
	public static bool RestoreStarts(ITextDocument2 doc, List<(int Start, int Level)>? starts) {
		if (starts == null) return false;
		var items = Paragraphs(doc).Where(p => p.Legal).ToList();
		if (items.Count != starts.Count) return false;
		bool changed = false;
		for (int i = 0; i < items.Count; i++) {
			var para = doc.Range2(items[i].Start, items[i].Start).Para;
			if (items[i].ListStart != starts[i].Start) { para.ListStart = starts[i].Start; changed = true; }
			if (starts[i].Level >= 0 && items[i].Level != starts[i].Level) { para.ListLevelIndex = starts[i].Level; changed = true; }
		}
		return changed;
	}

	// "Numbering off": a paragraph of a numbered style whose numbering was turned off keeps that through other styles, as
	// Word's numbering off (numId 0) stays with its paragraph: under Normal and back to the numbered style it is still
	// unnumbered. The mark is a ListType of NoNumber alone, on a paragraph in no list: RichEdit draws nothing for it,
	// keeps it through typing and Enter (the new paragraph has it too, as Word's), and numbering the paragraph replaces
	// it. RichEdit's RTF doesn't keep it: AhkWord's keeps the marked paragraphs, by their index among Paragraphs, in
	// {\*\rtewriteroff count:i,j,...} (count: the paragraphs then, so that a file changed elsewhere isn't misread).
	public static bool IsNumberingOff(int listType) => listType == NoNumber;

	/// Marks a paragraph in no list as one whose numbering is off (its indents as they were).
	public static void SetNumberingOff(ITextPara para) {
		if ((para.ListType & 0xFFFF) != tomConstants.tomListNone || para.ListType == NoNumber) return;
		float first = para.FirstLineIndent, left = para.LeftIndent, right = para.RightIndent;
		para.ListType = NoNumber;
		if (Math.Abs(para.FirstLineIndent - first) > 0.01f || Math.Abs(para.LeftIndent - left) > 0.01f || Math.Abs(para.RightIndent - right) > 0.01f)
			para.SetIndents(first, left, right);
	}

	static readonly Regex OffGroupRx = new(@"\{\\\*\\rtewriteroff (\d+):([0-9,]*)\}");

	/// The group with the paragraphs whose numbering is off ("": none).
	public static string OffGroup(ITextDocument2 doc) {
		var paragraphs = Paragraphs(doc);
		var off = Enumerable.Range(0, paragraphs.Count).Where(i => IsNumberingOff(paragraphs[i].ListType)).ToList();
		return off.Count == 0 ? "" : $@"{{\*\rtewriteroff {paragraphs.Count}:{string.Join(",", off)}}}";
	}

	/// Marks the paragraphs an RTF kept as numbering off again (if it has as many paragraphs as it had).
	public static void RestoreOff(ITextDocument2 doc, string rtf) {
		var m = OffGroupRx.Match(rtf);
		if (!m.Success) return;
		var paragraphs = Paragraphs(doc);
		if (!int.TryParse(m.Groups[1].Value, out int count) || count != paragraphs.Count) return;
		foreach (var part in m.Groups[2].Value.Split(',', StringSplitOptions.RemoveEmptyEntries))
			if (int.TryParse(part, out int i) && i >= 0 && i < paragraphs.Count)
				SetNumberingOff(doc.Range2(paragraphs[i].Start, paragraphs[i].Start).Para);
	}

	/// The indents legal items had as loaded, by paragraph start, before LevelsFromNumbers gave them legal numbering's
	/// (styles are read after it: a numbered style's items get theirs back). Cleared when a document is loaded.
	public static readonly Dictionary<int, (float First, float Left)> LoadedIndents = new();

	/// After loading RTF: RichEdit's RTF doesn't keep the level of a list paragraph without a number, so legal items take
	/// it back from their numbers ("3.1.1" is the third level). Returns how many changed.
	public static int LevelsFromNumbers(ITextDocument2 doc) {
		int changed = 0;
		foreach (var p in Paragraphs(doc)) {
			if (!p.Legal || Prefix.Match(p.Text) is not { Success: true } m) continue;
			var styled = StyleNumbering(p.Style);    // (a numbered style's item: the level its style gives, its indents)
			int level = styled?.Level ?? Math.Min(8, m.Value.TrimEnd('\t', '.').Count(c => c == '.'));
			int type = tomConstants.tomListNumberAsArabic | NoNumber;    // (RichEdit reads them back as bullets without a number)
			if (level == p.Level && p.ListType == type) continue;
			var para = doc.Range2(p.Start, p.Start).Para;
			if (p.ListType != type) para.ListType = type;
			para.ListLevelIndex = level;
			if (styled == null) {
				// (indents as loaded kept, for an item whose style turns out to be numbered once styles are read: its
				// own; StyleEngine.NumberAfterLoad puts them back)
				LoadedIndents[p.Start] = (para.FirstLineIndent, para.LeftIndent);
				var (first, left) = Indents(level);
				para.SetIndents(first, left, para.RightIndent);
			}
			changed++;
		}
		return changed;
	}

	// ---------------------------------------------------------------- typed numbers (simple lists)

	// A typed list marker at a paragraph's start: optional spaces, then "1." "1)" "(1)" "a." "(iv)" "IV" ... or a bullet,
	// then a tab (or, if allowed, spaces).
	public static readonly Regex TypedMarker = new(@"^ *(?:(?<open>\()?(?<n>\d{1,4}|[a-zA-Z]{1,6})(?<close>[.)])?|(?<bullet>[•·▪◦‣\-–*]))(?<gap>\t| +)");

	static int RomanValue(string s) {
		var values = new Dictionary<char, int> { ['i'] = 1, ['v'] = 5, ['x'] = 10, ['l'] = 50, ['c'] = 100, ['d'] = 500, ['m'] = 1000 };
		s = s.ToLowerInvariant();
		if (s.Length == 0 || s.Any(c => !values.ContainsKey(c))) return -1;
		int total = 0;
		for (int i = 0; i < s.Length; i++) {
			int v = values[s[i]];
			total += i + 1 < s.Length && values[s[i + 1]] > v ? -v : v;
		}
		return total;
	}

	static int LetterValue(string s) {
		int total = 0;
		foreach (char c in s.ToLowerInvariant()) { if (c < 'a' || c > 'z') return -1; total = total * 26 + (c - 'a' + 1); }
		return total;
	}

	/// Paragraphs from start to end that begin with typed list numbers or bullets ("1.<tab>", "a)<tab>", "(iv)<tab>",
	/// "•<tab>"), as LibreOffice's RTF leaves its lists, become real list items: the numbering style, punctuation and
	/// start from the typed numbers, the level from the paragraphs' indents (more indent, deeper). Runs of at least two
	/// such paragraphs that follow each other; compound numbers ("1.2") become legal items (AdoptTypedNumbers).
	/// requireTab: only markers followed by a tab (as converters write them); otherwise spaces do too. Returns how many
	/// paragraphs became list items.
	/// indents: where each paragraph's typed number starts (points), for the paragraphs that start with one, in order,
	/// when the editor didn't get the paragraphs' indents (HTML loads drop them); used if it has one for each.
	public static int ConvertTypedLists(RichEditBox editor, int start, int end, bool requireTab, IReadOnlyList<float>? indents = null) {
		int legal = AdoptTypedNumbers(editor);
		var doc = editor.TextDocument;
		var paragraphs = Paragraphs(doc).Where(p => p.End > start && p.Start <= Math.Max(start, end - 1) || p.Start == start).ToList();
		var runs = new List<List<(Paragraph P, Match M, float Indent)>>();
		var run = new List<(Paragraph, Match, float)>();
		int typedCount = paragraphs.Count(p => (p.ListType & 0xFFFF) == tomConstants.tomListNone && TypedMarker.IsMatch(p.Text));
		bool useIndents = indents != null && indents.Count == typedCount;
		int typedIndex = 0;
		foreach (var p in paragraphs) {
			var m = (p.ListType & 0xFFFF) == tomConstants.tomListNone ? TypedMarker.Match(p.Text) : Match.Empty;
			int index = m.Success ? typedIndex++ : -1;
			bool typed = m.Success && (!requireTab || m.Groups["gap"].Value == "\t") && m.Length < p.Text.Length;
			if (typed) {
				var para = doc.Range2(p.Start, p.Start).Para;
				float indent = useIndents ? indents![index] : para.LeftIndent + para.FirstLineIndent;    // where the number starts
				run.Add((p, m, indent));
				continue;
			}
			if (run.Count > 0) { runs.Add(run); run = new(); }
		}
		if (run.Count > 0) runs.Add(run);
		runs = runs.Where(r => r.Count >= 2).ToList();
		if (runs.Count == 0) return legal;
		int count = 0;
		Grouped(editor, () => {
			foreach (var r in Enumerable.Reverse(runs)) {
				// levels: by where the numbers start (more indent, deeper); if they all start at the same place, by the kind
				// of marker, as outlines are written (1. then a) then (i) ...: a new kind goes a level deeper)
				var indents = new List<float>();
				foreach (var x in r.Select(x => x.Indent).OrderBy(i => i))
					if (indents.Count == 0 || x - indents[^1] > 2) indents.Add(x);
				static string KindOf(Match m) => m.Groups["bullet"].Success ? "•"
					: (char.IsDigit(m.Groups["n"].Value[0]) ? "1" : char.IsUpper(m.Groups["n"].Value[0]) ? "A" : "a")
					  + (m.Groups["open"].Success ? "(" : "") + m.Groups["close"].Value;
				var kinds = new List<string>();
				var levels = r.Select(x => {
					if (indents.Count > 1) return Math.Min(8, Math.Max(0, indents.FindLastIndex(i => x.Indent >= i - 2)));
					string kind = KindOf(x.M);
					if (!kinds.Contains(kind)) kinds.Add(kind);
					return Math.Min(8, kinds.IndexOf(kind));
				}).ToList();
				// sub-lists: the items of a level under one parent item; each gets its style from its own markers (so that
				// "i" is a Roman numeral in one and a letter in another) and starts at its first item's number
				var sublistOf = new int[r.Count];
				var open = new int[9];
				int sublists = 0, previousLevel = -1;
				for (int k = 0; k < r.Count; k++) {
					int level = levels[k];
					if (level > previousLevel || previousLevel < 0 || open[level] == 0) open[level] = ++sublists;
					for (int deeper = level + 1; deeper < 9; deeper++) if (level < previousLevel) open[deeper] = 0;
					sublistOf[k] = open[level];
					previousLevel = level;
				}
				var styleOf = new Dictionary<int, int>();
				var startOf = new Dictionary<int, int>();
				foreach (var group in Enumerable.Range(0, r.Count).GroupBy(k => sublistOf[k])) {
					var markers = group.Select(k => r[k].M).ToList();
					int style;
					if (markers.All(m => m.Groups["bullet"].Success)) style = tomConstants.tomListBullet;
					else {
						var numbers = markers.Where(m => m.Groups["n"].Success).Select(m => m.Groups["n"].Value).ToList();
						string first = numbers.FirstOrDefault() ?? "1";
						bool digits = numbers.All(n => char.IsDigit(n[0]));
						bool roman = !digits && numbers.All(n => RomanValue(n) > 0) && (RomanValue(first) == 1 || numbers.Any(n => n.Length > 1));
						bool upper = char.IsUpper(first[0]);
						style = digits ? tomConstants.tomListNumberAsArabic
							: roman ? (upper ? tomConstants.tomListNumberAsUCRoman : tomConstants.tomListNumberAsLCRoman)
							: (upper ? tomConstants.tomListNumberAsUCLetter : tomConstants.tomListNumberAsLCLetter);
					}
					styleOf[group.Key] = style;
					string n0 = markers[0].Groups["n"].Value;
					startOf[group.Key] = style == tomConstants.tomListBullet ? 1
						: style == tomConstants.tomListNumberAsArabic ? (int.TryParse(n0, out int d) ? d : 1)
						: style is tomConstants.tomListNumberAsLCRoman or tomConstants.tomListNumberAsUCRoman ? Math.Max(1, RomanValue(n0))
						: Math.Max(1, LetterValue(n0));
				}
				var items = Enumerable.Range(0, r.Count).Select(k => {
					var m = r[k].M;
					int style = styleOf[sublistOf[k]];
					int punctuation = style == tomConstants.tomListBullet ? 0
						: m.Groups["open"].Success ? tomConstants.tomListParentheses
						: m.Groups["close"].Value == ")" ? 0
						: m.Groups["close"].Value == "." ? tomConstants.tomListPeriod : tomConstants.tomListPlain;
					return (r[k].P, M: m, Level: levels[k], Type: style | punctuation, Start: startOf[sublistOf[k]]);
				}).ToList();
				// the typed numbers go (from the end, as that moves positions), then the items are made front to back (as
				// loaders do; RichEdit's numbering of a list made back to front didn't go on past its sub-items)
				foreach (var item in Enumerable.Reverse(items)) doc.Range2(item.P.Start, item.P.Start + item.M.Length).Text = "";
				int removed = 0;
				foreach (var item in items) {
					var para = doc.Range2(item.P.Start - removed, item.P.Start - removed).Para;
					removed += item.M.Length;
					para.ListType = item.Type;
					para.ListLevelIndex = item.Level;
					para.ListStart = item.Start;
					var (first, left) = IndentsFor(item.Level, 18);    // (its level's: the document's list levels, else automatic)
					para.SetIndents(first, left, para.RightIndent);
					count++;
				}
			}
		});
		return legal + count;
	}

	// ---------------------------------------------------------------- typed numbers (legal)

	/// Paragraphs that start with typed compound numbers ("1.2<tab>Text", as Word's and LibreOffice's RTF of a legal list
	/// leaves them, without list formatting) become legal items: runs of such paragraphs where at least one number has
	/// two or more parts. Returns how many paragraphs changed.
	public static int AdoptTypedNumbers(RichEditBox editor) {
		var doc = editor.TextDocument;
		var paragraphs = Paragraphs(doc);
		var runs = new List<List<Paragraph>>();
		var run = new List<Paragraph>();
		foreach (var p in paragraphs) {
			bool typed = !p.Legal && Regex.IsMatch(p.Text, @"^ *\d+(?:\.\d+)*\.?\t");    // (LibreOffice's RTF puts a space first)
			if (typed) { run.Add(p); continue; }
			if (run.Count > 0) { runs.Add(run); run = new List<Paragraph>(); }
		}
		if (run.Count > 0) runs.Add(run);
		runs = runs.Where(r => r.Any(p => Regex.IsMatch(p.Text, @"^ *\d+\.\d+"))).ToList();
		if (runs.Count == 0) return 0;
		Grouped(editor, () => {
			foreach (var r in Enumerable.Reverse(runs)) {
				for (int i = r.Count - 1; i >= 0; i--) {
					var p = r[i];
					var m = Regex.Match(p.Text, @"^ *(\d+(?:\.\d+)*)\.?\t");
					int level = Math.Min(8, m.Groups[1].Value.Count(c => c == '.'));
					int top = int.Parse(m.Groups[1].Value.Split('.')[0]);
					doc.Range2(p.Start, p.Start + m.Length).Text = "";
					// each top-level item gives its number: one that isn't the next number is a restart
					SetLegal(doc.Range2(p.Start, p.Start).Para, level, level == 0 ? Loaded + top : 0);
				}
			}
			Normalize(doc);
		});
		return runs.Sum(r => r.Count);
	}
}

/// The formats of a numbered style's numbers (ParagraphStyle.Numbering): the count of its own level as a number, a
/// letter or a Roman numeral, with the punctuation around it ("1.", "(a)", "iv)"), or Legal: the counts down to its
/// level, "1.2.1", as legal numbering writes them.
public static class NumberFormats
{
	public const string Legal = "1.1";

	public static readonly string[] All = { "1.", "1)", "(1)", "a.", "a)", "(a)", "A.", "A)", "(A)", "i.", "i)", "(i)", "I.", "I)", "(I)", Legal };

	public static bool Valid(string format) => All.Contains(format);

	/// The number n in a format: "(a)" and 2 give "(b)".
	public static string Render(string format, int n) {
		int at = format.IndexOfAny(new[] { '1', 'a', 'A', 'i', 'I' });
		if (at < 0) return n.ToString();
		string count = format[at] switch {
			'a' => Letters(n).ToLowerInvariant(),
			'A' => Letters(n),
			'i' => Roman(n).ToLowerInvariant(),
			'I' => Roman(n),
			_ => n.ToString()
		};
		return format[..at] + count + format[(at + 1)..];
	}

	// A, B ... Z, AA, BB ... ZZ, AAA ... (as Word counts in letters); 0 and below as a number.
	static string Letters(int n) => n <= 0 ? n.ToString() : new string((char)('A' + (n - 1) % 26), (n - 1) / 26 + 1);

	static string Roman(int n) {
		if (n <= 0 || n > 3999) return n.ToString();
		var parts = new (int Value, string Text)[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
		var text = new System.Text.StringBuilder();
		foreach (var (value, roman) in parts) while (n >= value) { text.Append(roman); n -= value; }
		return text.ToString();
	}
}
