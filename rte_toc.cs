// A table of contents, for AhkWord (AhkWord.cs, through #load): the headings (Heading 1 to the level chosen), each with
// its page as printed, as Word's References > Table of Contents makes it.
//
// RichEdit has no fields, so the table is paragraphs of text between hidden markers (CFE_HIDDEN, RTF \v), as
// cross-references are (rte_crossrefs.cs):
//   [[tocs:1-3]]  at the start of its first paragraph, its options after the colon (Options), and [[/tocs]] at the end
//                 of its last: the table
//   [[toc=_Toc123456789|2]]Heading text<tab>[[pg]]4[[/toc]]   an entry, a paragraph of its own: the name of its
//                 heading's target and its level (for a file that keeps no paragraph styles: Markdown; Restyle), the
//                 heading's text, and its page (nothing after [[pg]] without page numbers)
// Each heading has a target's marker ([[xref:_Toc123456789]], CrossRefs.TargetMark): Ctrl+click on an entry goes there.
// The entries' paragraphs have the styles TOC 1 to TOC 6 (made if the document hasn't got them, as Word's), and a right
// tab at the text's width with a leader for the page numbers. Update (F9) makes the table again from the headings;
// the page numbers alone are brought up to date before saving and printing (UpdatePages).
// Files: RTF (and so Word and LibreOffice files) gets Word's own TOC field, its entries HYPERLINK fields to bookmarks at
// the headings with PAGEREF fields for the numbers, which Word and LibreOffice can update themselves; Word's TOC fields
// are read back as tables. The built-in .docx converter gets the same (IntoDocx, MarkDocx). HTML gets links to anchors
// at the headings, Markdown [Heading](#_Toc123) links; both are read back as tables.
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET.TextObjectModel2;
using W = DocumentFormat.OpenXml.Wordprocessing;

public static class Toc
{
	const string Id = @"[A-Za-z_][A-Za-z0-9_]{0,39}";
	public const string EndMark = "[[/tocs]]";
	public static readonly Regex TableStart = new(@"\[\[tocs:([^\]\r]*)\]\]");
	/// An entry: its heading's target (group 1), its text with the separator (2) and its page number (3).
	public static readonly Regex Entry = new(@"\[\[toc=(" + Id + @")(?:\|\d)?\]\]([^\r]*?)\[\[pg\]\]([^\r\[]*?)\[\[/toc\]\]");
	// An entry's marker with its level (group 1 the name, 2 the level).
	static readonly Regex Leveled = new(@"\[\[toc=(" + Id + @")\|(\d)\]\]");
	static string LevelIn(string marker) => Leveled.Match(marker) is { Success: true } m ? "|" + m.Groups[2].Value : "";
	// An entry as a file can give it: its marker and its end, its page's marker anywhere in it or none.
	static readonly Regex LooseEntry = new(@"\[\[toc=(" + Id + @")(?:\|\d)?\]\]([^\r]*?)\[\[/toc\]\]");

	/// The markers, for CrossRefs.AnyMark (hidden, and left out of plain text).
	public const string Marks = @"\[\[tocs:[^\]\r]*\]\]|\[\[/tocs\]\]|\[\[toc=" + Id + @"(?:\|\d)?\]\]|\[\[pg\]\]|\[\[/toc\]\]";

	/// What the table has when no heading is found (as Word's).
	public const string NoEntries = "No table of contents entries found.";

	// A HYPERLINK field's instruction to a bookmark: Word's \l "_Toc1", LibreOffice's "#_Toc1" (the name: group 1 or 2).
	static readonly Regex LinkInstruction = new(@"^\s*HYPERLINK\s+(?:\\l\s+""?(" + Id + @")""?|""#(" + Id + @")"")");
	static string? LinkName(string instruction) => LinkInstruction.Match(instruction) is { Success: true } m ? (m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value) : null;

	/// The styles of the entries (TOC 1 to TOC 6) and of the table's title.
	public static string StyleName(int level) => "TOC " + Math.Clamp(level, 1, 6);
	public const string TitleStyle = "TOC Heading";

	/// A table's options: the headings' levels it takes (1 to Levels), whether it has page numbers, whether they are
	/// at the right margin (else just after the text), and the leader before them there (dots, dashes, underline, none).
	public sealed record Options(int Levels = 3, bool PageNumbers = true, bool RightAlign = true, string Leader = "dots")
	{
		public static readonly string[] Leaders = { "dots", "dashes", "underline", "none" };

		/// As the marker has them: "1-3", then ",nopages", ",left", ",leader=..." where not the default.
		public string Encode() {
			var sb = new StringBuilder("1-" + Math.Clamp(Levels, 1, 6));
			if (!PageNumbers) sb.Append(",nopages");
			if (!RightAlign) sb.Append(",left");
			if (Leader != "dots") sb.Append(",leader=").Append(Leader);
			return sb.ToString();
		}

		public static Options Parse(string text) {
			var o = new Options();
			foreach (string part in text.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)) {
				var levels = Regex.Match(part, @"^(\d)\s*-\s*(\d)$");
				if (levels.Success) o = o with { Levels = Math.Clamp(int.Parse(levels.Groups[2].Value), 1, 6) };
				else if (part == "nopages") o = o with { PageNumbers = false };
				else if (part == "left") o = o with { RightAlign = false };
				else if (part.StartsWith("leader=", StringComparison.Ordinal) && Leaders.Contains(part[7..])) o = o with { Leader = part[7..] };
			}
			return o;
		}

		// TOM's leader (tomSpaces 0, tomDots 1, tomDashes 2, tomLines 3).
		public int TomLeader => Leader switch { "dashes" => 2, "underline" => 3, "none" => 0, _ => 1 };

		/// Word's TOC field instruction for them: TOC \o "1-3" \h \z \u (\n: no page numbers; \p " ": just after the text).
		public string Instruction() => $"TOC \\o \"1-{Math.Clamp(Levels, 1, 6)}\" \\h \\z \\u" + (!PageNumbers ? " \\n" : "") + (PageNumbers && !RightAlign ? " \\p \" \"" : "");

		/// From a TOC field's instruction (Word's, LibreOffice's).
		public static Options FromInstruction(string instruction) {
			var o = new Options();
			var levels = Regex.Match(instruction, @"\\o\s+""?\s*(\d)\s*-\s*(\d)");
			if (levels.Success) o = o with { Levels = Math.Clamp(int.Parse(levels.Groups[2].Value), 1, 6) };
			if (Regex.IsMatch(instruction, @"\\n\b")) o = o with { PageNumbers = false };
			if (Regex.IsMatch(instruction, @"\\p\b")) o = o with { RightAlign = false };
			return o;
		}
	}

	/// A table in the text: from its [[tocs:]] marker to the end of its [[/tocs]] marker.
	public sealed record Table(int Start, int End, Options Options);

	/// A heading the table takes: where its paragraph starts, its level, its text (as shown, its list number too), its
	/// target's name (null: none yet), and its text without a number RichEdit's list gives it.
	public sealed record Heading(int Start, int Level, string Text, string? Name, string Bare);

	/// Whether text has a table or an entry (quick).
	public static bool Any(string story) => story.Contains("[[toc", StringComparison.Ordinal) || story.Contains("[[pg]]", StringComparison.Ordinal);

	/// The tables in the text, in order (a start marker without its end: to the end of its last entry, or its paragraph).
	public static List<Table> Tables(string story) {
		var tables = new List<Table>();
		if (!Any(story)) return tables;
		foreach (Match m in TableStart.Matches(story)) {
			int end = story.IndexOf(EndMark, m.Index, StringComparison.Ordinal);
			int next = story.IndexOf("[[tocs:", m.Index + m.Length, StringComparison.Ordinal);
			if (end < 0 || next >= 0 && next < end) {
				// (its end marker deleted: the entries straight after it)
				int at = m.Index + m.Length;
				var entries = Entry.Matches(story, at);
				end = at;
				foreach (Match e in entries) {
					if (next >= 0 && e.Index > next) break;
					string between = story[end..e.Index];
					if (between.Trim('\r', ' ', '\t').Length > 0) break;
					end = e.Index + e.Length;
				}
				if (end == at) { int cr = story.IndexOf('\r', at); end = cr < 0 ? story.Length : cr; }
				tables.Add(new Table(m.Index, end, Options.Parse(m.Groups[1].Value)));
				continue;
			}
			tables.Add(new Table(m.Index, end + EndMark.Length, Options.Parse(m.Groups[1].Value)));
		}
		return tables;
	}

	/// The table a position is in (its paragraphs), or null.
	public static Table? TableAt(string story, int position) {
		foreach (var t in Tables(story)) {
			int start = story.LastIndexOf('\r', Math.Max(0, t.Start - 1)) + 1;
			if (t.Start == 0) start = 0;
			int end = story.IndexOf('\r', Math.Max(t.Start, t.End - 1));
			if (end < 0) end = story.Length;
			if (position >= start && position <= end) return t;
		}
		return null;
	}

	/// The heading an entry at a position goes to (its target's name), or null.
	public static string? EntryAt(string story, int position) {
		if (!story.Contains("[[toc=", StringComparison.Ordinal)) return null;
		foreach (Match m in Entry.Matches(story))
			if (position >= m.Index && position <= m.Index + m.Length) return m.Groups[1].Value;
		return null;
	}

	/// The document's headings of levels 1 to levels, in order (not those in a table of contents).
	public static List<Heading> Headings(ITextDocument2 doc, StyleSheet sheet, int levels) {
		var story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var tables = Tables(story);
		var headings = new List<Heading>();
		foreach (var (start, end) in StyleEngine.Paragraphs(doc, 0, story.Length)) {
			int level = StyleEngine.LevelAt(doc, sheet, start);
			if (level < 1 || level > levels) continue;
			if (tables.Any(t => start >= t.Start && start < t.End)) continue;
			var range = doc.Range2(start, end);
			string shown = range.GetText2(tomConstants.tomIncludeNumbering) ?? "";
			string text = EntryText(shown);
			if (text == "") continue;
			var mark = CrossRefs.TargetMark.Match(range.Text ?? "");
			headings.Add(new Heading(start, level, text, mark.Success ? mark.Groups[1].Value : null, EntryText(range.Text ?? "")));
		}
		return headings;
	}

	// A heading's text for its entry: without the markers (a footnote with its number), pictures, line breaks and tabs.
	static string EntryText(string shown) {
		string text = Footnotes.Any(shown) ? Footnotes.Note.Replace(shown, "") : shown;
		text = CrossRefs.PlainText(text);
		text = Regex.Replace(text, "[\uFFFC\u0007]", "");
		text = Regex.Replace(text, @"[\s\u000B]+", " ").Trim();
		return text.Replace("[[", "[ [").Replace("]]", "] ]");
	}

	/// The headings' targets, each put in where it has none (from the last: the positions before stay), at the end of
	/// its text: a paragraph or a page break put in at its start (before it, as often) leaves it where it is.
	static void EnsureTargets(ITextDocument2 doc, List<Heading> headings) {
		var story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		for (int i = headings.Count - 1; i >= 0; i--) {
			if (headings[i].Name != null) continue;
			string name;
			do name = "_Toc" + Random.Shared.Next(100_000_000, 999_999_999); while (story.Contains(name, StringComparison.Ordinal));
			story += name;    // (no two alike)
			var paragraph = doc.Range2(headings[i].Start, headings[i].Start);
			paragraph.Expand(tomConstants.tomParagraph);
			int at = paragraph.Start + (paragraph.Text ?? "").TrimEnd('\r').Length;
			CrossRefs.Hidden(doc, at, "[[xref:" + name + "]]");
			targetsAdded.Add((at, name.Length + 9));    // ("[[xref:" and "]]" around the name)
		}
	}

	// The targets EnsureTargets put in last (where, how long), and how much text they put in before a position.
	static readonly List<(int At, int Length)> targetsAdded = new();
	static int TargetsBefore(ITextDocument2 doc, int position) => targetsAdded.Where(t => t.At < position).Sum(t => t.Length);

	/// The entries' styles (TOC 1 to TOC 6) and the title's (TOC Heading), put in the sheet where it hasn't got them:
	/// the text's font, each level 11 points further in (as Word's), 5 points after; the title as Heading 1, not a heading.
	public static void EnsureStyles(StyleSheet sheet) {
		for (int level = 1; level <= 6; level++) {
			if (sheet.Find(StyleName(level)) != null) continue;
			var style = sheet.Normal.Copy();
			style.Name = StyleName(level);
			style.Level = 0;
			style.Next = "Normal";
			style.LeftIndent = (level - 1) * 11;
			style.FirstLine = 0;
			style.Before = 0;
			style.After = 5;
			style.Numbering = "";
			style.Contextual = false;
			sheet.Styles.Add(style);
		}
		if (sheet.Find(TitleStyle) == null) {
			var title = (sheet.Find("Heading 1") ?? sheet.Normal).Copy();
			title.Name = TitleStyle;
			title.Level = 0;
			title.Next = "Normal";
			title.Numbering = "";
			sheet.Styles.Add(title);
		}
	}

	// A table's text (markers and all) for headings, and where its paragraphs start in it with their levels; the hidden
	// parts as (offset, length).
	static (string Text, List<(int At, int Level)> Paragraphs, List<(int At, int Length)> Hidden) Build(List<Heading> headings, Options options) {
		var sb = new StringBuilder();
		var paragraphs = new List<(int, int)>();
		var hidden = new List<(int, int)>();
		void Mark(string text) { hidden.Add((sb.Length, text.Length)); sb.Append(text); }
		Mark("[[tocs:" + options.Encode() + "]]");
		if (headings.Count == 0) {
			paragraphs.Add((0, 1));
			sb.Append(NoEntries);
		}
		for (int i = 0; i < headings.Count; i++) {
			if (i > 0) sb.Append('\r');
			paragraphs.Add((i == 0 ? 0 : sb.Length, headings[i].Level));
			Mark("[[toc=" + headings[i].Name + "|" + headings[i].Level + "]]");
			sb.Append(headings[i].Text);
			if (options.PageNumbers) sb.Append(options.RightAlign ? '\t' : ' ');
			Mark("[[pg]]");
			if (options.PageNumbers) sb.Append('0');    // (its page: UpdatePages, once the table is in)
			Mark("[[/toc]]");
		}
		Mark(EndMark);
		return (sb.ToString(), paragraphs, hidden);
	}

	/// Inserts a table at a position: in place of its paragraph if that is empty, before it if the position is at its
	/// start (or only spaces before it), else there, the paragraph split (as Word: a Word document's paragraph can go on
	/// after a page break, the next heading's text after it), with a title above it ("": none) in TOC Heading; returns
	/// where the table starts. Its page numbers: UpdatePages after.
	/// textWidth: the text's width on the page (points), for the right tab.
	public static int Insert(ITextDocument2 doc, StyleSheet sheet, int position, Options options, string title, float textWidth) {
		EnsureStyles(sheet);
		var headings = Headings(doc, sheet, options.Levels);
		targetsAdded.Clear();
		EnsureTargets(doc, headings);
		headings = Headings(doc, sheet, options.Levels);
		// (the targets went in: the position moves with the text before it)
		position = Math.Min(position + TargetsBefore(doc, position), doc.Range2(0, 0).StoryLength - 1);
		// (the paragraph from the text: TOM's paragraph ends at a line break too)
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		int paragraphStart = position == 0 ? 0 : story.LastIndexOf('\r', Math.Max(0, position - 1)) + 1;
		int paragraphEnd = story.IndexOf('\r', position);
		string paragraphText = story[paragraphStart..(paragraphEnd < 0 ? story.Length : paragraphEnd)];
		bool empty = CrossRefs.PlainText(paragraphText).Trim(' ', '\t').Length == 0;
		int at = paragraphStart;
		bool split = false;
		if (!empty && position > paragraphStart && CrossRefs.PlainText(story[paragraphStart..position]).Trim(' ', '\t').Length > 0) {
			at = position;
			split = true;
		}
		var (text, paragraphs, hidden) = Build(headings, options);
		title = Regex.Replace(title ?? "", @"[\r\n]+", " ").Trim();
		if (split) {
			// (the paragraph split first, on its own: put in with the table's text, a paragraph mark just after a line
			// break was lost)
			var mark = doc.Range2(at, at);
			mark.Text = "\r";
			at++;
		}
		string lead = title == "" ? "" : title + "\r";
		string all = lead + text + (empty ? "" : "\r");
		Put(doc, at, at, all, hidden.Select(h => (h.At + lead.Length, h.Length)));
		if (title != "") {
			StyleEngine.Apply(doc, sheet, sheet.Find(TitleStyle)!, at, at);
			ClearTabs(doc, at);
		}
		Format(doc, sheet, at + lead.Length, paragraphs, options, textWidth);
		return at + lead.Length;
	}

	/// Makes a table again from the headings (its options as given); returns where it starts. Its page numbers:
	/// UpdatePages after.
	public static int Rebuild(ITextDocument2 doc, StyleSheet sheet, Table table, Options options, float textWidth) {
		EnsureStyles(sheet);
		var headings = Headings(doc, sheet, options.Levels);
		EnsureTargets(doc, headings);
		// (the targets went in before or after the table: where it is now)
		var story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var now = Tables(story).FirstOrDefault(t => t.Start >= table.Start) ?? Tables(story).LastOrDefault();
		if (now == null) return -1;
		headings = Headings(doc, sheet, options.Levels);
		var (text, paragraphs, hidden) = Build(headings, options);
		Put(doc, now.Start, now.End, text, hidden);
		Format(doc, sheet, now.Start, paragraphs, options, textWidth);
		return now.Start;
	}

	// The text in place of start..end, shown but for its hidden parts (offsets from start).
	static void Put(ITextDocument2 doc, int start, int end, string text, IEnumerable<(int At, int Length)> hidden) {
		var range = doc.Range2(start, end);
		range.Text = text;
		range.SetRange(start, start + text.Length);
		range.Font.Hidden = tomConstants.tomFalse;
		range.Font.Superscript = tomConstants.tomFalse;
		foreach (var (at, length) in hidden) doc.Range2(start + at, start + at + length).Font.Hidden = tomConstants.tomTrue;
	}

	// Each entry's paragraph its level's style and the page numbers' tab.
	static void Format(ITextDocument2 doc, StyleSheet sheet, int start, List<(int At, int Level)> paragraphs, Options options, float textWidth) {
		foreach (var (at, level) in Enumerable.Reverse(paragraphs)) {
			var style = sheet.Find(StyleName(level))!;
			StyleEngine.Apply(doc, sheet, style, start + at, start + at);
			var para = doc.Range2(start + at, start + at).Para;
			para.ClearAllTabs();
			// (no leader in the editor: AhkWord draws the table's own, DrawLeader, as RichEdit draws some tabs' leaders and not
			// others; the files get it: ToWordRtf, IntoDocx)
			if (options.PageNumbers && options.RightAlign && textWidth > 20)
				LineSpacing.AddTab(para, textWidth - Reserve(style.Size), 2, 0);    // tomAlignRight, tomSpaces
		}
	}

	static void ClearTabs(ITextDocument2 doc, int at) => doc.Range2(at, at).Para.ClearAllTabs();

	/// How far short of the text's width an entry's right tab is (points): RichEdit puts the number on a line of its own
	/// when the tab and the number after it go past the line's end (it measures as for a left tab), and ignores a tab past
	/// it; so room for a number of three digits. (Word's and LibreOffice's files have it at the margin: ToWordRtf, FitTabs.)
	public static float Reserve(float size) => Math.Max(24f, size * 2f);

	/// Entries' right tabs that a file put at the margin (Word's, LibreOffice's, AhkWord's own .rtf) put Reserve short of
	/// it, without their leader (AhkWord draws the table's). Returns how many were changed.
	public static int FitTabs(ITextDocument2 doc, float textWidth) {
		var story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (textWidth <= 20 || !story.Contains("[[toc=", StringComparison.Ordinal)) return 0;
		int moved = 0;
		foreach (Match m in Entry.Matches(story)) {
			var para = doc.Range2(m.Index, m.Index).Para;
			float size = Math.Max(1, doc.Range2(m.Groups[2].Index, m.Groups[2].Index + 1).Font.Size);
			float wanted = textWidth - Reserve(size);
			for (int i = para.TabCount - 1; i >= 0; i--) {
				para.GetTab(i, out float position, out int alignment, out int leader);
				if (alignment != 2 || position <= wanted + 0.5f && leader == 0) continue;
				para.DeleteTab(position);
				LineSpacing.AddTab(para, Math.Min(position, wanted), 2, 0);
				moved++;
			}
		}
		return moved;
	}

	/// After loading a file that keeps no paragraph styles (Markdown): each entry not in its level's style (its marker
	/// has the level) given it, with the page numbers' tab. Returns how many.
	public static int Restyle(ITextDocument2 doc, StyleSheet sheet, float textWidth) {
		var story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (!story.Contains("[[toc=", StringComparison.Ordinal)) return 0;
		var tables = Tables(story);
		var wrong = new List<(int At, int Level, Options Options)>();
		foreach (Match m in Leveled.Matches(story)) {
			int level = Math.Clamp(int.Parse(m.Groups[2].Value), 1, 6);
			if (string.Equals(StyleEngine.StyleAt(doc, sheet, m.Index).Name, StyleName(level), StringComparison.OrdinalIgnoreCase)) continue;
			var options = tables.LastOrDefault(t => t.Start <= m.Index && m.Index < t.End)?.Options ?? new Options();
			var paragraph = doc.Range2(m.Index, m.Index);
			paragraph.StartOf(tomConstants.tomParagraph, 0);
			wrong.Add((paragraph.Start, level, options));
		}
		if (wrong.Count == 0) return 0;
		EnsureStyles(sheet);
		foreach (var (at, level, options) in Enumerable.Reverse(wrong)) Format(doc, sheet, at, new List<(int, int)> { (0, level) }, options, textWidth);
		return wrong.Count;
	}

	/// A leader RichEdit doesn't draw (it keeps a tab's leader, but draws none): an entry's tab (where its text ends),
	/// where its page number starts, and the leader's kind (dots, dashes, underline).
	public sealed record Leader(int Tab, int Number, string Kind);

	/// The entries' leaders (those of tables with page numbers at the right margin, and a leader).
	public static List<Leader> Leaders(string story) {
		var leaders = new List<Leader>();
		if (!story.Contains("[[toc=", StringComparison.Ordinal)) return leaders;
		var tables = Tables(story);
		foreach (Match m in Entry.Matches(story)) {
			var table = tables.LastOrDefault(t => t.Start <= m.Index && m.Index < t.End);
			var options = table?.Options ?? new Options();
			if (!options.PageNumbers || !options.RightAlign || options.Leader == "none" || m.Groups[3].Length == 0) continue;
			int tab = m.Groups[2].Value.LastIndexOf('\t');
			if (tab < 0) continue;
			leaders.Add(new Leader(m.Groups[2].Index + tab, m.Groups[3].Index, options.Leader));
		}
		return leaders;
	}

	/// Draws a leader from x1 (the text's end) to x2 (the number's start) on a baseline, for text em high (in g's units).
	public static void DrawLeader(Graphics g, string kind, float x1, float x2, float baseline, float em, Color color) {
		float gap = em * 0.2f;
		float from = x1 + gap, to = x2 - gap;
		if (to - from < em * 0.5f) return;
		switch (kind) {
			case "dashes": {
				using var pen = new Pen(color, Math.Max(0.5f, em / 16));
				float y = baseline - em * 0.3f, dash = em * 0.32f, step = em * 0.55f;
				for (float x = from; x + dash <= to; x += step) g.DrawLine(pen, x, y, x + dash, y);
				break;
			}
			case "underline": {
				using var pen = new Pen(color, Math.Max(0.5f, em / 16));
				g.DrawLine(pen, from, baseline + em * 0.05f, to, baseline + em * 0.05f);
				break;
			}
			default: {
				// (dots as a period's, a period's width apart: as Word's dot leader)
				using var brush = new SolidBrush(color);
				float d = Math.Max(0.8f, em * 0.09f), step = em * 0.28f;
				float start = to - (float)Math.Floor((to - from) / step) * step;    // (lined up from the number, as on every line)
				for (float x = start; x <= to - d; x += step) g.FillEllipse(brush, x, baseline - d, d, d);
				break;
			}
		}
	}

	/// Brings the entries' page numbers up to date (pageOf: the page a position is printed on, 1 the first). Returns how
	/// many changed.
	public static int UpdatePages(ITextDocument2 doc, Func<int, int> pageOf) {
		var story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (!story.Contains("[[toc=", StringComparison.Ordinal)) return 0;
		var tables = Tables(story);
		var targets = new Dictionary<string, int>(StringComparer.Ordinal);
		foreach (Match t in CrossRefs.TargetMark.Matches(story)) targets.TryAdd(t.Groups[1].Value, t.Index);
		int changed = 0;
		foreach (Match m in Entry.Matches(story).Reverse()) {
			var table = tables.LastOrDefault(t => t.Start <= m.Index);
			bool pages = table == null || table.End < m.Index || table.Options.PageNumbers;
			string wanted = !pages ? "" : targets.TryGetValue(m.Groups[1].Value, out int at) ? pageOf(at).ToString(System.Globalization.CultureInfo.InvariantCulture) : "?";
			if (m.Groups[3].Value == wanted) continue;
			var number = doc.Range2(m.Groups[3].Index, m.Groups[3].Index + m.Groups[3].Length);
			number.Text = wanted;
			number.SetRange(m.Groups[3].Index, m.Groups[3].Index + wanted.Length);
			number.Font.Hidden = tomConstants.tomFalse;    // (an empty one took the marker's format)
			changed++;
		}
		return changed;
	}

	/// The first table's entries, one a line: the level (its style's), a tab, the text, a tab, the page.
	public static string Listing(ITextDocument2 doc, StyleSheet sheet) {
		var story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (Tables(story).FirstOrDefault() is not { } table) return "";
		var lines = new List<string>();
		foreach (Match m in Entry.Matches(story[..table.End])) {
			if (m.Index < table.Start) continue;
			var style = StyleEngine.StyleAt(doc, sheet, m.Index).Name;
			var level = Regex.Match(style, @"^TOC (\d)$", RegexOptions.IgnoreCase);
			string text = m.Groups[2].Value.TrimEnd('\t', ' ');
			lines.Add((level.Success ? level.Groups[1].Value : "1") + "\t" + text + "\t" + m.Groups[3].Value);
		}
		return string.Join("\n", lines);
	}

	/// After loading (Word's RTF, HTML, Markdown, the built-in reader): a table's end marker at the start of a paragraph
	/// (Word's field ends after its last paragraph) put at the end of the one before, and each entry's separator before
	/// its number as its table has it (a tab, for the right tab; a space).
	public static void AfterLoad(ITextDocument2 doc) {
		var story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (!Any(story)) return;
		for (int i = story.LastIndexOf(EndMark, StringComparison.Ordinal); i >= 0; i = i == 0 ? -1 : story.LastIndexOf(EndMark, i - 1, StringComparison.Ordinal)) {
			int before = i;
			while (before > 0 && story[before - 1] == ' ') before--;
			if (before == 0 || story[before - 1] != '\r') continue;
			doc.Range2(i, i + EndMark.Length).Text = "";
			var mark = doc.Range2(before - 1, before - 1);
			mark.Text = EndMark;
			mark.SetRange(before - 1, before - 1 + EndMark.Length);
			mark.Font.Hidden = tomConstants.tomTrue;
		}
		// an entry without its page's marker (LibreOffice writes no PAGEREF; its number after the text): the marker before
		// a number at its end; one whose marker isn't just before the number (LibreOffice's PAGEREF around the whole text):
		// put there
		story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var loaded = Tables(story);
		foreach (Match m in LooseEntry.Matches(story).Reverse()) {
			if (loaded.LastOrDefault(t => t.Start <= m.Index && m.Index < t.End) is { Options.PageNumbers: false }) continue;
			string inner = m.Groups[2].Value;
			int innerAt = m.Groups[2].Index;
			int pg = inner.IndexOf("[[pg]]", StringComparison.Ordinal);
			string plain = pg >= 0 ? inner.Remove(pg, 6) : inner;
			// (the page: digits at the end, or a roman number after a space or a tab)
			var number = Regex.Match(plain, @"(?:(?<=[\s\u00A0])[ivxlcdmIVXLCDM]+|[0-9]+)(?=[\s\u00A0]*$)");
			int wanted = number.Success ? number.Index : plain.TrimEnd().Length;
			if (pg == wanted) continue;
			if (pg >= 0) { doc.Range2(innerAt + pg, innerAt + pg + 6).Text = ""; }
			var mark = doc.Range2(innerAt + wanted, innerAt + wanted);
			mark.Text = "[[pg]]";
			mark.SetRange(innerAt + wanted, innerAt + wanted + 6);
			mark.Font.Hidden = tomConstants.tomTrue;
		}
		// each entry's level in its marker where a file gave none (RTF, Word's and LibreOffice's: their paragraphs' styles
		// have it), from its paragraph's style, TOC 1 to TOC 6 (for a file that keeps no styles: Markdown; Restyle)
		story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		foreach (Match m in Regex.Matches(story, @"\[\[toc=(" + Id + @")\]\]").Reverse()) {
			string? style = StyleEngine.NameOf(doc.Range2(m.Index, m.Index).Para.Style);
			if (style == null || Regex.Match(style, @"^TOC ([1-6])$", RegexOptions.IgnoreCase) is not { Success: true } level) continue;
			var marker = doc.Range2(m.Index, m.Index + m.Length);
			marker.Text = "[[toc=" + m.Groups[1].Value + "|" + level.Groups[1].Value + "]]";
			marker.Font.Hidden = tomConstants.tomTrue;
		}
		story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var tables = Tables(story);
		foreach (Match m in Entry.Matches(story).Reverse()) {
			var table = tables.LastOrDefault(t => t.Start <= m.Index && m.Index < t.End);
			if (table == null || !table.Options.PageNumbers) continue;
			string text = m.Groups[2].Value;
			string wanted = text.TrimEnd(' ', '\t', '\u00A0') + (table.Options.RightAlign ? "\t" : " ");
			if (wanted == text) continue;
			int trimmed = text.TrimEnd(' ', '\t', '\u00A0').Length;
			var tail = doc.Range2(m.Groups[2].Index + trimmed, m.Groups[2].Index + text.Length);
			tail.Text = table.Options.RightAlign ? "\t" : " ";
			tail.Font.Hidden = tomConstants.tomFalse;
		}
	}

	// ================================================================ RTF (Word's TOC field)

	const string Words = @"((?:\\[a-z]+-?\d* ?)*)", V = @"\\v(?![a-z0-9]) ?";

	/// RTF as RichEdit wrote it, with Word's TOC field, HYPERLINK and PAGEREF fields for the markers (the bookmarks at
	/// the headings: CrossRefs.ToWordRtf, their targets).
	public static string ToWordRtf(string rtf) {
		if (!rtf.Contains("[[toc", StringComparison.Ordinal)) return rtf;
		// the entries' right tab at the margin, as Word has it (in the editor Reserve short of it; 24 points as a rule)
		rtf = Regex.Replace(rtf, @"\\pard(?:(?!\\pard).)*?\[\[tocs:([^\]\r]*)\]\].*?\[\[/tocs\]\]", region => {
			// (the table's leader with it: none in the editor, as AhkWord draws it)
			string leader = Options.Parse(region.Groups[1].Value).Leader switch { "dashes" => @"\tlhyph", "underline" => @"\tlul", "none" => "", _ => @"\tldot" };
			return Regex.Replace(region.Value, @"(?:\\tl[a-z]+)*\\tqr(?:\\tl[a-z]+)*\\tx(\d+)", t => leader + @"\tqr\tx" + (int.Parse(t.Groups[1].Value) + 480));
		}, RegexOptions.Singleline);
		// an entry: its text and number in a HYPERLINK field to its heading's bookmark, the number in a PAGEREF field
		rtf = Regex.Replace(rtf, V + Words + @"\[\[toc=(" + Id + @")(?:\|\d)?\]\]" + Words + @"\\v0 ?((?:[^{}]|\{[^{}]*\})*?)" + V + Words + @"\[\[pg\]\]" + Words + @"\\v0 ?([^{}\\]*?)" + V + Words + @"\[\[/toc\]\]" + Words + @"\\v0 ?",
			m => {
				// (groups: 1 the formatting before the entry's marker, 2 its name, 3 after it; 4 the text; 5 and 6 around [[pg]];
				// 7 the number; 8 and 9 around [[/toc]])
				string name = m.Groups[2].Value;
				string number = m.Groups[7].Value;
				return m.Groups[1].Value + m.Groups[3].Value + @"{\field{\*\fldinst HYPERLINK \\l """ + name + @"""}{\fldrslt " + m.Groups[4].Value
					+ Footnotes.Delimited(m.Groups[5].Value + m.Groups[6].Value)
					// (locked: Word works PAGEREF fields out again as it reads RTF, before it has laid the pages out, and every one
					// said 1; Word's Update Table makes them again, unlocked)
					+ (number.Trim() == "" ? "" : @"{\field\fldlock{\*\fldinst PAGEREF " + name + @" \\h}{\fldrslt " + number + "}}")
					+ "}}" + Footnotes.Delimited(m.Groups[8].Value + m.Groups[9].Value);
			}, RegexOptions.Singleline);
		rtf = Regex.Replace(rtf, V + Words + @"\[\[tocs:([^\]\r]*)\]\]" + Words + @"\\v0 ?",
			m => m.Groups[1].Value + @"{\field{\*\fldinst " + Options.Parse(m.Groups[2].Value).Instruction().Replace(@"\", @"\\") + @"}{\fldrslt " + Footnotes.Delimited(m.Groups[3].Value));
		rtf = Regex.Replace(rtf, V + Words + @"\[\[/tocs\]\]" + Words + @"\\v0 ?", m => m.Groups[1].Value + "}}" + Footnotes.Delimited(m.Groups[2].Value));
		return rtf;
	}

	/// Hidden markers next to each other in one run of hidden text (RichEdit writes them so: "\v [[tocs:1-3]][[toc=_Toc1]]\v0")
	/// each in a run of its own, for the conversions (each takes a marker from \v to \v0).
	public static string SplitHidden(string rtf) => Regex.Replace(rtf, @"\]\](?=\[\[(?:tocs:|/tocs\]|toc=|pg\]|/toc\]|xref|/xref\]|fn))", @"]]\v0 \v ");

	/// Word's (and LibreOffice's) RTF with markers for its TOC fields (their HYPERLINK and PAGEREF fields, and the
	/// bookmarks they go to), before RichEdit loads it.
	public static string FromWordRtf(string rtf) {
		if (!rtf.Contains(@"\fldinst", StringComparison.Ordinal) || !Regex.IsMatch(rtf, @"TOC\s")) return rtf;
		var names = new HashSet<string>(StringComparer.Ordinal);
		string Entries(string instruction, string result) {
			if (LinkName(instruction) is { } link) {
				names.Add(link);
				return @"{\v [[toc=" + link + @"]]}{" + RewriteFields(result, Entries) + @"}{\v [[/toc]]}";
			}
			var page = Regex.Match(instruction, @"^\s*PAGEREF\s+(" + Id + ")");
			if (page.Success) {
				names.Add(page.Groups[1].Value);
				return @"{\v [[pg]]}{" + result + "}";
			}
			return "{" + result + "}";    // (another field in the table: its result)
		}
		string? Tables(string instruction, string result) {
			if (!Regex.IsMatch(instruction, @"^\s*TOC\b")) return null;
			return @"{\v [[tocs:" + Options.FromInstruction(instruction).Encode() + @"]]}{" + RewriteFields(result, Entries) + @"}{\v [[/tocs]]}";
		}
		string done = RewriteFields(rtf, Tables);
		if (done == rtf) return rtf;
		// the bookmarks the entries go to: a target's marker where they start
		done = Regex.Replace(done, @"\{\\\*\\bkmkstart (" + Id + @")\}", m => names.Contains(m.Groups[1].Value) ? @"{\v [[xref:" + m.Groups[1].Value + "]]}" : m.Value);
		return Regex.Replace(done, @"\{\\\*\\bkmkend (" + Id + @")\}", m => names.Contains(m.Groups[1].Value) ? "" : m.Value);
	}

	// The fields in rtf (not those inside them) given to rewrite (their instruction, as text, and their result's RTF);
	// a field it gives null for is left as it is.
	static string RewriteFields(string rtf, Func<string, string, string?> rewrite) {
		var sb = new StringBuilder();
		int at = 0;
		for (int i = rtf.IndexOf(@"{\field", StringComparison.Ordinal); i >= 0; i = rtf.IndexOf(@"{\field", Math.Max(at, i + 1), StringComparison.Ordinal)) {
			if (i < at) continue;
			if (i + 7 < rtf.Length && char.IsLetter(rtf[i + 7])) continue;
			int end = GroupEnd(rtf, i);
			if (end < 0) break;
			string field = rtf[i..(end + 1)];
			int inst = field.IndexOf(@"\fldinst", StringComparison.Ordinal), result = field.IndexOf(@"{\fldrslt", StringComparison.Ordinal);
			if (inst < 0 || result < inst) continue;
			// (RTF's line breaks aren't text: Word breaks a line in the middle of an instruction, "HYPERLINK \\<line>l")
			string instruction = Regex.Replace(field[(inst + 8)..result].Replace("\r", "").Replace("\n", "").Replace(@"\\", "\u0001"), @"\\[a-z]+-?\d* ?|[{}]|\\\*", "").Replace('\u0001', '\\');
			int resultEnd = GroupEnd(field, result);
			if (resultEnd < 0) continue;
			string shown = field[(result + 9)..resultEnd];
			if (shown.StartsWith(' ')) shown = shown[1..];
			if (rewrite(instruction, shown) is not { } replaced) continue;
			sb.Append(rtf, at, i - at).Append(replaced);
			at = end + 1;
		}
		if (at == 0) return rtf;
		return sb.Append(rtf, at, rtf.Length - at).ToString();
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

	// ================================================================ .docx (the built-in converter)

	/// The built-in converter's .docx (made from the HTML: each entry a link to its heading's anchor, its title
	/// "toc:<options>") with Word's TOC field around the entries, each a hyperlink with its page in a PAGEREF field after
	/// a tab. (The bookmarks at the headings: CrossRefs.BookmarksIntoDocx.)
	public static void IntoDocx(string path) {
		using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, true);
		var body = package.MainDocumentPart?.Document?.Body;
		if (body == null) return;
		var links = body.Descendants<W.Hyperlink>().Where(l => l.Tooltip?.Value?.StartsWith("toc:", StringComparison.Ordinal) == true).ToList();
		if (links.Count == 0) return;
		static W.Run Field(W.FieldCharValues type) => new(new W.FieldChar { FieldCharType = type });
		static W.Run Code(string text) => new(new W.FieldCode(" " + text + " ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
		var paragraphs = new List<(W.Paragraph Paragraph, string Options)>();
		foreach (var link in links) {
			string options = link.Tooltip!.Value![4..].Split('|')[0];
			link.Tooltip = null;
			string name = link.Anchor?.Value ?? "";
			// the link's runs: the text, then the page (the last run, after the page's span): a tab and a PAGEREF field
			var runs = link.Elements<W.Run>().ToList();
			var o = Options.Parse(options);
			if (o.PageNumbers && runs.Count >= 2 && Regex.IsMatch(runs[^1].InnerText.Trim(), @"^\S+$")) {
				var last = runs[^1];
				string number = last.InnerText.Trim();
				var properties = last.RunProperties?.CloneNode(true) as W.RunProperties;
				foreach (var text in runs[^2].Elements<W.Text>()) { text.Text = text.Text.TrimEnd(); }
				W.Run Styled(W.Run run) { if (properties != null) run.PrependChild((W.RunProperties)properties.CloneNode(true)); return run; }
				last.InsertBeforeSelf(Styled(new W.Run(o.RightAlign ? new W.TabChar() : new W.Text(" ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })));
				last.InsertBeforeSelf(Styled(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin, FieldLock = true })));    // (locked, as in RTF)
				last.InsertBeforeSelf(Styled(Code("PAGEREF " + name + " \\h")));
				last.InsertBeforeSelf(Styled(Field(W.FieldCharValues.Separate)));
				last.InsertBeforeSelf(Styled(new W.Run(new W.Text(number))));
				last.InsertBeforeSelf(Styled(Field(W.FieldCharValues.End)));
				last.Remove();
			}
			if (link.Ancestors<W.Paragraph>().FirstOrDefault() is { } paragraph && (paragraphs.Count == 0 || paragraphs[^1].Paragraph != paragraph)) {
				paragraphs.Add((paragraph, options));
				// (the right tab at the margin, with the table's leader: in the editor short of it, without)
				foreach (var tab in paragraph.ParagraphProperties?.Tabs?.Elements<W.TabStop>() ?? Enumerable.Empty<W.TabStop>()) {
					if (tab.Val?.Value != W.TabStopValues.Right) continue;
					tab.Position = (tab.Position?.Value ?? 0) + 480;
					tab.Leader = o.Leader switch { "dashes" => W.TabStopLeaderCharValues.Hyphen, "underline" => W.TabStopLeaderCharValues.Underscore, "none" => W.TabStopLeaderCharValues.None, _ => W.TabStopLeaderCharValues.Dot };
				}
			}
		}
		// the field around each run of the entries' paragraphs (one after another in the body)
		var all = body.Descendants<W.Paragraph>().ToList();
		int k = 0;
		while (k < paragraphs.Count) {
			int first = k;
			while (k + 1 < paragraphs.Count && all.IndexOf(paragraphs[k + 1].Paragraph) == all.IndexOf(paragraphs[k].Paragraph) + 1) k++;
			var start = paragraphs[first].Paragraph;
			var options = Options.Parse(paragraphs[first].Options);
			var begin = new List<W.Run> { Field(W.FieldCharValues.Begin), Code(options.Instruction()), Field(W.FieldCharValues.Separate) };
			var anchor = (DocumentFormat.OpenXml.OpenXmlElement?)start.ParagraphProperties;
			foreach (var run in begin) {
				if (anchor == null) start.PrependChild(run); else anchor.InsertAfterSelf(run);
				anchor = run;
			}
			paragraphs[k].Paragraph.Append(Field(W.FieldCharValues.End));
			k++;
		}
		package.MainDocumentPart!.Document.Save();
	}

	/// A .docx's tables of contents as markers in its text (for the built-in reader, which goes through HTML): Word's
	/// TOC field around its entries, their hyperlinks (w:hyperlink, or HYPERLINK fields) to bookmarks, their PAGEREF
	/// fields, and the bookmarks. Returns whether there were any. (Called by CrossRefs.MarkedDocx on its copy.)
	public static bool MarkDocx(DocumentFormat.OpenXml.Packaging.WordprocessingDocument package) {
		var body = package.MainDocumentPart?.Document?.Body;
		if (body == null) return false;
		var names = new HashSet<string>(StringComparer.Ordinal);
		static void Say(W.Run run, string text) {
			var properties = run.RunProperties?.CloneNode(true);
			run.RemoveAllChildren();
			if (properties != null) run.Append(properties);
			run.Append(new W.Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
		}
		// the complex fields (begin, instruction, separate, result, end: each a run's w:fldChar), in the body's order
		var stack = new Stack<(W.Run Begin, List<W.Run> Code, StringBuilder Instruction, bool Separated)>();
		bool any = false;
		foreach (var run in body.Descendants<W.Run>().ToList()) {
			var type = run.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value;
			if (type == W.FieldCharValues.Begin) { stack.Push((run, new List<W.Run>(), new StringBuilder(), false)); continue; }
			if (stack.Count == 0) continue;
			var top = stack.Peek();
			if (type == W.FieldCharValues.Separate) { stack.Pop(); top.Code.Add(run); stack.Push((top.Begin, top.Code, top.Instruction, true)); continue; }
			if (type == W.FieldCharValues.End) {
				stack.Pop();
				string instruction = top.Instruction.ToString();
				string? start = null, finish = null;
				if (Regex.IsMatch(instruction, @"^\s*TOC\b")) { start = "[[tocs:" + Options.FromInstruction(instruction).Encode() + "]]"; finish = EndMark; }
				else if (Regex.Match(instruction, @"^\s*PAGEREF\s+(" + Id + ")") is { Success: true } page) { names.Add(page.Groups[1].Value); start = "[[pg]]"; finish = ""; }
				else if (LinkName(instruction) is { } link && link.StartsWith("_Toc", StringComparison.Ordinal)) {
					names.Add(link); start = "[[toc=" + link + "]]"; finish = "[[/toc]]";
				}
				if (start == null) continue;
				Say(top.Begin, start);
				foreach (var code in top.Code) code.Remove();
				if (finish == "") run.Remove(); else Say(run, finish!);
				any = true;
				continue;
			}
			if (!top.Separated && run.Elements<W.FieldCode>().Any()) {
				foreach (var code in run.Elements<W.FieldCode>()) top.Instruction.Append(code.Text);
				top.Code.Add(run);
			}
		}
		// hyperlinks to a _Toc bookmark: entries
		foreach (var link in body.Descendants<W.Hyperlink>().ToList()) {
			if (link.Anchor?.Value is not { } anchor || !anchor.StartsWith("_Toc", StringComparison.Ordinal)) continue;
			names.Add(anchor);
			var parent = link.Parent!;
			parent.InsertBefore(new W.Run(new W.Text("[[toc=" + anchor + "]]")), link);
			foreach (var child in link.ChildElements.ToList()) { child.Remove(); parent.InsertBefore(child, link); }
			parent.InsertBefore(new W.Run(new W.Text("[[/toc]]")), link);
			link.Remove();
			any = true;
		}
		foreach (var bookmark in body.Descendants<W.BookmarkStart>().ToList())
			if (bookmark.Name?.Value is { } name && names.Contains(name)) bookmark.InsertAfterSelf(new W.Run(new W.Text("[[xref:" + name + "]]")));
		return any;
	}

	// ================================================================ HTML and Markdown

	// An entry with the options of the table it is in: its heading's name, its text (without the separator), its page.
	static List<(Match Entry, string Options)> EntriesIn(string text, Regex start, Regex entry, Regex end) {
		var found = new List<(Match, string)>();
		string options = new Options().Encode();
		int at = 0;
		while (true) {
			var s = start.Match(text, at);
			var e = entry.Match(text, at);
			if (!e.Success) break;
			if (s.Success && s.Index < e.Index) { options = s.Groups[1].Value; at = s.Index + s.Length; continue; }
			found.Add((e, options));
			at = e.Index + e.Length;
		}
		return found;
	}

	/// HTML from the editor with each entry a link to its heading's anchor (CrossRefs.ToHtml: <a id>), titled with its
	/// table's options, its page in a span inside it; the table's own markers left out.
	public static string ToHtml(string html) {
		if (!html.Contains("[[toc", StringComparison.Ordinal)) return html;
		var entries = EntriesIn(html, TableStart, Entry, new Regex(Regex.Escape(EndMark)));
		var sb = new StringBuilder();
		int at = 0;
		foreach (var (m, options) in entries) {
			sb.Append(html, at, m.Index - at);
			string text = m.Groups[2].Value;
			string trimmed = Regex.Replace(text, @"(\s|&#9;|&nbsp;|&emsp;)+$", "");
			sb.Append($"<a href=\"#{m.Groups[1].Value}\" title=\"toc:{options}{LevelIn(m.Value)}\">{trimmed}");
			if (m.Groups[3].Value != "") sb.Append($"<span class=\"rte-toc-page\"> {m.Groups[3].Value}</span>");
			sb.Append("</a>");
			at = m.Index + m.Length;
		}
		html = sb.Append(html, at, html.Length - at).ToString();
		html = TableStart.Replace(html, "");
		return html.Replace(EndMark, "");
	}

	static readonly Regex HtmlEntry = new(@"<a\s+href=""#(" + Id + @")""\s+title=""toc:([^""]*)""[^>]*>(.*?)(?:<span\s+class=""rte-toc-page""\s*>\s*(.*?)\s*</span>\s*)?</a>", RegexOptions.IgnoreCase | RegexOptions.Singleline);

	/// HTML with markers for the entries ToHtml wrote (and the anchors at their headings), each run of entries in
	/// paragraphs one after another a table, before it is loaded.
	public static string FromHtml(string html) {
		if (!html.Contains("title=\"toc:", StringComparison.OrdinalIgnoreCase)) return html;
		var names = new HashSet<string>(StringComparer.Ordinal);
		var matches = HtmlEntry.Matches(html).ToList();
		if (matches.Count == 0) return html;
		var sb = new StringBuilder();
		int at = 0;
		for (int i = 0; i < matches.Count; i++) {
			var m = matches[i];
			static string OptionsOf(Match e) => e.Groups[2].Value.Split('|')[0];
			bool first = i == 0 || Regex.Replace(html[(matches[i - 1].Index + matches[i - 1].Length)..m.Index], "<[^>]*>|\\s", "") != "" || OptionsOf(matches[i - 1]) != OptionsOf(m);
			bool last = i == matches.Count - 1 || Regex.Replace(html[(m.Index + m.Length)..matches[i + 1].Index], "<[^>]*>|\\s", "") != "" || OptionsOf(matches[i + 1]) != OptionsOf(m);
			names.Add(m.Groups[1].Value);
			sb.Append(html, at, m.Index - at);
			if (first) sb.Append("[[tocs:").Append(OptionsOf(m)).Append("]]");
			string level = m.Groups[2].Value.Contains('|') ? "|" + m.Groups[2].Value.Split('|')[1] : "";
			sb.Append("[[toc=").Append(m.Groups[1].Value).Append(level).Append("]]").Append(m.Groups[3].Value).Append(' ').Append("[[pg]]").Append(m.Groups[4].Value).Append("[[/toc]]");
			if (last) sb.Append(EndMark);
			at = m.Index + m.Length;
		}
		html = sb.Append(html, at, html.Length - at).ToString();
		return Regex.Replace(html, @"<a\s+(?:id|name)=""(" + Id + @")""\s*>\s*</a>", m => names.Contains(m.Groups[1].Value) ? $"[[xref:{m.Groups[1].Value}]]" : m.Value, RegexOptions.IgnoreCase);
	}

	// Markdown escapes the markers' brackets and underscores ("\[\[toc=\_Toc1\]\]").
	const string E = @"\\?";
	static readonly Regex MarkdownStart = new(E + @"\[" + E + @"\[tocs:([^\]\r\n\\]*)" + E + @"\]" + E + @"\]");
	static readonly Regex MarkdownEntry = new(E + @"\[" + E + @"\[toc=((?:\\?[A-Za-z0-9_]){1,40})(?:" + E + @"\|(\d))?" + E + @"\]" + E + @"\](.*?)" + E + @"\[" + E + @"\[pg" + E + @"\]" + E + @"\](.*?)" + E + @"\[" + E + @"\[/toc" + E + @"\]" + E + @"\]");
	static readonly Regex MarkdownEnd = new(E + @"\[" + E + @"\[/tocs" + E + @"\]" + E + @"\]");

	/// Markdown from the editor with each entry a [heading](#name "toc:<options>") link and its page after it; the
	/// table's own markers left out.
	public static string ToMarkdown(string markdown) {
		if (!markdown.Contains("toc", StringComparison.Ordinal)) return markdown;
		var entries = EntriesIn(markdown, MarkdownStart, MarkdownEntry, MarkdownEnd);
		if (entries.Count == 0) return markdown;
		var sb = new StringBuilder();
		int at = 0;
		foreach (var (m, options) in entries) {
			sb.Append(markdown, at, m.Index - at);
			// (groups: 1 the name, 2 the level, 3 the text, 4 the page)
			string text = m.Groups[3].Value.TrimEnd(' ', '\t');
			string name = m.Groups[1].Value.Replace("\\", "");
			string level = m.Groups[2].Success ? "|" + m.Groups[2].Value : "";
			sb.Append($"[{text}](#{name} \"toc:{options.Replace("\\", "")}{level}\")");
			if (m.Groups[4].Value != "") sb.Append(' ').Append(m.Groups[4].Value);
			at = m.Index + m.Length;
		}
		markdown = sb.Append(markdown, at, markdown.Length - at).ToString();
		return MarkdownEnd.Replace(MarkdownStart.Replace(markdown, ""), "");
	}

	/// Markdown with markers for the entries ToMarkdown wrote (and the anchors at their headings), each run of entries
	/// on lines one after another a table, before it is loaded (escaped, so that Markdown keeps them as text).
	public static string FromMarkdown(string markdown) {
		if (!markdown.Contains("\"toc:", StringComparison.Ordinal)) return markdown;
		var names = new HashSet<string>(StringComparer.Ordinal);
		var line = new Regex(@"^(\s*(?:[-*+]\s+|\d+[.)]\s+)?)\[([^\]\r\n]*)\]\(#(" + Id + @")\s+""toc:([^""\r\n]*)""\)[ \t]*(\S*)[ \t]*$", RegexOptions.Multiline);
		var matches = line.Matches(markdown).ToList();
		if (matches.Count == 0) return markdown;
		var sb = new StringBuilder();
		int at = 0;
		for (int i = 0; i < matches.Count; i++) {
			var m = matches[i];
			static string OptionsOf(Match e) => e.Groups[4].Value.Split('|')[0];
			bool first = i == 0 || markdown[(matches[i - 1].Index + matches[i - 1].Length)..m.Index].Trim() != "" || OptionsOf(matches[i - 1]) != OptionsOf(m);
			bool last = i == matches.Count - 1 || markdown[(m.Index + m.Length)..matches[i + 1].Index].Trim() != "" || OptionsOf(matches[i + 1]) != OptionsOf(m);
			string name = m.Groups[3].Value;
			names.Add(name);
			sb.Append(markdown, at, m.Index - at).Append(m.Groups[1].Value);
			if (first) sb.Append(@"\[\[tocs:").Append(OptionsOf(m)).Append(@"\]\]");
			string level = m.Groups[4].Value.Contains('|') ? @"\|" + m.Groups[4].Value.Split('|')[1] : "";
			sb.Append($@"\[\[toc={name.Replace("_", "\\_")}{level}\]\]{m.Groups[2].Value} \[\[pg\]\]{m.Groups[5].Value}\[\[/toc\]\]");
			if (last) sb.Append(@"\[\[/tocs\]\]");
			at = m.Index + m.Length;
		}
		markdown = sb.Append(markdown, at, markdown.Length - at).ToString();
		// (the entries one paragraph each: Markdown would join lines one after another into one)
		markdown = Regex.Replace(markdown, @"(\\\[\\\[/toc\\\]\\\](?:\\\[\\\[/tocs\\\]\\\])?)\r?\n(?=\S)", "$1\n\n");
		return Regex.Replace(markdown, @"<a\s+(?:id|name)=""(" + Id + @")""\s*>\s*</a>", m => names.Contains(m.Groups[1].Value) ? $@"\[\[xref:{m.Groups[1].Value.Replace("_", "\\_")}\]\]" : m.Value, RegexOptions.IgnoreCase);
	}
}
