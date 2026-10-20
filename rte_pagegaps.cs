// Page view for AhkWord's Print Layout (AhkWord.cs, through #load): the printed pages as sheets of their own,
// with their headers and footers and a gap between them, as in Word, while the text is edited as before.
// RichEdit shows a document as one long page and can't leave a gap between two of its lines. So (iteration 42):
//   - where a page starts with a paragraph (most pages do), the gap is space before that paragraph: the rest of the
//     sheet above, the bottom margin, a band between the sheets and the next sheet's top margin. Page 1's top margin
//     is space before the first paragraph too. Space before is ordinary paragraph formatting: the caret and typing
//     behave there as anywhere.
//   - where a page starts in the middle of a paragraph, the line before it is raised (ITextFont.Position, as for
//     superscript): a raised line is that much taller, its text at its top, so the gap shows after it. Text typed in
//     that line is raised as the line is, so that it stays at the top, and the caret with it.
//   - the last page has no gap after it: the end of the document, where typing goes on, is as it is.
// The line breaks stay as they were, so the pages do too. The gaps are the view's, never the document's: they are put
// in with undo suspended (Modified as it was), raised by a value of their own (whole points and 0.35: GapMark) and
// spaced by a value known to them (each paragraph gap's Applied), so that they are found wherever editing has moved
// them, and taken out before anything reads the document as a file does (saving, HTML, Markdown, copying, printing and
// counting pages: the editor's WithoutGaps), then put back as they were. They are worked out again a moment after
// typing stops, from where the printed pages start. A page that starts inside a table gets no gap (Print Layout's
// "Page N" line marks it), and nor does a raised line spaced "exactly" (no raise makes it taller).
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

/// Runs an action when disposed (using: "afterwards").
public sealed class OnDispose(Action action) : IDisposable
{
	public void Dispose() => action();
}

public sealed class PageGaps
{
	/// A gap before a page. Paragraph: space before the paragraph at Next (its own space before is Original, as put in
	/// Applied); else a raise of the line LineStart-LineEnd, the page starting at Next in the middle of a paragraph.
	/// Its parts (points): the rest of the sheet above, the bottom margin, the band, the next sheet's top margin. Page:
	/// the page that ends there (0: before page 1, its top margin alone).
	public sealed record Gap(int LineStart, int LineEnd, int Next, float Raise, float Slack, float Bottom, float Band, float Top, int Page,
		bool Paragraph = false, float Original = 0, float Applied = 0)
	{
		/// How tall the gap is (points).
		public float Height => Slack + Bottom + Band + Top;

		/// Where it is now: a paragraph gap's paragraph start, a raised line's line (RichEdit moves ranges with the text
		/// edited before them; LineStart, LineEnd and Next are where it was when it was put in).
		public ITextRange2? Track { get; init; }
	}

	// The raise that makes a line taller by height points, under its paragraph's line spacing (null: none can: a line
	// spaced "exactly" keeps its height whatever it holds). RichEdit makes a raised line taller by the raise times the
	// spacing where that is a multiple (1.5 lines, double, multiple), and only by what it goes beyond the least height
	// where that is "at least" (a line's natural height: about 1.17 times its font's size).
	static float? RaiseFor(ITextRange2 line, float height) {
		var para = line.Para;
		float spacing = para.LineSpacing;
		switch (para.LineSpacingRule) {
			case 1: return Marked(height / 1.5f);    // tomLineSpace1pt5
			case 2: return Marked(height / 2);       // tomLineSpaceDouble
			case 5: return Marked(height / Math.Max(0.25f, spacing));    // tomLineSpaceMultiple (lines)
			case 4: return null;                     // tomLineSpaceExactly
			case 3: {                                // tomLineSpaceAtLeast (points)
				float natural = Math.Max(1, line.Font.Size) * 1.17f;
				return Marked(height + Math.Max(0, spacing - natural));
			}
			default: return Marked(height);
		}
	}

	/// The band between two sheets, in points.
	public const float BandPoints = 14;

	// A raise of AhkWord's own: whole points and 0.35 (7 twips).
	const float GapMark = 0.35f;
	static bool Ours(float position) => position > 0.5f && Math.Abs(position - MathF.Floor(position) - GapMark) < 0.02f;
	static float Marked(float points) => MathF.Floor(Math.Clamp(points, 1, 1580)) + GapMark;

	readonly RichEditBox editor;
	public PageGaps(RichEditBox editor) => this.editor = editor;

	/// The gaps as they are now (empty: none).
	public List<Gap> Gaps { get; private set; } = new();

	/// Page 1's top margin, as space before the first paragraph (points; 0: none).
	public float TopExtra => Gaps.FirstOrDefault(g => g.Page == 0)?.Top ?? 0;

	/// The number of pages they were worked out for.
	public int Pages { get; private set; }

	public bool Applied => Gaps.Count > 0;

	const char RowStart = '￹', RowEnd = '￻', CellEnd = '\u0007';
	static bool InTable(string story, int at) {
		if (at < 0 || at >= story.Length) return false;
		if (story[at] is RowStart or RowEnd or CellEnd) return true;
		int open = story.LastIndexOf(RowStart, Math.Max(0, at - 1)), close = at > 0 ? story.LastIndexOf(RowEnd, at - 1) : -1;
		return open >= 0 && open > close;
	}

	// A line in a table or touching one: its row's marks (raising a row's end mark raises the whole row).
	static bool TouchesTable(string story, int start, int end) {
		if (InTable(story, start)) return true;
		for (int i = Math.Max(0, start); i < Math.Min(end, story.Length); i++)
			if (story[i] is RowStart or RowEnd or CellEnd) return true;
		return start >= 1 && story[start - 1] is RowEnd;    // (a row's end mark is followed by its own paragraph mark)
	}

	const int EM_POSFROMCHAR = 0x00D6;
	[System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessageW")]
	static extern IntPtr SendPoint(IntPtr hWnd, int msg, ref Point point, IntPtr cp);

	/// Where a character is in the editor (its line's top: a paragraph's, its space before included), whether or not it
	/// is in view. (RichTextBox's GetPositionFromCharIndex gives (0, 0) past its Text's length, which is shorter than the
	/// document when it has fields: a gap drawn there would stay put while the text scrolls.)
	public Point PositionOf(int cp) {
		int storyEnd = editor.TextDocument.Range2(0, 0).StoryLength;
		var point = new Point();
		SendPoint(editor.Handle, EM_POSFROMCHAR, ref point, (IntPtr)Math.Clamp(cp, 0, Math.Max(0, storyEnd - 1)));
		return point;
	}

	// The start of the paragraph a position is in, and of those before and after it.
	int ParagraphStart(int at) {
		var range = editor.TextDocument.Range2(at, at);
		range.StartOf(4, 0);    // tomParagraph
		return range.Start;
	}

	List<int> Neighbours(int start) {
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var starts = new List<int>();
		if (start > 0) starts.Add(ParagraphStart(start - 1));
		starts.Add(start);
		var own = doc.Range2(start, start);
		own.Expand(4);
		if (own.End < storyEnd - 1) starts.Add(own.End);
		return starts;
	}

	static bool Spaced(ITextPara para, float applied) => Math.Abs(para.SpaceBefore - applied) < 0.03f;

	/// Takes the gaps out: every raise of its own (wherever editing has moved it) and every paragraph's space that a gap
	/// gave it (its own space before back), where editing has moved the paragraph, and in the paragraph Enter made of it.
	/// Keeps what they were, for Reapply.
	public void Remove() {
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		// the raised lines
		var run = doc.Range2(0, 0);
		var found = new List<(int Start, int End, float Raise)>();
		if (Gaps.Any(g => !g.Paragraph))
			while (run.Start < storyEnd) {
				run.MoveEnd(13, 1);    // tomCharFormat
				if (run.End <= run.Start) break;
				var font = run.Font;
				float position = font.Position;
				if (Ours(position)) {
					font.Position = 0;
					// (a line with a bold word or a link is several runs: one gap)
					if (found.Count > 0 && found[^1].End == run.Start && Math.Abs(found[^1].Raise - position) < 0.01f) found[^1] = (found[^1].Start, run.End, position);
					else found.Add((run.Start, run.End, position));
				}
				if (run.End >= storyEnd) break;
				run.SetRange(run.End, run.End);
			}
		var lines = Gaps.Where(g => !g.Paragraph).ToList();
		bool linesKept = found.Count == lines.Count;
		var kept = new List<Gap>();
		int lineIndex = 0;
		bool lost = !linesKept;
		foreach (var gap in Gaps) {
			if (!gap.Paragraph) {
				if (linesKept) {
					var (start, end, raise) = found[lineIndex];
					kept.Add(gap with { LineStart = start, LineEnd = end, Next = end, Raise = raise, Track = null });    // (the next page: after the raised line)
				}
				lineIndex++;
				continue;
			}
			// a paragraph's space: where its paragraph is now (and the one Enter made of it, and one joined to it)
			int at = gap.Track is { } track ? Math.Min(track.Start, Math.Max(0, storyEnd - 1)) : gap.Next;
			int own = ParagraphStart(at);
			bool here = false;
			foreach (int start in Neighbours(own)) {
				var para = doc.Range2(start, start).Para;
				if (!Spaced(para, gap.Applied)) continue;
				para.SpaceBefore = gap.Original;
				if (!here) { kept.Add(gap with { LineStart = start, LineEnd = start, Next = start, Track = null }); here = true; }
			}
			if (!here) lost = true;    // (joined to the paragraph before, which has its own spacing: worked out again)
		}
		removed = lost ? new List<Gap>() : kept;
		if (lost) Lost = true;
		Gaps = new();
	}

	List<Gap> removed = new();

	/// Remove couldn't find all the gaps (editing joined or deleted their lines or paragraphs), so Reapply can't put them
	/// back: they are to be worked out again. The editor clears it when it has asked for that.
	public bool Lost { get; set; }

	/// Puts back the gaps Remove took out, as they were (the text in between unchanged).
	public void Reapply() => Put(removed);

	// Raises the line gaps' lines and spaces the paragraph gaps' paragraphs.
	void Put(List<Gap> gaps) {
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		string story = doc.Range2(0, storyEnd).Text ?? "";
		var put = new List<Gap>();
		foreach (var gap in gaps) {
			if (gap.Paragraph) {
				if (gap.Next >= storyEnd || (gap.Next > 0 && story[gap.Next - 1] != '\r')) continue;    // (no longer a paragraph's start)
				var para = doc.Range2(gap.Next, gap.Next).Para;
				float original = Math.Max(0, para.SpaceBefore);
				para.SpaceBefore = original + gap.Height;
				put.Add(gap with { Original = original, Applied = para.SpaceBefore, Track = doc.Range2(gap.Next, gap.Next) });
			}
			else if (gap.LineEnd <= storyEnd && gap.LineEnd > gap.LineStart && !TouchesTable(story, gap.LineStart, gap.LineEnd)) {
				doc.Range2(gap.LineStart, gap.LineEnd).Font.Position = gap.Raise;
				put.Add(gap with { Track = doc.Range2(gap.LineStart, gap.LineEnd) });
			}
		}
		Gaps = put;
	}

	// A line without the paragraph mark that ends it, unless that is all it has (an empty paragraph): where the
	// paragraph's line spacing is a multiple (double, 1.5 lines), RichEdit spaces every line of the paragraph by its
	// mark's height, so a raised mark would make all of them as tall as the gap.
	static void WithoutParagraphMark(ITextRange2 line, string story) {
		if (line.End - 1 > line.Start && line.End - 1 < story.Length && story[line.End - 1] == '\r') line.End--;
	}

	/// Works the gaps out for the printed pages (where each starts, laid out without gaps, and how tall its text is, in
	/// points: DocumentPrinter.PageHeights) and puts them in. textArea, bottom, top: the page's text height and margins.
	public void Apply(List<(int Start, float Height)> pages, float textArea, float bottom, float top) {
		var starts = pages.Select(p => p.Start).ToList();
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		string story = doc.Range2(0, storyEnd).Text ?? "";
		Pages = starts.Count;
		var gaps = new List<Gap>();
		if (starts.Count == 0 || story.Length < 2) { Put(gaps); return; }
		// page 1's top margin: space before the first paragraph (not before a table's first row)
		if (!InTable(story, 0)) gaps.Add(new Gap(0, 0, 0, 0, 0, 0, 0, top, 0, Paragraph: true));
		for (int k = 1; k < starts.Count; k++) {
			int at = starts[k];
			if (at <= 0 || at >= story.Length - 1 || InTable(story, at) || InTable(story, at - 1)) continue;
			float slack = Math.Max(0, textArea - pages[k - 1].Height);
			// a page that starts with a paragraph: space before it
			if (story[at - 1] == '\r' && !(at >= 2 && story[at - 2] is RowEnd)) {
				gaps.Add(new Gap(at, at, at, 0, slack, bottom, BandPoints, top, k, Paragraph: true));
				continue;
			}
			// in the middle of a paragraph: the line before it raised
			var line = doc.Range2(at - 1, at - 1);
			line.Expand(5);    // tomLine
			if (line.End > at) line.End = at;
			if (TouchesTable(story, line.Start, line.End)) continue;
			WithoutParagraphMark(line, story);
			if (RaiseFor(line, slack + bottom + BandPoints + top) is not { } raise) continue;    // (spaced exactly: Print Layout's "Page N" line)
			gaps.Add(new Gap(line.Start, line.End, at, raise, slack, bottom, BandPoints, top, k));
		}
		Put(gaps);
	}

	/// Forgets the gaps without touching the document (a new one was loaded, or they were taken out for good).
	public void Forget() {
		Gaps = new();
		removed = new();
	}

	/// The document has raises of AhkWord's own (a gap's, moved by editing or left behind).
	public bool AnyLeft() {
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var run = doc.Range2(0, 0);
		while (run.Start < storyEnd) {
			run.MoveEnd(13, 1);
			if (run.End <= run.Start) break;
			if (Ours(run.Font.Position)) return true;
			if (run.End >= storyEnd) break;
			run.SetRange(run.End, run.End);
		}
		return false;
	}

	/// The space a gap adds before the paragraph a position is in (points; 0: none): formats read there leave it out.
	public float ExtraAt(int position) {
		if (!Applied) return 0;
		int start = ParagraphStart(position);
		foreach (var gap in Gaps)
			if (gap.Paragraph && gap.Track is { } track && ParagraphStart(Math.Min(track.Start, Math.Max(0, editor.TextDocument.Range2(0, 0).StoryLength - 1))) == start)
				return gap.Height;
		return 0;
	}

	/// Whether any paragraph gap is in a paragraph from start to end.
	public bool AnyParagraphGapIn(int start, int end) {
		if (!Applied) return false;
		int from = ParagraphStart(start);
		return Gaps.Any(g => g.Paragraph && g.Track is { } track && track.Start >= from && track.Start <= Math.Max(end, start));
	}

	/// After an edit: Enter at a paragraph gap's paragraph start gives both paragraphs its space; only the first keeps it
	/// (the page starts there), the other gets its own space back. Returns whether anything changed (check: whether
	/// anything would, changing nothing: cheap, for after every keystroke).
	public bool Tidy(bool check = false) {
		if (!Applied) return false;
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		bool changed = false;
		var gaps = Gaps.ToList();
		for (int i = 0; i < gaps.Count; i++) {
			var gap = gaps[i];
			if (!gap.Paragraph || gap.Track is not { } track) continue;
			int own = ParagraphStart(Math.Min(track.Start, Math.Max(0, storyEnd - 1)));
			var spaced = Neighbours(own).Where(s => Spaced(doc.Range2(s, s).Para, gap.Applied)).ToList();
			if (spaced.Count == 0) continue;
			if (check) { if (spaced.Count > 1) return true; continue; }
			foreach (int other in spaced.Skip(1)) { doc.Range2(other, other).Para.SpaceBefore = gap.Original; changed = true; }
			if (spaced[0] != own) gaps[i] = gap with { Track = doc.Range2(spaced[0], spaced[0]) };
		}
		Gaps = gaps;
		return changed;
	}

	/// What is typed at the caret: in a raised line, raised as it is (text that isn't would sit at the bottom of the
	/// line, the gap's bottom); at its end (where the next page starts), as the text there.
	public void TypingFormat() {
		var doc = editor.TextDocument;
		var selection = doc.Selection2;
		if (selection.Start != selection.End) return;
		int at = selection.Start;
		foreach (var gap in Gaps) {
			if (gap.Paragraph || gap.Track is not { } line) continue;
			if (at >= line.Start && at < line.End) { if (Math.Abs(selection.Font.Position - gap.Raise) > 0.01f) selection.Font.Position = gap.Raise; return; }
		}
		if (Ours(selection.Font.Position)) selection.Font.Position = 0;
	}

	/// The rectangles (the editor's coordinates) of a gap: the rest of its sheet with the footer, the band, the next
	/// sheet's top margin with its header.
	public (RectangleF Footer, RectangleF Band, RectangleF Header) Areas(Gap gap, float width, float pixelsPerPoint) {
		float topOfGap;
		if (gap.Paragraph) {
			// the paragraph's top: its space before begins with the gap
			int at = gap.Track is { } track ? track.Start : gap.Next;
			topOfGap = PositionOf(ParagraphStart(Math.Min(at, Math.Max(0, editor.TextDocument.Range2(0, 0).StoryLength - 1)))).Y;
		}
		else {
			// the next line: where the raised line ends now (typing on it moves that)
			int lineStart = gap.Track?.Start ?? gap.LineStart;
			var line = editor.TextDocument.Range2(lineStart, lineStart);
			line.Expand(5);    // tomLine
			topOfGap = PositionOf(gap.Track != null ? line.End : gap.Next).Y - gap.Height * pixelsPerPoint;
		}
		float footerTop = topOfGap, bandTop = footerTop + (gap.Slack + gap.Bottom) * pixelsPerPoint, headerTop = bandTop + gap.Band * pixelsPerPoint;
		float bottom = topOfGap + gap.Height * pixelsPerPoint;
		return (new RectangleF(0, footerTop, width, bandTop - footerTop), new RectangleF(0, bandTop, width, gap.Band * pixelsPerPoint), new RectangleF(0, headerTop, width, bottom - headerTop));
	}
}
