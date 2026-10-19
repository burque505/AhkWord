// Page view for RTE Writer's Print Layout (rte_writer36.cs, through #load): the printed pages as sheets of their own,
// with their headers and footers and a gap between them, as in Word, while the text is edited as before.
// RichEdit shows a document as one long page and can't leave a gap between two of its lines, but a line whose text is
// raised (ITextFont.Position, as for superscript) is that much taller, its text at its top (lowering is ignored): so
// each page's last printed line is raised by the rest of the sheet (what the page's lines leave of the text area), the
// bottom margin, a band between the sheets and the next sheet's top margin, and the space shows after it, in the middle
// of a paragraph too. The line breaks stay as they were, so the pages do too. Page 1's top margin is space before the
// first paragraph, and the last page's bottom margin a raise of the document's last line.
// The gaps are the view's, never the document's: they are put in with undo suspended (Modified as it was), raised by
// a value of their own (whole points and 0.35: GapMark), so that they are found wherever editing has moved them, and
// taken out before anything reads the document as a file does (saving, HTML, Markdown, copying, printing and counting
// pages: the editor's WithoutGaps), then put back as they were. They are worked out again a moment after typing stops,
// from where the printed pages start. A page that ends inside a table gets no gap (raising a cell's text would make
// its row taller), and nor does one whose last line is spaced "exactly" (no raise makes it taller): Print Layout's
// "Page N" line marks it. Where the line spacing multiplies a raise (1.5 lines, double), the raise is divided by it.
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
	/// A gap after a page: its last line (raised), where the next page starts (-1: the document's end), the raise and its
	/// parts (points): the rest of the sheet, the bottom margin, the band, the next sheet's top margin. Page: the page
	/// that ends there (1 for the first).
	public sealed record Gap(int LineStart, int LineEnd, int Next, float Raise, float Slack, float Bottom, float Band, float Top, int Page)
	{
		/// How tall the gap is (points): the raise is less where the paragraph's line spacing multiplies it.
		public float Height => Slack + Bottom + Band + Top;

		/// Its line as it is now: a range RichEdit moves with the text edited before it (LineStart, LineEnd and Next
		/// are where it was when it was put in, until the gaps are worked out again).
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

	// A raise of RTE Writer's own: whole points and 0.35 (7 twips).
	const float GapMark = 0.35f;
	static bool Ours(float position) => position > 0.5f && Math.Abs(position - MathF.Floor(position) - GapMark) < 0.02f;
	static float Marked(float points) => MathF.Floor(Math.Clamp(points, 1, 1580)) + GapMark;

	readonly RichEditBox editor;
	public PageGaps(RichEditBox editor) => this.editor = editor;

	/// The gaps as they are now (empty: none).
	public List<Gap> Gaps { get; private set; } = new();

	/// Page 1's top margin, as space before the first paragraph (points; 0: none), and that paragraph's own space.
	public float TopExtra { get; private set; }
	float topOriginal;

	/// The number of pages they were worked out for.
	public int Pages { get; private set; }

	public bool Applied => Gaps.Count > 0 || TopExtra > 0;

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

	/// Where a character is in the editor (its line's top), whether or not it is in view. (RichTextBox's
	/// GetPositionFromCharIndex gives (0, 0) past its Text's length, which is shorter than the document when it has
	/// fields: a gap drawn there would stay put while the text scrolls.)
	public Point PositionOf(int cp) {
		int storyEnd = editor.TextDocument.Range2(0, 0).StoryLength;
		var point = new Point();
		SendPoint(editor.Handle, EM_POSFROMCHAR, ref point, (IntPtr)Math.Clamp(cp, 0, Math.Max(0, storyEnd - 1)));
		return point;
	}

	/// Takes the gaps out: every raise of its own (wherever editing has moved it) and page 1's space (also in the
	/// paragraphs that Enter made of the first, which have its format). Keeps what they were, for Reapply.
	public void Remove() {
		Unlift();    // (the lifted gap's raise back first, so that all the gaps are found)
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var run = doc.Range2(0, 0);
		var found = new List<(int Start, int End, float Raise)>();
		while (run.Start < storyEnd) {    // (the final paragraph mark too: the last page's raise has it)
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
			if (run.End >= storyEnd) break;    // (RichEdit puts no range after the final paragraph mark)
			run.SetRange(run.End, run.End);
		}
		// where the gaps are now (editing since they were worked out moves them): Reapply puts them back there
		var gaps = Gaps;
		if (found.Count == gaps.Count)
			gaps = gaps.Select((gap, i) => gap with { LineStart = found[i].Start, LineEnd = found[i].End, Next = gap.Next < 0 ? -1 : found[i].End + (gap.Next - gap.LineEnd), Raise = found[i].Raise }).ToList();    // (the next line: after the paragraph mark the raise leaves out)
		if (TopExtra > 0) {
			float inflated = topOriginal + TopExtra;
			foreach (var (start, _) in StyleEngine.Paragraphs(doc, 0, Math.Min(storyEnd, 4000))) {
				var para = doc.Range2(start, start).Para;
				if (Math.Abs(para.SpaceBefore - inflated) > 0.05f) break;
				para.SpaceBefore = topOriginal;
			}
		}
		removed = (found.Count == Gaps.Count ? gaps : new List<Gap>(), TopExtra);    // (gaps lost to editing: back when they are worked out again)
		if (found.Count != Gaps.Count) Lost = true;
		Gaps = new();
		TopExtra = 0;
	}

	(List<Gap> Gaps, float Top) removed = (new(), 0);

	/// The gap whose line has the caret (-1: none). Its raise is lifted while the caret is there: RichEdit draws the
	/// caret at the end of a raised line at the line's bottom (the bottom of the gap), and what is typed there, which
	/// isn't raised, there too, so that typing at a page's last line (a new document's only one) went to the bottom
	/// of the page. Meanwhile that page has Print Layout's "Page N" line; the gap comes back when the caret leaves.
	public int Lifted { get; private set; } = -1;
	int liftedLength;

	/// The gap whose line has a position (-1: none).
	public int GapAt(int position) => Gaps.FindIndex(g => g.Track is { } line && position >= line.Start && position <= line.End);

	/// Lifts a gap's raise (index -1: none), putting back the one lifted before.
	public void Lift(int index) {
		if (index == Lifted) return;
		Unlift();
		if (index < 0 || index >= Gaps.Count || Gaps[index].Track is not { } line || line.End <= line.Start) return;
		editor.TextDocument.Range2(line.Start, line.End).Font.Position = 0;
		Lifted = index;
		liftedLength = line.End - line.Start;
		liftedStory = editor.TextDocument.Range2(0, 0).StoryLength;
	}
	int liftedStory;

	/// Puts the lifted gap's raise back, if the document is as it was; if anything was typed or deleted meanwhile (its
	/// line may be longer, or no longer a page's last: Enter at a new document's only line), the gaps are worked out
	/// again instead (Lost), the line left as it is.
	public void Unlift() {
		if (Lifted < 0) return;
		if (Lifted < Gaps.Count && Gaps[Lifted].Track is { } line && line.End > line.Start && line.End - line.Start == liftedLength
			&& editor.TextDocument.Range2(0, 0).StoryLength == liftedStory)
			editor.TextDocument.Range2(line.Start, line.End).Font.Position = Gaps[Lifted].Raise;
		else Lost = true;
		Lifted = -1;
	}

	/// Remove found fewer or more raises than there were gaps (editing joined or deleted their lines), so Reapply can't
	/// put them back: they are to be worked out again. The editor clears it when it has asked for that.
	public bool Lost { get; set; }

	/// Puts back the gaps Remove took out, as they were (the text in between unchanged).
	public void Reapply() {
		var (gaps, top) = removed;
		Put(gaps, top);
	}

	// Raises the gaps' lines and gives page 1 its space.
	void Put(List<Gap> gaps, float top) {
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		string story = doc.Range2(0, storyEnd).Text ?? "";
		var put = new List<Gap>();
		foreach (var gap in gaps)
			if (gap.LineEnd <= storyEnd && gap.LineEnd > gap.LineStart && !TouchesTable(story, gap.LineStart, gap.LineEnd)) {
				doc.Range2(gap.LineStart, gap.LineEnd).Font.Position = gap.Raise;
				put.Add(gap with { Track = doc.Range2(gap.LineStart, gap.LineEnd) });
			}
		Gaps = put;
		Lifted = -1;
		if (top > 0 && storyEnd > 1) {
			var para = doc.Range2(0, 0).Para;
			topOriginal = Math.Max(0, para.SpaceBefore);
			para.SpaceBefore = topOriginal + top;
			TopExtra = top;
		}
	}

	// A gap's line without the paragraph mark that ends it, unless that is all it has (an empty paragraph): where the
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
		if (starts.Count == 0 || story.Length < 2) { Put(gaps, 0); return; }
		for (int k = 1; k < starts.Count; k++) {
			int at = starts[k];
			if (at <= 0 || at >= story.Length || InTable(story, at) || InTable(story, at - 1)) continue;
			var line = doc.Range2(at - 1, at - 1);
			line.Expand(5);    // tomLine
			if (line.End > at) line.End = at;
			if (TouchesTable(story, line.Start, line.End)) continue;
			WithoutParagraphMark(line, story);
			float slack = Math.Max(0, textArea - pages[k - 1].Height);
			if (RaiseFor(line, slack + bottom + BandPoints + top) is not { } raise) continue;    // (spaced exactly: Print Layout's "Page N" line)
			gaps.Add(new Gap(line.Start, line.End, at, raise, slack, bottom, BandPoints, top, k));
		}
		// the last page: the rest of its sheet and its bottom margin, after the document's last line
		int lastChar = storyEnd - 2;    // (before the final paragraph mark)
		if (lastChar >= 0 && !InTable(story, lastChar)) {
			var line = doc.Range2(lastChar, lastChar);
			line.Expand(5);
			if (line.End > storyEnd - 1) line.End = storyEnd - 1;    // (not the story's final paragraph mark)
			WithoutParagraphMark(line, story);
			float slack = Math.Max(0, textArea - pages[^1].Height);
			if (!TouchesTable(story, line.Start, Math.Max(line.End, line.Start + 1)))
				if (RaiseFor(line, slack + bottom) is { } raise)
					gaps.Add(new Gap(line.Start, Math.Max(line.End, line.Start + 1), -1, raise, slack, bottom, 0, 0, starts.Count));
		}
		Put(gaps, InTable(story, 0) ? 0 : top);
	}

	/// Forgets the gaps without touching the document (a new one was loaded, or they were taken out for good).
	public void Forget() {
		Gaps = new();
		TopExtra = 0;
		removed = (new(), 0);
		Lifted = -1;
	}

	/// The document has raises of RTE Writer's own (a gap's, moved by editing or left behind).
	public bool AnyLeft() {
		var doc = editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var run = doc.Range2(0, 0);
		while (run.Start < storyEnd) {    // (the final paragraph mark too: the last page's raise has it)
			run.MoveEnd(13, 1);
			if (run.End <= run.Start) break;
			if (Ours(run.Font.Position)) return true;
			if (run.End >= storyEnd) break;
			run.SetRange(run.End, run.End);
		}
		return false;
	}

	/// What is typed at the caret isn't raised (on a gap's line, it would be).
	public static void PlainTyping(ITextDocument2 doc) {
		var selection = doc.Selection2;
		if (selection.Start == selection.End && Ours(selection.Font.Position)) selection.Font.Position = 0;
	}

	/// The rectangles (the editor's coordinates) of a gap: the rest of its sheet with the footer, the band, the next
	/// sheet's top margin with its header. Bottom of the text above it: where the gap begins.
	public (RectangleF Footer, RectangleF Band, RectangleF Header) Areas(Gap gap, float width, float pixelsPerPoint) {
		float bottom;
		int lineStart = gap.Track?.Start ?? gap.LineStart;    // (where editing has moved it)
		if (gap.Next >= 0) {
			// the next line: where the gap's line ends now (typing on it moves that)
			var line = editor.TextDocument.Range2(lineStart, lineStart);
			line.Expand(5);    // tomLine
			bottom = PositionOf(gap.Track != null ? line.End : gap.Next).Y;
		}
		else {
			float lineHeight = Math.Max(1, editor.TextDocument.Range2(lineStart, lineStart + 1).Font.Size) * 1.2f * pixelsPerPoint;
			bottom = PositionOf(lineStart).Y + lineHeight + gap.Height * pixelsPerPoint;
		}
		float topOfGap = bottom - gap.Height * pixelsPerPoint;
		float footerTop = topOfGap + gap.Slack * pixelsPerPoint, bandTop = footerTop + gap.Bottom * pixelsPerPoint, headerTop = bandTop + gap.Band * pixelsPerPoint;
		return (new RectangleF(0, topOfGap, width, bandTop - topOfGap), new RectangleF(0, bandTop, width, gap.Band * pixelsPerPoint), new RectangleF(0, headerTop, width, bottom - headerTop));
	}
}
