// The Paragraph dialog, tab stops and the page-keeping rules for RTE Writer (rte_writer33.cs, through #load).
//   - ParagraphFormat: what the Paragraph dialog sets (alignment, indents, spacing, line spacing, the line and page
//     break rules, suppressed line numbers, tab stops), read from the selected paragraphs (a value they don't share is
//     left empty) and given to them (only what was set), or to a style; as "Key=Value" lines for scripts.
//   - PageKeeping: RichEdit lays out printed pages (EM_FORMATRANGE) without Word's rules for where a page may end, so
//     each page's end is moved back where a rule says: page break before, keep lines together, widow and orphan control
//     (no single line of a paragraph alone at the foot or the top of a page) and keep with next, as Word does; a rule
//     that would leave a page empty is set aside, as in Word. PrintedLines gives the printed lines (a hidden RichEdit
//     laid out for the printer, as rte_linenumbers2.cs has it).
//   - ParagraphDialog (Indents and Spacing, Line and Page Breaks, with a preview) and TabsDialog, as Word's.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

/// A tab stop: its position (points from the left indent's edge, as TOM has it), alignment and leader.
public readonly record struct ParagraphTab(float Position, string Alignment, string Leader)
{
	public static readonly string[] Alignments = { "left", "center", "right", "decimal", "bar" };    // (TOM's order)
	public static readonly string[] Leaders = { "none", "dots", "dashes", "underline", "thick", "equals" };

	public override string ToString() => Position.ToString("0.##", CultureInfo.InvariantCulture) + "/" + Alignment + "/" + Leader;

	/// "72/right/dots" (points / alignment / leader; the last two may be left out).
	public static ParagraphTab Parse(string text) {
		var parts = text.Split('/').Select(p => p.Trim()).ToArray();
		if (!ParagraphFormat.TryMeasure(parts[0], false, out float position) || position < 0 || position > 1584)
			throw new ArgumentException($"\"{text}\" is not a tab stop: give its position (72, 1in, 2.5cm), then /left, /center, /right, /decimal or /bar, then /dots, /dashes or /underline.");
		string alignment = parts.Length > 1 && parts[1] != "" ? parts[1].ToLowerInvariant() : "left";
		string leader = parts.Length > 2 && parts[2] != "" ? parts[2].ToLowerInvariant() : "none";
		if (alignment == "centre") alignment = "center";
		if (leader is "dotted") leader = "dots";
		if (leader is "lines" or "line") leader = "underline";
		if (!Alignments.Contains(alignment)) throw new ArgumentException($"A tab stop's alignment is left, center, right, decimal or bar (not \"{parts[1]}\").");
		if (!Leaders.Contains(leader)) throw new ArgumentException($"A tab stop's leader is none, dots, dashes, underline, thick or equals (not \"{parts[2]}\").");
		return new ParagraphTab(position, alignment, leader);
	}
}

public sealed class ParagraphFormat
{
	static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

	// (null: the paragraphs differ, or, given to paragraphs, not to be changed)
	public string? Align;                       // left, center, right, justify
	public float? Left, Right, FirstLine;       // points; FirstLine from the left indent (negative: hanging)
	public float? Before, After;                // points
	public string? LineSpacing;                 // as Rte.SetLineSpacing: "1", "1.15", "exactly 14pt", "at least 12pt"
	public bool? KeepWithNext, KeepTogether, WidowControl, PageBreakBefore, NoLineNumber;
	public bool? Contextual;                    // a style's: no space between paragraphs of the same style
	public List<ParagraphTab>? Tabs;
	public int? Level;                          // the outline level the style gives (0: body text), shown only

	public ParagraphFormat Copy() {
		var copy = (ParagraphFormat)MemberwiseClone();
		copy.Tabs = Tabs?.ToList();
		return copy;
	}

	static bool Is(int tom) => tom == tomConstants.tomTrue;
	static int Tom(bool on) => on ? tomConstants.tomTrue : tomConstants.tomFalse;
	static string AlignText(int tom) => tom switch { 1 => "center", 2 => "right", 3 => "justify", _ => "left" };
	static int AlignOf(string align) => align switch { "center" => 1, "right" => 2, "justify" => 3, _ => 0 };

	/// A paragraph's tab stops.
	public static List<ParagraphTab> TabsOf(ITextPara para) {
		var tabs = new List<ParagraphTab>();
		for (int i = 0; i < para.TabCount; i++) {
			para.GetTab(i, out float position, out int alignment, out int leader);
			tabs.Add(new ParagraphTab(position, ParagraphTab.Alignments[Math.Clamp(alignment, 0, 4)], ParagraphTab.Leaders[Math.Clamp(leader, 0, 5)]));
		}
		return tabs;
	}

	/// The format of the paragraphs from start to end (what they share; the rest left empty).
	public static ParagraphFormat Read(ITextDocument2 doc, StyleSheet sheet, int start, int end) {
		ParagraphFormat? result = null;
		foreach (var (pStart, _) in StyleEngine.Paragraphs(doc, start, end)) {
			var range = doc.Range2(pStart, pStart);
			var para = range.Para;
			string line = global::LineSpacing.Of(range);
			var style = StyleEngine.StyleAt(doc, sheet, pStart);
			var one = new ParagraphFormat {
				Align = AlignText(para.Alignment), Left = para.LeftIndent, Right = para.RightIndent, FirstLine = para.FirstLineIndent,
				Before = Math.Max(0, para.SpaceBefore), After = Math.Max(0, para.SpaceAfter), LineSpacing = line == "mixed" ? null : line,
				KeepWithNext = Is(para.KeepWithNext), KeepTogether = Is(para.KeepTogether), WidowControl = StyleEngine.WidowControlOf(para),
				PageBreakBefore = Is(para.PageBreakBefore), NoLineNumber = Is(para.NoLineNumber), Contextual = style.Contextual,
				Tabs = TabsOf(para), Level = StyleEngine.LevelAt(doc, sheet, pStart)
			};
			result = result == null ? one : result.Shared(one);
		}
		return result ?? new ParagraphFormat();
	}

	// What this and another have in common.
	ParagraphFormat Shared(ParagraphFormat o) {
		static T? Same<T>(T? a, T? b) where T : struct => a.HasValue && b.HasValue && a.Value.Equals(b.Value) ? a : null;
		static float? Near(float? a, float? b) => a.HasValue && b.HasValue && Math.Abs(a.Value - b.Value) < 0.05f ? a : null;
		return new ParagraphFormat {
			Align = Align == o.Align ? Align : null, Left = Near(Left, o.Left), Right = Near(Right, o.Right), FirstLine = Near(FirstLine, o.FirstLine),
			Before = Near(Before, o.Before), After = Near(After, o.After), LineSpacing = LineSpacing == o.LineSpacing ? LineSpacing : null,
			KeepWithNext = Same(KeepWithNext, o.KeepWithNext), KeepTogether = Same(KeepTogether, o.KeepTogether), WidowControl = Same(WidowControl, o.WidowControl),
			PageBreakBefore = Same(PageBreakBefore, o.PageBreakBefore), NoLineNumber = Same(NoLineNumber, o.NoLineNumber), Contextual = Same(Contextual, o.Contextual),
			Tabs = Tabs != null && o.Tabs != null && Tabs.SequenceEqual(o.Tabs) ? Tabs : null, Level = Same(Level, o.Level)
		};
	}

	/// Gives the paragraphs from start to end what this sets (the rest of their format stays). Contextual and Level are
	/// a style's, and stay.
	public void Apply(ITextDocument2 doc, int start, int end) {
		(int Rule, float Value)? line = LineSpacing == null ? null : global::LineSpacing.Parse(LineSpacing);
		foreach (var (pStart, _) in StyleEngine.Paragraphs(doc, start, end)) {
			var range = doc.Range2(pStart, pStart);
			var para = range.Para;
			if (Align != null) para.Alignment = AlignOf(Align);
			if (Left != null || Right != null || FirstLine != null)
				para.SetIndents(FirstLine ?? para.FirstLineIndent, Left ?? para.LeftIndent, Right ?? para.RightIndent);
			if (Before is float before) para.SpaceBefore = before;
			if (After is float after) para.SpaceAfter = after;
			if (line is { } spacing) global::LineSpacing.Apply(range, spacing.Rule, spacing.Value);
			if (KeepWithNext is bool keepWith) para.KeepWithNext = Tom(keepWith);
			if (KeepTogether is bool keepTogether) para.KeepTogether = Tom(keepTogether);
			if (WidowControl is bool widow) StyleEngine.SetWidowControl(para, widow);
			if (PageBreakBefore is bool breakBefore) para.PageBreakBefore = Tom(breakBefore);
			if (NoLineNumber is bool noNumber) para.NoLineNumber = Tom(noNumber);
			if (Tabs != null) {
				para.ClearAllTabs();
				foreach (var tab in Tabs.OrderBy(t => t.Position))
					para.AddTab(tab.Position, Array.IndexOf(ParagraphTab.Alignments, tab.Alignment), Math.Max(0, Array.IndexOf(ParagraphTab.Leaders, tab.Leader)));
			}
		}
	}

	/// A style's paragraph format (everything set).
	public static ParagraphFormat Of(ParagraphStyle style) => new() {
		Align = style.Align, Left = style.LeftIndent, Right = style.RightIndent, FirstLine = style.FirstLine, Before = style.Before,
		After = style.After, LineSpacing = style.LineSpacing, KeepWithNext = style.KeepWithNext, KeepTogether = style.KeepTogether,
		WidowControl = style.WidowControl, PageBreakBefore = style.PageBreakBefore, Contextual = style.Contextual, Level = style.Level
	};

	/// Gives a style what this sets (tab stops and suppressed line numbers aren't a style's).
	public void ToStyle(ParagraphStyle style) {
		if (Align != null) style.Align = Align;
		if (Left is float left) style.LeftIndent = left;
		if (Right is float right) style.RightIndent = right;
		if (FirstLine is float first) style.FirstLine = first;
		if (Before is float before) style.Before = before;
		if (After is float after) style.After = after;
		if (LineSpacing != null) style.LineSpacing = LineSpacing;
		if (KeepWithNext is bool keepWith) style.KeepWithNext = keepWith;
		if (KeepTogether is bool keepTogether) style.KeepTogether = keepTogether;
		if (WidowControl is bool widow) style.WidowControl = widow;
		if (PageBreakBefore is bool breakBefore) style.PageBreakBefore = breakBefore;
		if (Contextual is bool contextual) style.Contextual = contextual;
	}

	static readonly string[] Keys = { "Align", "LeftIndent", "RightIndent", "FirstLine", "Before", "After", "LineSpacing", "KeepWithNext",
		"KeepTogether", "WidowControl", "PageBreakBefore", "NoLineNumber", "Tabs", "Contextual", "Level" };

	/// As "Key=Value" lines (what is set): Align, LeftIndent, RightIndent, FirstLine, Before, After (points), LineSpacing,
	/// KeepWithNext, KeepTogether, WidowControl, PageBreakBefore, NoLineNumber (1 or 0), Tabs (72/right/dots,...),
	/// Contextual and Level (the style's).
	public string Describe() {
		var lines = new List<string>();
		void Add(string key, object? value) {
			if (value == null) return;
			lines.Add(key + "=" + value switch {
				bool b => b ? "1" : "0",
				float f => f.ToString("0.##", Invariant),
				List<ParagraphTab> tabs => string.Join(",", tabs),
				_ => value.ToString()
			});
		}
		Add("Align", Align); Add("LeftIndent", Left); Add("RightIndent", Right); Add("FirstLine", FirstLine); Add("Before", Before);
		Add("After", After); Add("LineSpacing", LineSpacing); Add("KeepWithNext", KeepWithNext); Add("KeepTogether", KeepTogether);
		Add("WidowControl", WidowControl); Add("PageBreakBefore", PageBreakBefore); Add("NoLineNumber", NoLineNumber); Add("Tabs", Tabs);
		Add("Contextual", Contextual); Add("Level", Level);
		return string.Join("\n", lines);
	}

	/// "Key=Value" lines (or ; between them) as Describe gives them; measures may have units: 0.5in, 1.27cm, 36pt.
	public static ParagraphFormat Parse(string text) {
		var format = new ParagraphFormat();
		foreach (var raw in text.Split(new[] { '\n', '\r', ';' }, StringSplitOptions.RemoveEmptyEntries)) {
			int eq = raw.IndexOf('=');
			if (eq < 0) throw new ArgumentException($"\"{raw.Trim()}\" is not Key=Value.");
			string key = raw[..eq].Trim(), value = raw[(eq + 1)..].Trim();
			string? known = Keys.FirstOrDefault(k => string.Equals(k, key, StringComparison.OrdinalIgnoreCase))
				?? (string.Equals(key, "Alignment", StringComparison.OrdinalIgnoreCase) ? "Align" : null);
			bool Flag() => value.ToLowerInvariant() switch {
				"1" or "true" or "yes" or "on" => true,
				"0" or "false" or "no" or "off" => false,
				_ => throw new ArgumentException($"{key} is 1 or 0 (not \"{value}\").")
			};
			float Points(bool inches) => TryMeasure(value, inches, out float points) && Math.Abs(points) <= 1584 ? points
				: throw new ArgumentException($"{key}: \"{value}\" is not a measure (72, 1in, 2.5cm, 12pt).");
			switch (known) {
				case "Align":
					string a = value.ToLowerInvariant() switch { "centre" or "centered" or "centred" => "center", "justified" => "justify", var v => v };
					if (a is not ("left" or "center" or "right" or "justify")) throw new ArgumentException($"Align is left, center, right or justify (not \"{value}\").");
					format.Align = a;
					break;
				case "LeftIndent": format.Left = Points(false); break;
				case "RightIndent": format.Right = Points(false); break;
				case "FirstLine": format.FirstLine = Points(false); break;
				case "Before": format.Before = Math.Max(0, Points(false)); break;
				case "After": format.After = Math.Max(0, Points(false)); break;
				case "LineSpacing": global::LineSpacing.Parse(value); format.LineSpacing = value; break;
				case "KeepWithNext": format.KeepWithNext = Flag(); break;
				case "KeepTogether": format.KeepTogether = Flag(); break;
				case "WidowControl": format.WidowControl = Flag(); break;
				case "PageBreakBefore": format.PageBreakBefore = Flag(); break;
				case "NoLineNumber": format.NoLineNumber = Flag(); break;
				case "Contextual": format.Contextual = Flag(); break;
				case "Tabs": format.Tabs = value == "" ? new() : value.Split(',').Select(ParagraphTab.Parse).ToList(); break;
				case "Level": break;    // (the style's)
				default: throw new ArgumentException($"A paragraph has no \"{key}\" ({string.Join(", ", Keys.Take(13))}).");
			}
		}
		return format;
	}

	/// A measure as points: a number with a unit (in, ", cm, mm, pt, pi) or without one (inches if inches, else points).
	public static bool TryMeasure(string text, bool inches, out float points) {
		points = 0;
		var m = Regex.Match(text.Trim().Replace(',', '.').ToLowerInvariant(), @"^(-?\d*\.?\d+)\s*(in|inch|inches|""|″|cm|mm|pt|pts|points?|pi)?$");
		if (!m.Success || !float.TryParse(m.Groups[1].Value, NumberStyles.Float, Invariant, out float value)) return false;
		string unit = m.Groups[2].Value;
		points = unit switch {
			"in" or "inch" or "inches" or "\"" or "″" => value * 72,
			"cm" => value * 72 / 2.54f,
			"mm" => value * 72 / 25.4f,
			"pi" => value * 12,
			"pt" or "pts" or "point" or "points" => value,
			_ => inches ? value * 72 : value
		};
		return true;
	}
}

// Where printed pages end, with the paragraphs' rules. RichEdit's page layout (EM_FORMATRANGE) fits as many lines as
// it can; a rule moves the page's end back to the start of a paragraph or a line, never so far that the page would be
// empty (then the page ends where RichEdit had it, as Word ignores a rule it can't keep). Tables are left as RichEdit
// lays them out. Keep with next moves the whole paragraph with the next (a heading's one line; Word keeps only its last
// line with the next, which differs only for a paragraph of several lines with keep with next).
public static class PageKeeping
{
	const char RowStart = '￹', RowEnd = '￻', CellEnd = '\u0007';

	static bool InTable(string story, int at) {
		if (at < 0 || at >= story.Length) return false;
		if (story[at] is RowStart or RowEnd or CellEnd) return true;
		int open = story.LastIndexOf(RowStart, Math.Max(0, at - 1)), close = at > 0 ? story.LastIndexOf(RowEnd, at - 1) : -1;
		return open >= 0 && open > close;
	}

	// The paragraph a position is in: its start and end (after its mark).
	static (int Start, int End) ParagraphAt(string story, int at) {
		int start = at <= 0 ? 0 : story.LastIndexOfAny(new[] { '\r', CellEnd }, at - 1) + 1;
		int end = story.IndexOfAny(new[] { '\r', CellEnd }, at);
		return (start, end < 0 ? story.Length : end + 1);
	}

	/// A page laid out from start ends before next (where RichEdit ended it): where it ends with the rules. lines: the
	/// printed lines' starts from a paragraph's start up to its end (null: not known, widow control set aside).
	public static int Adjust(ITextDocument2 doc, string story, int start, int next, Func<int, int, IReadOnlyList<int>?> lines) {
		if (next <= start || next >= story.Length - 1) return next;
		if (story[next - 1] == PageBreaks.Break || InTable(story, next) || InTable(story, next - 1)) return next;
		ITextPara Para(int at) => doc.Range2(at, at).Para;
		// page break before: the first paragraph on the page (after its first) that has it ends the page
		for (var (s, e) = ParagraphAt(story, start); e <= next && e < story.Length; (s, e) = ParagraphAt(story, e))
			if (e > start && e < next && !InTable(story, e) && Para(e).PageBreakBefore == tomConstants.tomTrue) return e;
		int end = next;
		var (pStart, pEnd) = ParagraphAt(story, end);
		if (pStart < end) {
			// the page ends inside a paragraph: kept together, it goes to the next page; with widow control, not a single
			// line at the foot of this page or at the top of the next
			var para = Para(pStart);
			if (para.KeepTogether == tomConstants.tomTrue) {
				if (pStart > start) end = pStart;
			} else if (StyleEngine.WidowControlOf(para) && lines(pStart, pEnd) is { Count: > 1 } starts) {
				int here = starts.Count(l => l >= start && l < end), there = starts.Count(l => l >= end);
				if (here < 2) {
					if (pStart > start) end = pStart;
				} else if (there < 2) {
					if (here - 1 >= 2) end = starts.Last(l => l < end);
					else if (pStart > start) end = pStart;
				}
			}
		}
		// keep with next: the paragraphs before the page's end that are to stay with the next one go with it
		if (ParagraphAt(story, end).Start == end) {
			int at = end;
			while (at > start) {
				var (qStart, _) = ParagraphAt(story, at - 1);
				if (InTable(story, qStart) || Para(qStart).KeepWithNext != tomConstants.tomTrue) break;
				if (qStart <= start) { at = end; break; }    // (the whole page: set aside)
				at = qStart;
			}
			end = at;
		}
		return end > start ? end : next;
	}
}

// The document's printed lines: a hidden RichEdit with the document (its RTF), laid out for the printer (its
// information context, the text area's width), whose lines are the printed lines.
public sealed class PrintedLines : IDisposable
{
	[DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	const int EM_SETTARGETDEVICE = 0x0448, EM_LINEINDEX = 0x00BB, EM_EXLINEFROMCHAR = 0x0436, EM_GETLINECOUNT = 0x00BA;
	readonly RichEditBox? hidden;
	readonly Graphics information;
	readonly IntPtr hdc;

	/// information: the printer's information context (kept until Dispose); width: the text area's, twips.
	public PrintedLines(RichEditBox editor, Graphics information, int width) {
		this.information = information;
		hdc = information.GetHdc();
		var box = new RichEditBox { WordWrap = true, Width = 600, Height = 400 };
		IntPtr h = box.Handle;
		((RichTextBox)box).Rtf = ((RichTextBox)editor).Rtf;
		if (box.TextDocument.Range2(0, 0).StoryLength != editor.TextDocument.Range2(0, 0).StoryLength) { box.Dispose(); return; }
		SendMessage(h, EM_SETTARGETDEVICE, hdc, (IntPtr)width);
		hidden = box;
	}

	/// The starts of the printed lines from start up to end (null: not known).
	public IReadOnlyList<int>? Starts(int start, int end) {
		if (hidden == null) return null;
		IntPtr h = hidden.Handle;
		var starts = new List<int>();
		int count = (int)SendMessage(h, EM_GETLINECOUNT, IntPtr.Zero, IntPtr.Zero);
		for (int line = (int)SendMessage(h, EM_EXLINEFROMCHAR, IntPtr.Zero, (IntPtr)start); line < count; line++) {
			int at = (int)SendMessage(h, EM_LINEINDEX, (IntPtr)line, IntPtr.Zero);
			if (at < 0 || at >= end) break;
			starts.Add(Math.Max(at, start));
		}
		return starts;
	}

	public void Dispose() {
		hidden?.Dispose();
		information.ReleaseHdc(hdc);
		information.Dispose();
	}
}

// A measure in a spin box, shown with its unit (0.5" or 12 pt) and typed with any (1.27 cm, 36pt, 0.5in); empty when
// the paragraphs differ (then nothing is changed unless a value is given).
sealed class MeasureBox : NumericUpDown
{
	readonly bool inches;
	bool blank;

	/// Raised when the user gives a value (spinning, typing).
	public event EventHandler? Edited;

	public MeasureBox(bool inches, float minimumPoints, float maximumPoints) {
		this.inches = inches;
		DecimalPlaces = inches ? 2 : 1;
		Increment = inches ? 0.1m : 6m;
		Minimum = (decimal)(inches ? minimumPoints / 72 : minimumPoints);
		Maximum = (decimal)(inches ? maximumPoints / 72 : maximumPoints);
		Width = 80;
		TextAlign = HorizontalAlignment.Left;
	}

	/// The value in points (null: empty).
	public float? Points {
		get => blank ? null : (float)(inches ? Value * 72 : Value);
		set {
			blank = value == null;
			if (value is float points) {
				decimal v = (decimal)Math.Round(inches ? points / 72 : points, inches ? 3 : 2);
				Value = Math.Clamp(v, Minimum, Maximum);
			}
			UpdateEditText();
		}
	}

	protected override void OnValueChanged(EventArgs e) {
		base.OnValueChanged(e);
		if (!Created || !Focused && !UserEdit) return;
		blank = false;
		Edited?.Invoke(this, e);
	}

	public override void UpButton() { bool was = blank; blank = false; base.UpButton(); if (was) { UpdateEditText(); Edited?.Invoke(this, EventArgs.Empty); } }
	public override void DownButton() { bool was = blank; blank = false; base.DownButton(); if (was) { UpdateEditText(); Edited?.Invoke(this, EventArgs.Empty); } }

	protected override void UpdateEditText() {
		if (blank) { Text = ""; return; }
		Text = inches ? Value.ToString("0.##", CultureInfo.CurrentCulture) + "\"" : Value.ToString("0.#", CultureInfo.CurrentCulture) + " pt";
	}

	protected override void ValidateEditText() {
		string text = Text.Trim();
		UserEdit = false;    // (first: reading Value validates a pending edit, which would come back here)
		if (text == "") { if (!blank) UpdateEditText(); return; }
		if (ParagraphFormat.TryMeasure(text.Replace(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, "."), inches, out float points)) {
			bool was = blank;
			blank = false;
			decimal before = Value;
			Points = points;
			if (was || before != Value) Edited?.Invoke(this, EventArgs.Empty);
		}
		UpdateEditText();
	}
}

// Home > Paragraph (the group's corner button), as Word's Paragraph dialog. For the selected paragraphs (a value they
// don't share is left empty, and stays as each has it unless one is given), or (forStyle) a style's paragraph format.
sealed class ParagraphDialog : Form
{
	readonly ComboBox align = Combo("Left", "Centered", "Right", "Justified");
	readonly ComboBox level = Combo("Body Text", "Level 1", "Level 2", "Level 3", "Level 4", "Level 5", "Level 6");
	readonly MeasureBox left = new(true, -1584, 1584), right = new(true, -1584, 1584), by = new(true, 0, 1584);
	readonly ComboBox special = Combo("(none)", "First line", "Hanging");
	readonly MeasureBox before = new(false, 0, 1584), after = new(false, 0, 1584);
	readonly ComboBox line = Combo("Single", "1.5 lines", "Double", "At least", "Exactly", "Multiple");
	readonly NumericUpDown at = new() { DecimalPlaces = 2, Increment = 0.25m, Minimum = 0.25m, Maximum = 1584, Width = 70 };
	readonly CheckBox contextual = Check("Don't add spa&ce between paragraphs of the same style");
	readonly CheckBox widow = Check("&Widow/Orphan control"), keepNext = Check("Keep with ne&xt"), keepLines = Check("&Keep lines together"),
		pageBreak = Check("Page &break before"), noNumber = Check("&Suppress line numbers");
	readonly Panel preview = new() { Height = 120, Dock = DockStyle.Fill, BackColor = SystemColors.Window, BorderStyle = BorderStyle.FixedSingle };
	readonly Label error = new() { AutoSize = true, ForeColor = Color.Firebrick };
	readonly ParagraphFormat original;
	readonly bool forStyle;
	readonly HashSet<string> changed = new();
	readonly string fontName;
	readonly float fontSize;
	List<ParagraphTab>? tabs;
	bool tabsChanged, loading = true;

	/// The default tab stops' spacing (points), as the Tabs dialog left it (null: not changed).
	public float? DefaultTabStop { get; private set; }
	float defaultTab;

	/// Set As Default was chosen: the paragraph's style is to get the format (and, NewDocumentsToo, new documents).
	public bool SetAsDefault { get; private set; }
	public bool NewDocumentsToo { get; private set; }

	static ComboBox Combo(params string[] items) { var c = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130 }; c.Items.AddRange(items); return c; }
	static CheckBox Check(string text) => new() { Text = text, AutoSize = true, ThreeState = false, Margin = new Padding(3, 3, 3, 3) };
	static readonly string[] Aligns = { "left", "center", "right", "justify" };

	/// format: the paragraphs' (or style's) format now; styleName: their style (shown, and Set As Default's); font:
	/// their text's, for the preview; defaultTab: the document's default tab stops (points).
	public ParagraphDialog(ParagraphFormat format, bool forStyle, string styleName, string fontName = "Times New Roman", float fontSize = 12, float defaultTab = 36) {
		original = format.Copy();
		this.forStyle = forStyle;
		this.fontName = fontName;
		this.fontSize = fontSize;
		this.defaultTab = defaultTab;
		tabs = format.Tabs?.ToList();
		Text = forStyle ? "Paragraph: " + styleName : "Paragraph";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MaximizeBox = MinimizeBox = false;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoScaleMode = AutoScaleMode.Dpi;
		Font = SystemFonts.MessageBoxFont ?? Font;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;

		// the values
		align.SelectedIndex = format.Align == null ? -1 : Math.Max(0, Array.IndexOf(Aligns, format.Align));
		level.SelectedIndex = format.Level is int lv ? Math.Clamp(lv, 0, 6) : -1;
		level.Enabled = false;
		left.Points = format.Left;
		right.Points = format.Right;
		if (format.FirstLine is float first) {
			special.SelectedIndex = Math.Abs(first) < 0.05f ? 0 : first > 0 ? 1 : 2;
			by.Points = Math.Abs(first) < 0.05f ? null : Math.Abs(first);
		} else {
			special.SelectedIndex = -1;
			by.Points = null;
		}
		before.Points = format.Before;
		after.Points = format.After;
		SetLine(format.LineSpacing);
		SetCheck(contextual, format.Contextual);
		SetCheck(widow, format.WidowControl);
		SetCheck(keepNext, format.KeepWithNext);
		SetCheck(keepLines, format.KeepTogether);
		SetCheck(pageBreak, format.PageBreakBefore);
		SetCheck(noNumber, format.NoLineNumber);
		contextual.Enabled = forStyle;
		var tips = new ToolTip();
		if (!forStyle) tips.SetToolTip(contextual, $"A setting of the paragraph's style ({styleName}): Home > Styles > Modify Style > Paragraph.");
		tips.SetToolTip(level, "The style's heading level (Home > Styles > Modify Style): headings are levels 1 to 6.");

		// what was changed (only that is given to the paragraphs)
		void Mark(string what) { if (loading) return; changed.Add(what); preview.Invalidate(); }
		align.SelectedIndexChanged += (s, e) => Mark("align");
		left.Edited += (s, e) => Mark("left");
		right.Edited += (s, e) => Mark("right");
		special.SelectedIndexChanged += (s, e) => {
			if (loading) return;
			if (special.SelectedIndex == 0) by.Points = null;
			else if (by.Points is not > 0) by.Points = 36;
			by.Enabled = special.SelectedIndex > 0;
			Mark("first");
		};
		by.Edited += (s, e) => { if (special.SelectedIndex <= 0) special.SelectedIndex = 1; Mark("first"); };
		before.Edited += (s, e) => Mark("before");
		after.Edited += (s, e) => Mark("after");
		line.SelectedIndexChanged += (s, e) => { if (loading) return; FitAt(changedRule: true); Mark("line"); };
		at.ValueChanged += (s, e) => { if (!loading && line.SelectedIndex >= 0) Mark("line"); };
		foreach (var (box, what) in new[] { (contextual, "contextual"), (widow, "widow"), (keepNext, "keepNext"), (keepLines, "keepLines"), (pageBreak, "pageBreak"), (noNumber, "noNumber") })
			box.CheckStateChanged += (s, e) => { if (loading) return; box.ThreeState = false; Mark(what); };
		preview.Paint += PaintPreview;

		// Indents and Spacing
		var spacingPage = new TabPage("Indents and Spacing") { Padding = new Padding(8), UseVisualStyleBackColor = true };
		var grid = Grid();
		Heading(grid, "General");
		Row(grid, "&Alignment:", align, "&Outline level:", level);
		Heading(grid, "Indentation");
		Row(grid, "&Left:", left, "S&pecial:", special);
		Row(grid, "&Right:", right, "B&y:", by);
		Heading(grid, "Spacing");
		Row(grid, "&Before:", before, "L&ine spacing:", line);
		Row(grid, "Aft&er:", after, "&At:", at);
		int cr = grid.RowCount++;
		grid.Controls.Add(contextual, 0, cr);
		grid.SetColumnSpan(contextual, 4);
		spacingPage.Controls.Add(grid);

		// Line and Page Breaks
		var breaksPage = new TabPage("Line and Page Breaks") { Padding = new Padding(8), UseVisualStyleBackColor = true };
		var grid2 = Grid();
		Heading(grid2, "Pagination");
		foreach (var box in new[] { widow, keepNext, keepLines, pageBreak }) { int r = grid2.RowCount++; grid2.Controls.Add(box, 0, r); grid2.SetColumnSpan(box, 4); }
		if (!forStyle) {
			Heading(grid2, "Formatting exceptions");
			int r = grid2.RowCount++;
			grid2.Controls.Add(noNumber, 0, r);
			grid2.SetColumnSpan(noNumber, 4);
		}
		var note = new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(400), 0), ForeColor = SystemColors.GrayText, Margin = new Padding(3, 12, 3, 3),
			Text = "Printing, Print Preview, Export as PDF and Multiple Pages keep these; Word and LibreOffice get them in their files."
		};
		int nr = grid2.RowCount++;
		grid2.Controls.Add(note, 0, nr);
		grid2.SetColumnSpan(note, 4);
		breaksPage.Controls.Add(grid2);

		var tabsControl = new TabControl();
		tabsControl.TabPages.Add(spacingPage);
		tabsControl.TabPages.Add(breaksPage);
		// (a tab control doesn't size itself to its pages)
		var need = new Size(Math.Max(grid.PreferredSize.Width, grid2.PreferredSize.Width), Math.Max(grid.PreferredSize.Height, grid2.PreferredSize.Height));
		tabsControl.Size = new Size(need.Width + 30, need.Height + 50);

		var layout = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Padding = new Padding(10), Dock = DockStyle.Fill };
		layout.Controls.Add(tabsControl);
		layout.Controls.Add(new Label { Text = "Preview", AutoSize = true, Margin = new Padding(3, 8, 3, 2) });
		layout.Controls.Add(preview);
		layout.Controls.Add(error);
		var buttons = new TableLayoutPanel { ColumnCount = 5, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 6, 0, 0) };
		buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
		buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
		buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
		buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
		buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
		var tabsButton = new Button { Text = "&Tabs...", AutoSize = true };
		var defaultButton = new Button { Text = "Set As &Default...", AutoSize = true };
		var ok = new Button { Text = "OK", AutoSize = true, MinimumSize = new Size(LogicalToDeviceUnits(80), 0) };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(LogicalToDeviceUnits(80), 0) };
		tabsButton.Click += (s, e) => TabsCommand();
		defaultButton.Click += (s, e) => DefaultCommand(styleName);
		ok.Click += (s, e) => { if (Validate(out string message)) { DialogResult = DialogResult.OK; Close(); } else error.Text = message; };
		if (!forStyle) { buttons.Controls.Add(tabsButton, 0, 0); buttons.Controls.Add(defaultButton, 1, 0); }
		buttons.Controls.Add(ok, 3, 0);
		buttons.Controls.Add(cancel, 4, 0);
		layout.Controls.Add(buttons);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		by.Enabled = special.SelectedIndex != 0;
		loading = false;
	}

	static TableLayoutPanel Grid() {
		var grid = new TableLayoutPanel { ColumnCount = 4, AutoSize = true, Dock = DockStyle.Fill };
		for (int i = 0; i < 4; i++) grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
		return grid;
	}

	void Heading(TableLayoutPanel grid, string text) {
		int r = grid.RowCount++;
		var label = new Label { Text = text, AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(0, r == 0 ? 2 : 10, 3, 2) };
		grid.Controls.Add(label, 0, r);
		grid.SetColumnSpan(label, 4);
	}

	static void Row(TableLayoutPanel grid, string label, Control a, string label2, Control b) {
		int r = grid.RowCount++;
		grid.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 6, 6, 0) }, 0, r);
		grid.Controls.Add(a, 1, r);
		grid.Controls.Add(new Label { Text = label2, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(24, 6, 6, 0) }, 2, r);
		grid.Controls.Add(b, 3, r);
	}

	static void SetCheck(CheckBox box, bool? value) {
		box.ThreeState = value == null;
		box.CheckState = value == null ? CheckState.Indeterminate : value.Value ? CheckState.Checked : CheckState.Unchecked;
	}

	void SetLine(string? spacing) {
		if (spacing == null) { line.SelectedIndex = -1; at.Enabled = false; return; }
		var (rule, value) = global::LineSpacing.Parse(spacing);
		if (rule == global::LineSpacing.Multiple && Math.Abs(value - 1.5f) < 0.001) rule = global::LineSpacing.OneAndHalf;
		else if (rule == global::LineSpacing.Multiple && Math.Abs(value - 2) < 0.001) rule = global::LineSpacing.Double;
		line.SelectedIndex = rule;
		at.Value = (decimal)Math.Clamp(rule is global::LineSpacing.AtLeast or global::LineSpacing.Exactly or global::LineSpacing.Multiple ? value : 1, 0.25f, 1584f);
		FitAt(changedRule: false);
	}

	// The At box: lines for Multiple, points for At least and Exactly, nothing for the fixed rules.
	void FitAt(bool changedRule) {
		int r = line.SelectedIndex;
		bool points = r is global::LineSpacing.AtLeast or global::LineSpacing.Exactly;
		at.Enabled = points || r == global::LineSpacing.Multiple;
		at.Maximum = points ? 1584 : 132;
		at.Increment = points ? 1 : 0.25m;
		if (changedRule && points && at.Value < 6) at.Value = 12;
		if (changedRule && r == global::LineSpacing.Multiple && at.Value > 10) at.Value = 1.15m;
	}

	string? LineResult => line.SelectedIndex switch {
		-1 => null,
		global::LineSpacing.Single => "1",
		global::LineSpacing.OneAndHalf => "1.5",
		global::LineSpacing.Double => "2",
		global::LineSpacing.AtLeast => "at least " + at.Value.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
		global::LineSpacing.Exactly => "exactly " + at.Value.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
		_ => at.Value.ToString("0.##", CultureInfo.InvariantCulture)
	};

	float? FirstResult => special.SelectedIndex switch { -1 => null, 0 => 0, 1 => by.Points ?? 0, _ => -(by.Points ?? 0) };

	static bool? Checked(CheckBox box) => box.CheckState == CheckState.Indeterminate ? null : box.Checked;

	bool Validate(out string message) {
		message = "";
		if (special.SelectedIndex > 0 && by.Points is not > 0) { message = "Give the first line's indent (By)."; return false; }
		return true;
	}

	/// What the dialog sets: for paragraphs, only what was changed; for a style, everything.
	public ParagraphFormat Result {
		get {
			bool all = forStyle;
			bool Want(string what) => all || changed.Contains(what);
			return new ParagraphFormat {
				Align = Want("align") && align.SelectedIndex >= 0 ? Aligns[align.SelectedIndex] : all ? original.Align : null,
				Left = Want("left") ? left.Points ?? original.Left : null,
				Right = Want("right") ? right.Points ?? original.Right : null,
				FirstLine = Want("first") ? FirstResult ?? original.FirstLine : null,
				Before = Want("before") ? before.Points ?? original.Before : null,
				After = Want("after") ? after.Points ?? original.After : null,
				LineSpacing = Want("line") ? LineResult ?? original.LineSpacing : null,
				KeepWithNext = Want("keepNext") ? Checked(keepNext) : null,
				KeepTogether = Want("keepLines") ? Checked(keepLines) : null,
				WidowControl = Want("widow") ? Checked(widow) : null,
				PageBreakBefore = Want("pageBreak") ? Checked(pageBreak) : null,
				NoLineNumber = !forStyle && Want("noNumber") ? Checked(noNumber) : null,
				Contextual = forStyle ? Checked(contextual) : null,
				Tabs = !forStyle && tabsChanged ? tabs : null
			};
		}
	}

	void TabsCommand() {
		using var dlg = new TabsDialog(tabs, defaultTab);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		if (dlg.TabsChanged) { tabs = dlg.Tabs; tabsChanged = true; }
		if (dlg.DefaultTabStop is float spacing && Math.Abs(spacing - defaultTab) > 0.01f) { DefaultTabStop = defaultTab = spacing; }
	}

	// Set As Default: the paragraph's style gets this format (as Word's sets the Normal template's), for this document
	// or for new documents too.
	void DefaultCommand(string styleName) {
		using var ask = new Form {
			Text = "Set As Default", FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, ShowInTaskbar = false,
			StartPosition = FormStartPosition.CenterParent, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Font = Font
		};
		var layout = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Padding = new Padding(12), Dock = DockStyle.Fill };
		layout.Controls.Add(new Label { AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(380), 0), Text = $"Make these the paragraph settings of the style {styleName} (every paragraph with it changes, as when the style is modified):" });
		var thisOnly = new RadioButton { Text = "&This document only", AutoSize = true, Checked = true, Margin = new Padding(12, 8, 3, 3) };
		var all = new RadioButton { Text = "This document and &new documents (Save Styles as My Default)", AutoSize = true, Margin = new Padding(12, 3, 3, 3) };
		layout.Controls.Add(thisOnly);
		layout.Controls.Add(all);
		var row = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		row.Controls.Add(cancel);
		row.Controls.Add(ok);
		layout.Controls.Add(row);
		ask.Controls.Add(layout);
		ask.AcceptButton = ok;
		ask.CancelButton = cancel;
		if (ask.ShowDialog(this) != DialogResult.OK) return;
		if (!Validate(out string message)) { error.Text = message; return; }
		SetAsDefault = true;
		NewDocumentsToo = all.Checked;
		DialogResult = DialogResult.OK;
		Close();
	}

	// The preview: grey paragraphs before and after, and the paragraph in its font, with its alignment, indents,
	// spacing and line spacing to scale (a 6.5 inch text width across the box).
	void PaintPreview(object? sender, PaintEventArgs e) {
		var g = e.Graphics;
		g.Clear(SystemColors.Window);
		var area = preview.ClientRectangle;
		area.Inflate(-LogicalToDeviceUnits(10), -LogicalToDeviceUnits(6));
		float scale = area.Width / 468f;    // pixels per point
		using var grey = new SolidBrush(Color.FromArgb(0xC0, 0xC0, 0xC0));
		using var greyFont = new Font("Segoe UI", Math.Max(5f, 7f * scale * 0.9f), GraphicsUnit.Pixel);
		float pointSize = Math.Max(5f, fontSize * scale * 0.75f);
		Font sample;
		try { sample = new Font(fontName, pointSize, GraphicsUnit.Pixel); } catch (ArgumentException) { sample = new Font("Segoe UI", pointSize, GraphicsUnit.Pixel); }
		using (sample) {
			float y = area.Top;
			y = Paragraph(g, "Previous Paragraph ", 3, greyFont, grey, area, y, 0, 0, 0, 0, 1, scale) + 2;
			var result = Result;
			ParagraphFormat Effective() {
				var f = original.Copy();
				var r = result;
				f.Align = r.Align ?? f.Align; f.Left = r.Left ?? f.Left; f.Right = r.Right ?? f.Right; f.FirstLine = r.FirstLine ?? f.FirstLine;
				f.Before = r.Before ?? f.Before; f.After = r.After ?? f.After; f.LineSpacing = r.LineSpacing ?? f.LineSpacing;
				return f;
			}
			var shown = Effective();
			float lineFactor = 1;
			float? exact = null;
			try {
				var (rule, value) = global::LineSpacing.Parse(shown.LineSpacing ?? "1");
				lineFactor = rule switch { global::LineSpacing.OneAndHalf => 1.5f, global::LineSpacing.Double => 2, global::LineSpacing.Multiple => value, _ => 1 };
				if (rule == global::LineSpacing.Exactly) exact = value;
				else if (rule == global::LineSpacing.AtLeast) exact = Math.Max(value, fontSize * 1.15f);
			} catch (ArgumentException) { }
			y += (shown.Before ?? 0) * scale;
			using var ink = new SolidBrush(SystemColors.WindowText);
			y = Paragraph(g, "Sample Text ", 4, sample, ink, area, y, shown.Left ?? 0, shown.Right ?? 0, shown.FirstLine ?? 0,
				Array.IndexOf(Aligns, shown.Align ?? "left"), lineFactor, scale, exact);
			y += (shown.After ?? 0) * scale + 2;
			Paragraph(g, "Following Paragraph ", 4, greyFont, grey, area, y, 0, 0, 0, 0, 1, scale);
		}
	}

	// Draws a paragraph of a phrase repeated (lines' worth), wrapped in the area with the indents (points), the
	// alignment and the line spacing; returns the y below it.
	static float Paragraph(Graphics g, string phrase, int lines, Font font, Brush brush, Rectangle area, float y, float left, float right,
		float first, int alignment, float lineFactor, float scale, float? exactPoints = null) {
		float lineHeight = exactPoints is float pts ? pts * scale : font.GetHeight(g) * lineFactor;
		var words = string.Concat(Enumerable.Repeat(phrase, lines * 4)).Split(' ', StringSplitOptions.RemoveEmptyEntries);
		float space = font.GetHeight(g) * 0.3f;    // (a space's width; MeasureString leaves a lone space out)
		int index = 0;
		for (int n = 0; n < lines && index < words.Length && y < area.Bottom; n++) {
			float x0 = area.Left + (left + (n == 0 ? first : 0)) * scale, x1 = area.Right - right * scale;
			if (x1 - x0 < 20) x1 = x0 + 20;
			var row = new List<(string Word, float Width)>();
			float width = 0;
			while (index < words.Length) {
				float w = g.MeasureString(words[index], font, PointF.Empty, StringFormat.GenericTypographic).Width;
				if (row.Count > 0 && width + space + w > x1 - x0) break;
				width += (row.Count > 0 ? space : 0) + w;
				row.Add((words[index++], w));
			}
			bool last = n == lines - 1;
			float x = alignment switch { 1 => x0 + (x1 - x0 - width) / 2, 2 => x1 - width, _ => x0 };
			float gap = alignment == 3 && !last && row.Count > 1 ? space + (x1 - x0 - width) / (row.Count - 1) : space;
			foreach (var (word, w) in row) {
				g.DrawString(word, font, brush, x, y, StringFormat.GenericTypographic);
				x += w + gap;
			}
			y += lineHeight;
		}
		return y;
	}
}

// The Paragraph dialog's Tabs: the selected paragraphs' tab stops (a position, its alignment and leader) and the
// document's default tab stops, as Word's Tabs dialog.
sealed class TabsDialog : Form
{
	readonly MeasureBox position = new(true, 0, 1584), defaultStops = new(true, 1, 1584);
	readonly ListBox list = new() { Width = 120, Height = 110, IntegralHeight = false };
	readonly RadioButton[] alignments = ParagraphTab.Alignments.Select(a => new RadioButton { Text = Title(a), AutoSize = true, Margin = new Padding(3, 1, 3, 1) }).ToArray();
	readonly RadioButton[] leaders = new[] { "none", "dots", "dashes", "underline" }.Select((l, i) => new RadioButton { Text = (i + 1) + " " + l switch { "none" => "None", "dots" => ".......", "dashes" => "-------", _ => "_______" }, Tag = l, AutoSize = true, Margin = new Padding(3, 1, 3, 1) }).ToArray();
	readonly Label error = new() { AutoSize = true, ForeColor = Color.Firebrick };
	List<ParagraphTab> stops;

	public List<ParagraphTab> Tabs => stops.OrderBy(t => t.Position).ToList();
	public bool TabsChanged { get; private set; }
	public float? DefaultTabStop => defaultStops.Points;

	static string Title(string s) => char.ToUpperInvariant(s[0]) + s[1..];
	static string Show(float points) => (points / 72).ToString("0.##", CultureInfo.CurrentCulture) + "\"";

	/// tabs: the paragraphs' tab stops (null: they differ: none shown, and setting any replaces theirs); defaultTab: points.
	public TabsDialog(List<ParagraphTab>? tabs, float defaultTab) {
		stops = tabs?.ToList() ?? new List<ParagraphTab>();
		Text = "Tabs";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MaximizeBox = MinimizeBox = false;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoScaleMode = AutoScaleMode.Dpi;
		Font = SystemFonts.MessageBoxFont ?? Font;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		defaultStops.Points = defaultTab;
		alignments[0].Checked = true;
		leaders[0].Checked = true;
		Fill();
		list.SelectedIndexChanged += (s, e) => {
			if (list.SelectedIndex < 0 || list.SelectedIndex >= stops.Count) return;
			var tab = Tabs[list.SelectedIndex];
			position.Points = tab.Position;
			alignments[Math.Max(0, Array.IndexOf(ParagraphTab.Alignments, tab.Alignment))].Checked = true;
			var leader = leaders.FirstOrDefault(r => (string)r.Tag! == tab.Leader) ?? leaders[0];
			leader.Checked = true;
		};

		var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Padding = new Padding(10), Dock = DockStyle.Fill };
		var stopsColumn = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
		stopsColumn.Controls.Add(new Label { Text = "&Tab stop position:", AutoSize = true });
		stopsColumn.Controls.Add(position);
		stopsColumn.Controls.Add(list);
		var defaultColumn = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Margin = new Padding(16, 3, 3, 3) };
		defaultColumn.Controls.Add(new Label { Text = "De&fault tab stops:", AutoSize = true });
		defaultColumn.Controls.Add(defaultStops);
		defaultColumn.Controls.Add(new Label { Text = "(the whole document's)", AutoSize = true, ForeColor = SystemColors.GrayText });
		grid.Controls.Add(stopsColumn, 0, 0);
		grid.Controls.Add(defaultColumn, 1, 0);
		var alignBox = new GroupBox { Text = "Alignment", AutoSize = true, Dock = DockStyle.Fill };
		var alignFlow = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
		alignFlow.Controls.AddRange(alignments);
		alignBox.Controls.Add(alignFlow);
		var leaderBox = new GroupBox { Text = "Leader", AutoSize = true, Dock = DockStyle.Fill };
		var leaderFlow = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
		leaderFlow.Controls.AddRange(leaders);
		leaderBox.Controls.Add(leaderFlow);
		grid.Controls.Add(alignBox, 0, 1);
		grid.Controls.Add(leaderBox, 1, 1);
		var setButton = new Button { Text = "&Set", AutoSize = true };
		var clearButton = new Button { Text = "Cl&ear", AutoSize = true };
		var clearAllButton = new Button { Text = "Clear &All", AutoSize = true };
		setButton.Click += (s, e) => SetStop();
		clearButton.Click += (s, e) => ClearStop();
		clearAllButton.Click += (s, e) => { stops.Clear(); TabsChanged = true; Fill(); };
		var actions = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 8, 0, 0) };
		actions.Controls.AddRange(new Control[] { setButton, clearButton, clearAllButton });
		grid.Controls.Add(actions, 0, 2);
		grid.SetColumnSpan(actions, 2);
		grid.Controls.Add(error, 0, 3);
		grid.SetColumnSpan(error, 2);
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var ok = new Button { Text = "OK", AutoSize = true };
		ok.Click += (s, e) => {
			// (a position typed and not Set: set, as Word does on OK)
			if (position.Points is float p && !stops.Any(t => Math.Abs(t.Position - p) < 0.05f) && position.Text.Trim() != "") SetStop();
			DialogResult = DialogResult.OK;
			Close();
		};
		buttons.Controls.Add(cancel);
		buttons.Controls.Add(ok);
		grid.Controls.Add(buttons, 0, 4);
		grid.SetColumnSpan(buttons, 2);
		Controls.Add(grid);
		AcceptButton = ok;
		CancelButton = cancel;
		position.Points = null;
	}

	void Fill() {
		list.Items.Clear();
		foreach (var tab in Tabs) list.Items.Add(Show(tab.Position) + (tab.Alignment == "left" ? "" : " " + tab.Alignment) + (tab.Leader == "none" ? "" : " " + tab.Leader));
	}

	void SetStop() {
		if (position.Points is not float p) { error.Text = "Type the tab stop's position."; return; }
		error.Text = "";
		string alignment = ParagraphTab.Alignments[Math.Max(0, Array.FindIndex(alignments, r => r.Checked))];
		string leader = (string)(leaders.FirstOrDefault(r => r.Checked)?.Tag ?? "none");
		stops.RemoveAll(t => Math.Abs(t.Position - p) < 0.05f);
		stops.Add(new ParagraphTab(p, alignment, leader));
		TabsChanged = true;
		Fill();
		position.Points = null;
	}

	void ClearStop() {
		float? p = list.SelectedIndex >= 0 && list.SelectedIndex < stops.Count ? Tabs[list.SelectedIndex].Position : position.Points;
		if (p is not float at) return;
		if (stops.RemoveAll(t => Math.Abs(t.Position - at) < 0.05f) > 0) TabsChanged = true;
		Fill();
		position.Points = null;
	}
}
