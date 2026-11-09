// Paragraph borders and shading, for AhkWord (AhkWord.cs, through #load). Docxodus trial.
//
// RichEdit keeps a paragraph's border words in its RTF but draws none of it, and keeps no shading colour, so a
// paragraph's borders and shading are AhkWord's own: a hidden marker at the paragraph's start, [[para:{...}]] (hidden
// text, as the cross-references' and sections' markers are), with each side's line (style, width in eighths of a point,
// colour, distance from the text in points), the line between paragraphs boxed together, and the fill colour.
//   - Read from a .docx (whichever converter read its text): each paragraph's own w:pBdr and w:shd over its style's,
//     paired up with the editor's paragraphs in order (DocxLineSpacing).
//   - Drawn by AhkWord: on the screen (WriterEditBox: the shading under the text, by darkening what RichEdit painted, and
//     the lines), and on printed pages, Print Preview and PDF (DocumentPrinter: the shading before the text).
//     Paragraphs one after the other with the same borders and shading are one box, as in Word.
//   - Written into RTF as Word's paragraph words (\brdrt ... \cbpat, \shading), which the direct .docx writer turns into
//     w:pBdr and w:shd; the marker stays in the text (hidden), as AhkWord's other markers do.
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using EllipticBit.RichEditorNET.TextObjectModel2;

/// A border line of a paragraph: Word's style name (single, double, dotted, dashed, thick ...), its width in eighths of a
/// point, its colour (RRGGBB) and its distance from the text (points).
public sealed class FrameLine
{
	[JsonPropertyName("s")] public string Style { get; set; } = "single";
	[JsonPropertyName("w")] public int Size { get; set; } = 4;
	[JsonPropertyName("c")] public string Color { get; set; } = "000000";
	[JsonPropertyName("p")] public int Space { get; set; }
}

/// A paragraph's borders and shading (null: none on that side, no fill).
public sealed class ParagraphFrame
{
	[JsonPropertyName("t")] public FrameLine? Top { get; set; }
	[JsonPropertyName("l")] public FrameLine? Left { get; set; }
	[JsonPropertyName("b")] public FrameLine? Bottom { get; set; }
	[JsonPropertyName("r")] public FrameLine? Right { get; set; }
	/// The line between paragraphs boxed together (Word's w:between).
	[JsonPropertyName("h")] public FrameLine? Between { get; set; }
	/// The fill colour (RRGGBB).
	[JsonPropertyName("f")] public string? Fill { get; set; }

	[JsonIgnore] public bool Any => Top != null || Left != null || Bottom != null || Right != null || Between != null || Fill != null;

	public string ToJson() => JsonSerializer.Serialize(this, ParagraphFrames.Json);
	public bool SameAs(ParagraphFrame? other) => other != null && other.ToJson() == ToJson();
}

public static class ParagraphFrames
{
	internal static readonly JsonSerializerOptions Json = new() { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull };
	static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

	/// The marker: [[para:{...}]] at a paragraph's start.
	public static readonly Regex Mark = new(@"\[\[para:(\{(?:[^\]\r]|\](?!\]))*\})\]\]");
	public static string MarkerFor(ParagraphFrame frame) => "[[para:" + frame.ToJson() + "]]";
	public static ParagraphFrame? Parse(string json) {
		try { return JsonSerializer.Deserialize<ParagraphFrame>(json, Json); } catch (JsonException) { return null; }
	}

	public static bool Any(string story) => story.Contains("[[para:", StringComparison.Ordinal);

	/// A framed paragraph: where it starts and ends (its paragraph mark), and its frame.
	public sealed record Framed(int Start, int End, ParagraphFrame Frame);

	/// The framed paragraphs of a story, in order.
	public static List<Framed> All(string story) {
		var list = new List<Framed>();
		if (!Any(story)) return list;
		foreach (Match m in Mark.Matches(story)) {
			if (Parse(m.Groups[1].Value) is not { Any: true } frame) continue;
			int start = m.Index == 0 ? 0 : story.LastIndexOfAny(new[] { '\r', '\u0007' }, m.Index - 1) + 1;
			int end = story.IndexOfAny(new[] { '\r', '\u0007' }, m.Index);
			list.Add(new Framed(start, end < 0 ? story.Length : end, frame));
		}
		return list;
	}

	/// Whether a framed paragraph and the next one are one box: the same frame, and the next paragraph straight after.
	public static bool Joined(Framed a, Framed b) => b.Start == a.End + 1 && a.Frame.SameAs(b.Frame);

	// ================================================================ .docx

	/// A border line from a w:pBdr child (null: none).
	public static FrameLine? LineOf(XElement? element) {
		if (element == null) return null;
		string style = (string?)element.Attribute(W + "val") ?? "single";
		if (style is "nil" or "none") return null;
		int size = int.TryParse((string?)element.Attribute(W + "sz"), out int sz) ? sz : 4;
		int space = int.TryParse((string?)element.Attribute(W + "space"), out int sp) ? sp : 0;
		string color = (string?)element.Attribute(W + "color") is { Length: 6 } c && c != "auto" ? c.ToUpperInvariant() : "000000";
		return new FrameLine { Style = style, Size = Math.Clamp(size, 2, 96), Color = color, Space = Math.Clamp(space, 0, 31) };
	}

	/// A w:shd's colour as it shows (RRGGBB; null: none): its fill, or its pattern's colour over the fill (solid, pctN).
	public static string? FillOf(XElement? shd) {
		if (shd == null) return null;
		string val = (string?)shd.Attribute(W + "val") ?? "clear";
		if (val is "nil") return null;
		static Color? Parse(string? hex) => hex is { Length: 6 } && hex != "auto" && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int v) ? System.Drawing.Color.FromArgb(v >> 16 & 255, v >> 8 & 255, v & 255) : null;
		var fill = Parse((string?)shd.Attribute(W + "fill"));
		var pattern = Parse((string?)shd.Attribute(W + "color")) ?? System.Drawing.Color.Black;
		double share = val switch {
			"solid" => 1,
			"clear" => 0,
			_ when val.StartsWith("pct", StringComparison.Ordinal) && int.TryParse(val[3..], out int pct) => pct / 100.0,
			_ => 0.25    // (stripes and crosses: about a quarter of the pattern's colour)
		};
		if (fill == null && share == 0) return null;
		var under = fill ?? System.Drawing.Color.White;
		int Mix(int a, int b) => (int)Math.Round(a * (1 - share) + b * share);
		var shown = System.Drawing.Color.FromArgb(Mix(under.R, pattern.R), Mix(under.G, pattern.G), Mix(under.B, pattern.B));
		return shown.ToArgb() == System.Drawing.Color.White.ToArgb() && fill == null ? null : $"{shown.R:X2}{shown.G:X2}{shown.B:X2}";
	}

	/// A frame from a paragraph's (or style's) w:pBdr children by side and w:shd (null: nothing).
	public static ParagraphFrame? From(IReadOnlyDictionary<string, XElement> borders, XElement? shd) {
		var frame = new ParagraphFrame {
			Top = LineOf(borders.GetValueOrDefault("top")), Left = LineOf(borders.GetValueOrDefault("left") ?? borders.GetValueOrDefault("start")),
			Bottom = LineOf(borders.GetValueOrDefault("bottom")), Right = LineOf(borders.GetValueOrDefault("right") ?? borders.GetValueOrDefault("end")),
			Between = LineOf(borders.GetValueOrDefault("between")), Fill = FillOf(shd)
		};
		return frame.Any ? frame : null;
	}

	/// After a .docx is opened (any converter): its paragraphs' borders and shading as markers. The file's paragraphs are
	/// paired up with the editor's by their text, in order (the editor has paragraphs of its own: section breaks, page
	/// breaks; list numbers and tables of contents' page numbers differ): a paragraph that finds no partner gets
	/// nothing. Markers already there (AhkWord's own, written into the file as hidden text) give way to the file's.
	/// Returns how many paragraphs got one.
	public static int FromDocx(EllipticBit.RichEditorNET.RichEditBox editor, string path) {
		var doc = editor.TextDocument;
		List<DocxLineSpacing.Format> formats;
		List<string> texts;
		try {
			formats = DocxLineSpacing.Paragraphs(path);
			texts = ParagraphTexts(path);
		} catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException or NullReferenceException) { return 0; }
		if (formats.Count != texts.Count) return 0;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		bool hadMarkers = Any(story);
		if (!hadMarkers && formats.All(f => f.Borders.Count == 0 && f.Shading == null)) return 0;
		var (starts, partner) = Pair(doc, story, texts);
		// (from the end: each change leaves the paragraphs before it where they were)
		var frames = new ParagraphFrame?[starts.Count];
		for (int i = 0; i < formats.Count; i++) if (partner[i] >= 0) frames[partner[i]] = From(formats[i].Borders, formats[i].Shading);
		int count = 0;
		for (int k = starts.Count - 1; k >= 0; k--) {
			int start = starts[k];
			if (hadMarkers) {    // (a marker of AhkWord's own in the paragraph: out, the file's goes in)
				int end = story.IndexOfAny(new[] { '\r', '\u0007' }, start);
				foreach (Match m in Mark.Matches(story[start..(end < 0 ? story.Length : end)]).Cast<Match>().Reverse()) doc.Range2(start + m.Index, start + m.Index + m.Length).Text = "";
			}
			if (frames[k] is not { } frame) continue;
			Insert(doc, start, frame);
			count++;
		}
		return count;
	}

	/// The editor's paragraphs (their starts: cells' included, rows' marks left out) and, for each of a .docx's
	/// paragraphs (their texts), its partner among them (-1: none): by their text, in order, a few paragraphs ahead at most.
	public static (List<int> Starts, int[] Partner) Pair(ITextDocument2 doc, string story, List<string> texts) {
		var starts = DocxLineSpacing.Starts(doc, int.MaxValue);
		string EditorText(int k) {
			int end = story.IndexOfAny(new[] { '\r', '\u0007' }, starts[k]);
			return story[starts[k]..(end < 0 ? story.Length : end)];
		}
		var mine = starts.Select((_, k) => Key(EditorText(k))).ToList();
		var partner = new int[texts.Count];
		int next = 0;
		for (int i = 0; i < texts.Count; i++) {
			partner[i] = -1;
			string theirs = Key(texts[i]);
			for (int k = next; k < mine.Count && k < next + 12; k++) {
				bool same = mine[k] == theirs || theirs.Length >= 6 && mine[k].Length >= 6 && (mine[k].Contains(theirs) || theirs.Contains(mine[k]));
				if (!same) continue;
				partner[i] = k;
				next = k + 1;
				break;
			}
		}
		return (starts, partner);
	}

	// A paragraph's text as the pairing compares it: without markers, private-use characters, row marks and page
	// breaks; up to its first tab (a table of contents' page number, a list number's tab); spaces as one; lower case.
	static string Key(string text) {
		text = Regex.Replace(text, @"\[\[[^\]\r]*\]\]|[-￹-￻\f\v​-‏]", "");
		int tab = text.IndexOf('\t');
		if (tab >= 0) text = text[..tab].Trim().Length >= 6 ? text[..tab] : text[(tab + 1)..];
		return Regex.Replace(text, @"\s+", " ").Trim().ToLowerInvariant();
	}

	// The .docx's paragraphs' texts, in the order DocxLineSpacing.Paragraphs has them (w:p in the body, its tables'
	// included, text boxes left out; deleted text left out, tabs as tabs).
	internal static List<string> ParagraphTexts(string path) {
		using var zip = System.IO.Compression.ZipFile.OpenRead(path);
		XDocument document;
		using (var s = zip.GetEntry("word/document.xml")!.Open()) document = XDocument.Load(s);
		return document.Root!.Element(W + "body")!.Descendants(W + "p").Where(p => !p.Ancestors(W + "txbxContent").Any())
			.Select(p => string.Concat(p.Descendants().Where(e => e.Name == W + "t" && !e.Ancestors(W + "del").Any() || e.Name == W + "tab" && e.Parent?.Name == W + "r").Select(e => e.Name == W + "tab" ? "\t" : e.Value)))
			.ToList();
	}

	/// Puts a frame's marker at a paragraph's start (hidden text).
	public static void Insert(ITextDocument2 doc, int start, ParagraphFrame frame) {
		var at = doc.Range2(start, start);
		at.Text = MarkerFor(frame);
		at.Font.Hidden = tomConstants.tomTrue;
	}

	// ================================================================ drawing

	/// Where a framed paragraph's text is (its lines' top and bottom, the left and right of its indents), in any unit;
	/// joinedAbove/joinedBelow: boxed together with the paragraph before/after it (on the same page).
	public sealed record Place(Framed Paragraph, RectangleF Text, bool JoinedAbove, bool JoinedBelow);

	/// The area the shading covers: the text, out to the borders' distances (as Word: the lines are outside it).
	public static RectangleF ShadeArea(Place place, float unitsPerPoint) {
		var f = place.Paragraph.Frame;
		var r = place.Text;
		float left = (f.Left?.Space ?? 0) * unitsPerPoint, right = (f.Right?.Space ?? 0) * unitsPerPoint;
		float top = place.JoinedAbove ? 0 : (f.Top?.Space ?? 0) * unitsPerPoint, bottom = place.JoinedBelow ? 0 : (f.Bottom?.Space ?? 0) * unitsPerPoint;
		return RectangleF.FromLTRB(r.Left - left, r.Top - top, r.Right + right, r.Bottom + bottom);
	}

	/// Draws a paragraph's shading (shade: how, so that the screen can put it under text already painted) and its lines
	/// (fill, lines: which of them).
	public static void Draw(Graphics g, Place place, float unitsPerPoint, float minimumLine, Action<RectangleF, Color>? shade = null, bool fill = true, bool lines = true) {
		var f = place.Paragraph.Frame;
		var area = ShadeArea(place, unitsPerPoint);
		if (fill && f.Fill != null && ColorOf(f.Fill) is { } color) {
			if (shade != null) shade(area, color);
			else { using var brush = new SolidBrush(color); g.FillRectangle(brush, area); }
		}
		if (!lines) return;
		float Width(FrameLine line) => Math.Max(minimumLine, line.Size / 8f * unitsPerPoint);
		// (a line's middle half its width outside the shaded area)
		void Line(FrameLine? line, float x1, float y1, float x2, float y2) {
			if (line == null) return;
			float width = Width(line);
			using var pen = new Pen(ColorOf(line.Color) ?? Color.Black, line.Style is "double" ? Math.Max(minimumLine, width / 3) : width);
			if (line.Style is "dotted") pen.DashStyle = DashStyle.Dot;
			else if (line.Style is "dashed" or "dashSmallGap" or "dashDotStroked") pen.DashStyle = DashStyle.Dash;
			else if (line.Style is "dotDash") pen.DashStyle = DashStyle.DashDot;
			else if (line.Style is "dotDotDash") pen.DashStyle = DashStyle.DashDotDot;
			if (line.Style is "double") {
				bool across = y1 == y2;
				float step = width / 3 * 2;
				g.DrawLine(pen, x1 - (across ? 0 : step), y1 - (across ? step : 0), x2 - (across ? 0 : step), y2 - (across ? step : 0));
				g.DrawLine(pen, x1 + (across ? 0 : step), y1 + (across ? step : 0), x2 + (across ? 0 : step), y2 + (across ? step : 0));
			} else g.DrawLine(pen, x1, y1, x2, y2);
		}
		float half(FrameLine? line) => line == null ? 0 : Width(line) / 2;
		float l = area.Left - half(f.Left), r = area.Right + half(f.Right), t = area.Top - half(f.Top), b = area.Bottom + half(f.Bottom);
		if (!place.JoinedAbove) Line(f.Top, l, t, r, t);
		if (!place.JoinedBelow) Line(f.Bottom, l, b, r, b);
		else Line(f.Between, l, area.Bottom, r, area.Bottom);
		Line(f.Left, l, place.JoinedAbove ? area.Top : t, l, place.JoinedBelow ? area.Bottom : b);
		Line(f.Right, r, place.JoinedAbove ? area.Top : t, r, place.JoinedBelow ? area.Bottom : b);
	}

	/// A picture's area multiplied by a colour (white becomes the colour; what is darker stays darker): shading put under
	/// what is drawn already.
	public static void Multiply(Bitmap picture, Rectangle area, Color color) {
		area.Intersect(new Rectangle(0, 0, picture.Width, picture.Height));
		if (area.Width <= 0 || area.Height <= 0) return;
		var data = picture.LockBits(area, System.Drawing.Imaging.ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
		try {
			var row = new byte[area.Width * 4];
			for (int y = 0; y < area.Height; y++) {
				IntPtr line = data.Scan0 + y * data.Stride;
				System.Runtime.InteropServices.Marshal.Copy(line, row, 0, row.Length);
				for (int x = 0; x < row.Length; x += 4) {
					row[x] = (byte)(row[x] * color.B / 255);
					row[x + 1] = (byte)(row[x + 1] * color.G / 255);
					row[x + 2] = (byte)(row[x + 2] * color.R / 255);
				}
				System.Runtime.InteropServices.Marshal.Copy(row, 0, line, row.Length);
			}
		} finally {
			picture.UnlockBits(data);
		}
	}

	public static Color? ColorOf(string hex) =>
		hex is { Length: 6 } && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int v) ? Color.FromArgb(v >> 16 & 255, v >> 8 & 255, v & 255) : null;

	// ================================================================ RTF

	/// Word's RTF: each framed paragraph gets its border and shading words (in the paragraph, after its marker), and the
	/// paragraph after it, if it has no \pard of its own, words that take them off again. The fill colours go into the
	/// colour table.
	public static string ToWordRtf(string rtf) {
		if (!rtf.Contains("[[para:", StringComparison.Ordinal)) return rtf;
		var colors = ColorTable(ref rtf);
		var marks = RtfMark.Matches(rtf).Cast<Match>().ToList();
		for (int k = marks.Count - 1; k >= 0; k--) {
			var m = marks[k];
			if (Parse(RtfUnescape(m.Groups[1].Value)) is not { Any: true } frame) continue;
			// (the paragraph's end: its \par or \cell; the next paragraph's own \pard, if it has one)
			int from = m.Index + m.Length;
			var end = Regex.Match(rtf[from..], @"\\(par|cell)(?![a-zA-Z])");
			string words = Words(frame, colors);
			if (end.Success) {
				int after = from + end.Index + end.Length;
				var next = Regex.Match(rtf[after..Math.Min(rtf.Length, after + 200)], @"^(?:\s|\\[a-zA-Z]+-?\d* ?|\{\\\*?\\[a-zA-Z]+[^{}]*\})*");
				if (!next.Value.Contains(@"\pard", StringComparison.Ordinal)) rtf = rtf.Insert(after, @"\brdrt\brdrnone\brdrl\brdrnone\brdrb\brdrnone\brdrr\brdrnone\brdrbtw\brdrnone\cbpat0\shading0 ");
			}
			// (after the hidden marker's group or \v0: the words in the paragraph's own text)
			var close = Regex.Match(rtf[from..Math.Min(rtf.Length, from + 40)], @"^(?:\}|\\v0 ?)");
			int at = from + (close.Success ? close.Length : 0);
			rtf = rtf.Insert(at, words + " ");
		}
		return rtf;
	}

	// a frame's RTF paragraph words
	static string Words(ParagraphFrame f, Func<string, int> color) {
		var sb = new System.Text.StringBuilder();
		void Side(string word, FrameLine? line) {
			if (line == null) return;
			string style = line.Style switch {
				"double" => @"\brdrdb", "dotted" => @"\brdrdot", "dashed" => @"\brdrdash", "dashSmallGap" => @"\brdrdashsm", "dotDash" => @"\brdrdashd",
				"dotDotDash" => @"\brdrdashdd", "thick" => @"\brdrth", "triple" => @"\brdrtriple", "wave" => @"\brdrwavy", _ => @"\brdrs"
			};
			// (\brdrw: twips, the line's width; \brsp: twips, its distance)
			sb.Append($@"\{word}{style}\brdrw{Math.Max(1, (int)Math.Round(line.Size * 2.5))}\brsp{line.Space * 20}\brdrcf{color(line.Color)}");
		}
		Side("brdrt", f.Top); Side("brdrl", f.Left); Side("brdrb", f.Bottom); Side("brdrr", f.Right); Side("brdrbtw", f.Between);
		if (f.Fill != null) sb.Append($@"\cbpat{color(f.Fill)}\shading0");
		return sb.ToString();
	}

	static readonly Regex RtfMark = new(@"\[\[para:((?:[^\]\\]|\\[\\{}]|\\'[0-9a-fA-F]{2}|\\u-?\d+\??|\\[a-zA-Z]+-?\d* ?|\](?!\]))*)\]\]");

	// The RTF of a marker's JSON back to text: \{ \} \\ and \'hh, \uN
	static string RtfUnescape(string s) {
		s = Regex.Replace(s, @"\\'([0-9a-fA-F]{2})", m => ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString());
		s = Regex.Replace(s, @"\\u(-?\d+)\??", m => ((char)(ushort)short.Parse(m.Groups[1].Value)).ToString());
		s = Regex.Replace(s, @"\\([\\{}])", "$1");
		return Regex.Replace(s, @"\\[a-zA-Z]+-?\d* ?", "");
	}

	// The colour table's colours (RRGGBB) by index, a colour added to it where it isn't there: its index.
	static Func<string, int> ColorTable(ref string rtf) {
		var table = Regex.Match(rtf, @"\{\\colortbl\s*;?((?:\s*\\red\d+\\green\d+\\blue\d+\s*;)*)\s*\}");
		var known = new List<string>();
		if (table.Success)
			foreach (Match c in Regex.Matches(table.Groups[1].Value, @"\\red(\d+)\\green(\d+)\\blue(\d+)"))
				known.Add($"{int.Parse(c.Groups[1].Value):X2}{int.Parse(c.Groups[2].Value):X2}{int.Parse(c.Groups[3].Value):X2}");
		var added = new List<string>();
		int Index(string hex) {
			int at = known.IndexOf(hex);
			if (at >= 0) return at + 1;
			at = added.IndexOf(hex);
			if (at < 0) { added.Add(hex); at = added.Count - 1; }
			return known.Count + at + 1;
		}
		// (the colours this RTF's frames use, added now: the table goes in front of the text, so before the words do)
		foreach (Match m in RtfMark.Matches(rtf))
			if (Parse(RtfUnescape(m.Groups[1].Value)) is { } f)
				foreach (var hex in new[] { f.Fill, f.Top?.Color, f.Left?.Color, f.Bottom?.Color, f.Right?.Color, f.Between?.Color })
					if (hex != null) Index(hex);
		string entries = string.Concat(added.Select(h => $@"\red{Convert.ToInt32(h[..2], 16)}\green{Convert.ToInt32(h[2..4], 16)}\blue{Convert.ToInt32(h[4..], 16)};"));
		if (added.Count > 0) {
			if (table.Success) rtf = rtf.Insert(table.Index + table.Length - 1, entries);
			else {
				var fonts = Regex.Match(rtf, @"\{\\fonttbl(?:[^{}]|\{[^{}]*\})*\}");
				int at = fonts.Success ? fonts.Index + fonts.Length : rtf.IndexOf('{', 1) is int i and > 0 ? i : 6;
				rtf = rtf.Insert(at, @"{\colortbl ;" + entries + "}");
			}
		}
		var all = known.Concat(added).ToList();
		return hex => all.IndexOf(hex) + 1;
	}
}

/// (Docxodus trial) What follows a list number, where it isn't a tab: Word's "space" or "nothing". RichEdit always
/// starts a numbered paragraph's first line at its indent, the number in front of it; with no room in front of the text
/// (no hanging indent) it puts the text just after the number instead, a small gap after it. So such a paragraph has no
/// hanging indent in the editor (its lines start where its number does), and a hidden marker at its start keeps what
/// it is and the indents its file gave it: [[sfx:space,720,-360]] (twips: left, first line). Saved, the paragraph gets
/// its own indents back and \rtesuffix (1 space, 2 nothing), which the direct .docx writer makes the level's w:suff.
public static class ListSuffixes
{
	public static readonly Regex Mark = new(@"\[\[sfx:(space|nothing),(-?\d+),(-?\d+)\]\]");
	public static bool Any(string story) => story.Contains("[[sfx:", StringComparison.Ordinal);

	/// After a .docx is opened: its list paragraphs whose number is followed by a space or nothing get their markers
	/// (paired up with the editor's by their text, as the frames are; AhkWord's own markers give way to the file's), and
	/// their layout (Apply). Returns how many.
	public static int FromDocx(EllipticBit.RichEditorNET.RichEditBox editor, string path) {
		var doc = editor.TextDocument;
		List<(string Id, string Definition, int Level, bool Bullet, string Suffix)?> lists;
		List<string> texts;
		try {
			lists = DocxodusReader.ListParagraphs(path, allParagraphs: true);
			texts = ParagraphFrames.ParagraphTexts(path);
		} catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException or NullReferenceException) { return 0; }
		if (lists.Count != texts.Count) return 0;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		bool hadMarkers = Any(story);
		if (!hadMarkers && !lists.Any(l => l?.Suffix is "space" or "nothing")) return 0;
		var (starts, partner) = ParagraphFrames.Pair(doc, story, texts);
		var suffixes = new string?[starts.Count];
		for (int i = 0; i < lists.Count; i++) if (partner[i] >= 0 && lists[i]?.Suffix is "space" or "nothing") suffixes[partner[i]] = lists[i]!.Value.Suffix;
		int count = 0;
		for (int k = starts.Count - 1; k >= 0; k--) {
			int start = starts[k];
			if (hadMarkers) {
				int end = story.IndexOfAny(new[] { '\r', '\u0007' }, start);
				foreach (Match m in Mark.Matches(story[start..(end < 0 ? story.Length : end)]).Cast<Match>().Reverse()) doc.Range2(start + m.Index, start + m.Index + m.Length).Text = "";
			}
			if (suffixes[k] is not { } suffix) continue;
			var para = doc.Range2(start, start).Para;
			if ((para.ListType & 0xFFFF) == tomConstants.tomListNone) continue;    // (a number the editor keeps as text)
			var at = doc.Range2(start, start);
			at.Text = $"[[sfx:{suffix},{Twips(para.LeftIndent)},{Twips(para.FirstLineIndent)}]]";
			at.Font.Hidden = tomConstants.tomTrue;
			count++;
		}
		if (count > 0) Apply(doc);
		return count;
	}

	static int Twips(float points) => (int)Math.Round(points * 20);

	/// The marked list paragraphs laid out with their text after their number: no hanging indent, their lines from where
	/// the number starts (after opening; an .rtf of AhkWord's own has the file's indents in it).
	/// The markers back at their paragraphs' starts (an .rtf of AhkWord's own has them at the ends, where saving put
	/// them). Before the lists' restarts are given back: moving them loses those.
	public static void ToStarts(ITextDocument2 doc) {
		if (Any(doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "")) Move(doc, toEnd: false);
	}

	public static void Apply(ITextDocument2 doc) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (!Any(story)) return;
		foreach (Match m in Mark.Matches(story)) {
			var para = doc.Range2(m.Index, m.Index).Para;
			if ((para.ListType & 0xFFFF) == tomConstants.tomListNone) continue;
			float numberAt = (int.Parse(m.Groups[2].Value) + int.Parse(m.Groups[3].Value)) / 20f;
			if (Math.Abs(para.FirstLineIndent) < 0.01f && Math.Abs(para.LeftIndent - numberAt) < 0.01f) continue;
			para.SetIndents(0, Math.Max(0, numberAt), para.RightIndent);    // (the indents only: the list's type and restart stay as they are)
		}
	}

	/// For saving: the marked paragraphs with their file's indents while action runs (RichEdit writes a numbered
	/// paragraph with no hanging indent as one with its number typed in), laid out again after it; not undoable, the
	/// document's modified state as it was.
	public static void WithFileIndents(EllipticBit.RichEditorNET.RichEditBox editor, Action action) {
		var doc = editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (!Any(story)) { action(); return; }
		bool modified = editor.Modified;
		doc.Undo(-9999995);    // tomSuspend
		// (the marked paragraphs' list restarts: putting the markers back at their starts loses them)
		var restarts = OfficeConverters.NewNumbers(editor, Mark.Matches(story).Select(m => m.Index).ToList());
		try {
			foreach (Match m in Mark.Matches(story)) {
				var para = doc.Range2(m.Index, m.Index).Para;
				if ((para.ListType & 0xFFFF) == tomConstants.tomListNone) continue;
				para.SetIndents(int.Parse(m.Groups[3].Value) / 20f, int.Parse(m.Groups[2].Value) / 20f, para.RightIndent);
			}
			// (the markers at their paragraphs' ends meanwhile: the steps that make lists Word's read a list paragraph's
			// start for a number typed in, and took one there for that)
			Move(doc, toEnd: true);
			action();
		} finally {
			Move(doc, toEnd: false);
			Apply(doc);
			OfficeConverters.NewNumbersBack(editor, restarts);
			doc.Undo(-9999994);    // tomResume
			editor.Modified = modified;
		}
	}

	// The markers moved to their paragraphs' ends (before the paragraph mark), or back to their starts.
	static void Move(ITextDocument2 doc, bool toEnd) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		foreach (Match m in Mark.Matches(story).Cast<Match>().Reverse()) {
			int to = toEnd ? story.IndexOfAny(new[] { '\r', '\u0007' }, m.Index) : (m.Index == 0 ? 0 : story.LastIndexOfAny(new[] { '\r', '\u0007' }, m.Index - 1) + 1);
			if (to < 0 || (toEnd ? to == m.Index + m.Length : to == m.Index)) continue;
			doc.Range2(m.Index, m.Index + m.Length).Text = "";
			if (toEnd) to -= m.Length;
			var at = doc.Range2(to, to);
			at.Text = m.Value;
			at.Font.Hidden = tomConstants.tomTrue;
		}
	}

	/// Word's RTF: each marked paragraph its file's indents back and \rtesuffix (1 space, 2 nothing), after its marker;
	/// the paragraph after it, without a \pard of its own, the indents it had (and \rtesuffix0).
	public static string ToWordRtf(string rtf) {
		if (!rtf.Contains("[[sfx:", StringComparison.Ordinal)) return rtf;
		var marks = Mark.Matches(rtf).Cast<Match>().ToList();
		for (int k = marks.Count - 1; k >= 0; k--) {
			var m = marks[k];
			int from = m.Index + m.Length;
			// (the indents RichEdit wrote for it: the last \li and \fi since its \pard)
			int pard = rtf.LastIndexOf(@"\pard", m.Index, StringComparison.Ordinal);
			string head = pard < 0 ? "" : rtf[pard..m.Index];
			string li = Regex.Matches(head, @"\\li(-?\d+)").Cast<Match>().LastOrDefault()?.Groups[1].Value ?? "0";
			string fi = Regex.Matches(head, @"\\fi(-?\d+)").Cast<Match>().LastOrDefault()?.Groups[1].Value ?? "0";
			var end = Regex.Match(rtf[from..], @"\\(par|cell)(?![a-zA-Z])");
			if (end.Success) {
				int after = from + end.Index + end.Length;
				var next = Regex.Match(rtf[after..Math.Min(rtf.Length, after + 200)], @"^(?:\s|\\[a-zA-Z]+-?\d* ?|\{\\\*?\\[a-zA-Z]+[^{}]*\})*");
				if (!next.Value.Contains(@"\pard", StringComparison.Ordinal)) rtf = rtf.Insert(after, $@"\li{li}\fi{fi}\rtesuffix0 ");
			}
			var close = Regex.Match(rtf[from..Math.Min(rtf.Length, from + 40)], @"^(?:\}|\\v0 ?)");
			int at = from + (close.Success ? close.Length : 0);
			rtf = rtf.Insert(at, $@"\li{m.Groups[2].Value}\fi{m.Groups[3].Value}\rtesuffix{(m.Groups[1].Value == "space" ? 1 : 2)} ");
		}
		return rtf;
	}
}
