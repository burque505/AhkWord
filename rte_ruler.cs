// Iteration 78: the ruler (View > Show > Ruler), as Word's horizontal ruler, for looking only: the scale (inches, or
// centimetres where the region measures in them) with 0 at the text's left edge; the page's margins in grey and the
// text's width in white (in Print Layout; otherwise the window's text width, white throughout); the paragraph at the
// caret's (with a selection, its first paragraph's) first-line indent (the triangle at the top), hanging indent (the
// triangle at the bottom) and left indent (the box under it), right indent (the triangle at the right), its own tab
// stops (left, centre, right, decimal and bar marks) and the default tab stops after them (small grey ticks). In a
// table it shows the scale only. It sits above the page and doesn't scroll down with it; sideways it follows the page.

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

sealed class RulerView : Control
{
	/// Where the page is, in the ruler's pixels: the paper's left edge, the text's left and right edges, the paper's right
	/// edge; and pixels per point at the zoom.
	public readonly record struct Frame(int PaperLeft, int TextLeft, int TextRight, int PaperRight, float PixelsPerPoint);

	/// The paragraph shown, in points: the first line's indent (relative to the left one, as TOM gives it), the left and
	/// right indents, its tab stops (position, TOM alignment) and the document's default tab spacing.
	public sealed record Paragraph(float First, float Left, float Right, (float Position, int Alignment)[] Tabs, float DefaultTab)
	{
		public bool Same(Paragraph? other) => other != null && other.First == First && other.Left == Left && other.Right == Right
			&& other.DefaultTab == DefaultTab && other.Tabs.SequenceEqual(Tabs);
	}

	Frame frame;
	Paragraph? paragraph;
	readonly bool metric = RegionInfo.CurrentRegion.IsMetric;

	public static readonly Color Margin = Color.FromArgb(0xC8, 0xC8, 0xC8), Ink = Color.FromArgb(0x50, 0x50, 0x50), Marker = Color.FromArgb(0xF4, 0xF4, 0xF4);

	public RulerView() {
		SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
		SetStyle(ControlStyles.Selectable, false);
		TabStop = false;
		Dock = DockStyle.Top;
	}

	public Frame Shown => frame;
	public Paragraph? ShownParagraph => paragraph;

	/// What to show; repaints only if it changed.
	public void Show(Frame newFrame, Paragraph? newParagraph) {
		if (newFrame == frame && (newParagraph == null ? paragraph == null : newParagraph.Same(paragraph))) return;
		frame = newFrame;
		paragraph = newParagraph;
		Invalidate();
	}

	/// Points per unit of the scale: an inch, or a centimetre.
	public float PointsPerUnit => metric ? 72f / 2.54f : 72f;
	public bool Metric => metric;

	/// The ruler's x for a position in points from the text's left edge.
	public float X(float points) => frame.TextLeft + points * frame.PixelsPerPoint;

	protected override void OnPaint(PaintEventArgs e) {
		var g = e.Graphics;
		g.Clear(PrintLayoutPage.Surround);    // (around the paper: the theme's grey)
		if (frame.PixelsPerPoint <= 0 || frame.PaperRight <= frame.PaperLeft) return;
		int top = LogicalToDeviceUnits(3), bottom = Height - LogicalToDeviceUnits(5);
		using (var margin = new SolidBrush(Margin)) g.FillRectangle(margin, frame.PaperLeft, top, frame.PaperRight - frame.PaperLeft, bottom - top);
		using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, frame.TextLeft, top, Math.Max(0, frame.TextRight - frame.TextLeft), bottom - top);
		DrawScale(g, top, bottom);
		if (paragraph is { } p) {
			DrawDefaultTabs(g, p, bottom);
			foreach (var (position, alignment) in p.Tabs) DrawTab(g, X(position), alignment, bottom);
			g.SmoothingMode = SmoothingMode.AntiAlias;
			int size = LogicalToDeviceUnits(4);
			DrawTriangle(g, X(p.Left + p.First), top, size, down: true);    // first line indent
			DrawTriangle(g, X(p.Left), bottom, size, down: false);          // hanging indent
			float left = X(p.Left);
			using (var fill = new SolidBrush(Marker)) using (var pen = new Pen(Ink)) {    // left indent
				var box = new RectangleF(left - size, bottom, 2 * size, Height - bottom - 1);
				g.FillRectangle(fill, box);
				g.DrawRectangle(pen, box.X, box.Y, box.Width, box.Height);
			}
			DrawTriangle(g, frame.TextRight - p.Right * frame.PixelsPerPoint, bottom, size, down: false);    // right indent
		}
	}

	// The ticks and numbers: within the paper, counted from the text's left edge both ways (numbers without a sign).
	void DrawScale(Graphics g, int top, int bottom) {
		float unit = PointsPerUnit * frame.PixelsPerPoint;    // pixels per inch or centimetre
		int parts = metric ? 4 : 8;
		while (parts > 1 && unit / parts < LogicalToDeviceUnits(4)) parts /= 2;
		int middle = (top + bottom) / 2;
		using var pen = new Pen(Ink);
		using var font = new Font("Segoe UI", 7f);
		using var ink = new SolidBrush(Ink);
		int numberEvery = unit < LogicalToDeviceUnits(18) ? 2 : 1;
		int first = (int)Math.Floor((frame.PaperLeft - frame.TextLeft) / unit * parts), last = (int)Math.Ceiling((frame.PaperRight - frame.TextLeft) / unit * parts);
		for (int i = first; i <= last; i++) {
			float x = frame.TextLeft + i * unit / parts;
			if (x < frame.PaperLeft || x > frame.PaperRight) continue;
			int px = (int)Math.Round(x);
			if (i % parts == 0) {
				int n = Math.Abs(i / parts);
				if (n == 0 || n % numberEvery != 0) continue;
				string text = n.ToString(CultureInfo.InvariantCulture);
				var size = g.MeasureString(text, font);
				g.DrawString(text, font, ink, x - size.Width / 2, middle - size.Height / 2);
			} else {
				int length = i % Math.Max(1, parts / 2) == 0 ? LogicalToDeviceUnits(3) : LogicalToDeviceUnits(1);    // (half a unit longer)
				g.DrawLine(pen, px, middle - length, px, middle + length - 1);
			}
		}
	}

	// Word's default tab stops: after the last tab stop of its own (and the left indent), small ticks at the bottom.
	void DrawDefaultTabs(Graphics g, Paragraph p, int bottom) {
		if (p.DefaultTab <= 0) return;
		float from = Math.Max(p.Left, p.Tabs.Length > 0 ? p.Tabs.Max(t => t.Position) : 0);
		using var pen = new Pen(Color.FromArgb(0x90, 0x90, 0x90));
		for (float at = (float)(Math.Floor(from / p.DefaultTab) + 1) * p.DefaultTab; ; at += p.DefaultTab) {
			float x = X(at);
			if (x > frame.TextRight - p.Right * frame.PixelsPerPoint || x > frame.PaperRight) break;
			g.DrawLine(pen, x, bottom, x, bottom + LogicalToDeviceUnits(2));
		}
	}

	// A tab stop, at the strip's bottom (TOM: 0 left, 1 centre, 2 right, 3 decimal, 4 bar).
	void DrawTab(Graphics g, float x, int alignment, int bottom) {
		using var pen = new Pen(Ink, Math.Max(1, LogicalToDeviceUnits(2) / 1.5f));
		int h = LogicalToDeviceUnits(6), w = LogicalToDeviceUnits(4);
		float y = bottom - 1, up = y - h;
		switch (alignment) {
			case 1: g.DrawLine(pen, x, up, x, y); g.DrawLine(pen, x - w, y, x + w, y); break;
			case 2: g.DrawLine(pen, x, up, x, y); g.DrawLine(pen, x - w, y, x, y); break;
			case 3:
				g.DrawLine(pen, x, up, x, y); g.DrawLine(pen, x - w, y, x + w, y);
				using (var dot = new SolidBrush(Ink)) g.FillRectangle(dot, x + w / 2f, up + h / 3f, 2, 2);
				break;
			case 4: g.DrawLine(pen, x, up - 2, x, y); break;
			default: g.DrawLine(pen, x, up, x, y); g.DrawLine(pen, x, y, x + w, y); break;
		}
	}

	void DrawTriangle(Graphics g, float x, float y, int size, bool down) {
		var points = down
			? new[] { new PointF(x - size, y), new PointF(x + size, y), new PointF(x, y + size + 1) }
			: new[] { new PointF(x - size, y), new PointF(x + size, y), new PointF(x, y - size - 1) };
		using var fill = new SolidBrush(Marker);
		using var pen = new Pen(Ink);
		g.FillPolygon(fill, points);
		g.DrawPolygon(pen, points);
	}
}
