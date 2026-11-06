// Iteration 78: the ruler (View > Show > Ruler), as Word's horizontal ruler, for looking only: the scale (inches, or
// centimetres where the region measures in them) with 0 at the text's left edge; the page's margins in grey and the
// text's width in white (in Print Layout; otherwise the window's text width, white throughout); the paragraph at the
// caret's (with a selection, its first paragraph's) first-line indent (the triangle at the top), hanging indent (the
// triangle at the bottom) and left indent (the box under it), right indent (the triangle at the right), its own tab
// stops (left, centre, right, decimal and bar marks) and the default tab stops after them (small grey ticks). In a
// table it shows the scale only. It sits above the page and doesn't scroll down with it; sideways it follows the page.
// Iteration 83: a click on the white strip adds a left tab stop there (on Word's grid) to the selected paragraphs; a
// double-click opens Tabs on a tab stop, Paragraph on an indent marker, Page Setup on the grey margins.
// Iteration 84: dragging, as on Word's: a tab stop to move it (off the ruler to clear it), the indent markers (the box
// moves the first line with the left indent, the bottom triangle only the left indent), a press on the strip dragged to
// put the new tab stop where it's let go; on Word's grid (Alt: without it), a dotted line down the page while dragging.

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

	// ---- iteration 83: clicking on the ruler, as on Word's

	/// What a point of the ruler is on: nothing (outside the paper), the white text strip, the grey margins, a tab stop
	/// of the paragraph's or one of its indent markers.
	public enum Part { None, Text, Margin, Tab, Indent }

	/// A click on the white strip, off the markers (not in a table): where, in points from the text's left edge, on
	/// Word's grid (1/16 inch, or 0.25 cm).
	public event Action<float>? TextClicked;
	/// A double-click: what it was on.
	public event Action<Part>? PartDoubleClicked;

	/// Word's grid for a click: 1/16 inch, or a quarter of a centimetre (points).
	public float Snap => metric ? 72f / 2.54f / 4 : 4.5f;

	/// Where a click at x lands, in points from the text's left edge, on the grid and within the text.
	public float PointsAt(float x) {
		if (frame.PixelsPerPoint <= 0) return 0;
		float width = (frame.TextRight - frame.TextLeft) / frame.PixelsPerPoint;
		float points = (float)Math.Round((x - frame.TextLeft) / frame.PixelsPerPoint / Snap) * Snap;
		return Math.Clamp(points, 0, Math.Max(0, width));
	}

	/// What the point at is on (markers win over the strip; tab stops in the strip's lower half, the first line's
	/// triangle at its top, the others at its bottom).
	public Part HitTest(Point at) {
		if (frame.PixelsPerPoint <= 0 || at.X < frame.PaperLeft || at.X > frame.PaperRight) return Part.None;
		int top = LogicalToDeviceUnits(3), bottom = Height - LogicalToDeviceUnits(5), size = LogicalToDeviceUnits(4), grab = size + 1;
		int middle = (top + bottom) / 2;
		bool Near(float x) => Math.Abs(at.X - x) <= grab;
		if (paragraph is { } p) {
			if (at.Y <= top + size + 2 && Near(X(p.Left + p.First))) return Part.Indent;
			if (at.Y >= bottom - size - 2 && (Near(X(p.Left)) || Near(frame.TextRight - p.Right * frame.PixelsPerPoint))) return Part.Indent;
			if (at.Y >= middle && p.Tabs.Any(t => Near(X(t.Position)))) return Part.Tab;
		}
		return at.X >= frame.TextLeft && at.X <= frame.TextRight ? Part.Text : Part.Margin;
	}

	// ---- iteration 84: dragging on the ruler, as on Word's

	/// What a press on the ruler took hold of: a tab stop, the first line's triangle, the hanging indent's triangle, the
	/// left indent's box (it moves the first line with it), the right indent's triangle, or the white strip (a new tab).
	public enum Handle { None, Tab, FirstLine, Hanging, LeftIndent, RightIndent, NewTab }

	/// A marker dragged and let go: the tab stop at From (Handle.Tab) moved to To, or cleared (To null: dragged off the
	/// ruler); an indent marker's new place (To, in points from the text's left edge; From where it was).
	public event Action<Handle, float, float?>? Dragged;
	/// While dragging: what, and where it would go (null: off the ruler, a tab stop to clear); (None, null) at the end.
	public event Action<Handle, float?>? Dragging;

	Handle grabbed;
	float grabbedFrom, dragAt;
	int grabbedAlignment;
	bool dragging, dragOff;
	Point downAt;
	int guideX = int.MinValue;

	/// The text's width in points.
	public float TextWidth => frame.PixelsPerPoint <= 0 ? 0 : (frame.TextRight - frame.TextLeft) / frame.PixelsPerPoint;

	/// What a press at a point takes hold of, and where that is (points from the text's left edge).
	public (Handle Handle, float At) Grab(Point at) {
		if (paragraph is not { } p || frame.PixelsPerPoint <= 0 || at.X < frame.PaperLeft || at.X > frame.PaperRight) return (Handle.None, 0);
		int top = LogicalToDeviceUnits(3), bottom = Height - LogicalToDeviceUnits(5), size = LogicalToDeviceUnits(4), grab = size + 1;
		int middle = (top + bottom) / 2;
		bool Near(float x) => Math.Abs(at.X - x) <= grab;
		if (at.Y <= top + size + 2 && Near(X(p.Left + p.First))) return (Handle.FirstLine, p.Left + p.First);
		if (at.Y >= bottom - size - 2) {
			if (Near(X(p.Left))) return (at.Y >= bottom ? Handle.LeftIndent : Handle.Hanging, p.Left);
			if (Near(X(TextWidth - p.Right))) return (Handle.RightIndent, TextWidth - p.Right);
		}
		if (at.Y >= middle && p.Tabs.Where(t => Near(X(t.Position))).OrderBy(t => Math.Abs(at.X - X(t.Position))).ToArray() is { Length: > 0 } near)
			return (Handle.Tab, near[0].Position);
		return HitTest(at) == Part.Text ? (Handle.NewTab, PointsAt(at.X)) : (Handle.None, 0);
	}

	/// Where a dragged marker at x goes (points from the text's left edge): on Word's grid unless snap is off (Alt, as in
	/// Word), within the text for a tab stop, within the paper for an indent, keeping a quarter inch of text between them.
	public float DragTo(Handle handle, float x, bool snap = true) {
		if (paragraph is not { } p || frame.PixelsPerPoint <= 0) return 0;
		float points = (x - frame.TextLeft) / frame.PixelsPerPoint;
		float Grid(float v) => snap ? (float)Math.Round(v / Snap) * Snap : (float)Math.Round(v, 1);
		// (the right indent on the grid from the right edge: it's measured from there)
		points = handle == Handle.RightIndent ? TextWidth - Grid(TextWidth - points) : Grid(points);
		float paperMin = (frame.PaperLeft - frame.TextLeft) / frame.PixelsPerPoint, paperMax = (frame.PaperRight - frame.TextLeft) / frame.PixelsPerPoint;
		const float room = 18;
		float right = TextWidth - p.Right;
		float Clamp(float v, float min, float max) => Math.Clamp(v, min, Math.Max(min, max));
		return handle switch {
			Handle.Tab or Handle.NewTab => Clamp(points, Snap, TextWidth),
			Handle.FirstLine or Handle.Hanging => Clamp(points, paperMin, right - room),
			Handle.LeftIndent => Clamp(points, Math.Max(paperMin, paperMin - p.First), Math.Min(right - room, right - room - p.First)),
			Handle.RightIndent => Clamp(points, Math.Max(p.Left, p.Left + p.First) + room, paperMax),
			_ => points,
		};
	}

	/// The paragraph as the drag would leave it (the paragraph itself when nothing is dragged).
	public Paragraph? Preview() {
		if (!dragging || paragraph is not { } p) return paragraph;
		switch (grabbed) {
			case Handle.Tab or Handle.NewTab:
				var tabs = p.Tabs.Where(t => grabbed == Handle.NewTab || Math.Abs(t.Position - grabbedFrom) > 0.01f).ToList();
				if (!dragOff) { tabs.RemoveAll(t => Math.Abs(t.Position - dragAt) < 0.01f); tabs.Add((dragAt, grabbedAlignment)); }
				return p with { Tabs = tabs.OrderBy(t => t.Position).ToArray() };
			case Handle.FirstLine: return p with { First = dragAt - p.Left };
			case Handle.Hanging: return p with { Left = dragAt, First = p.Left + p.First - dragAt };
			case Handle.LeftIndent: return p with { Left = dragAt };
			case Handle.RightIndent: return p with { Right = TextWidth - dragAt };
			default: return p;
		}
	}

	protected override void OnMouseDown(MouseEventArgs e) {
		base.OnMouseDown(e);
		EndDrag();
		if (e.Button != MouseButtons.Left || e.Clicks > 1) return;
		(grabbed, grabbedFrom) = Grab(e.Location);
		grabbedAlignment = grabbed == Handle.Tab && paragraph is { } p ? p.Tabs.First(t => t.Position == grabbedFrom).Alignment : 0;
		downAt = e.Location;
		dragAt = grabbedFrom;
	}

	protected override void OnMouseMove(MouseEventArgs e) {
		base.OnMouseMove(e);
		if (grabbed == Handle.None) {
			Cursor = Grab(e.Location).Handle is Handle.None or Handle.NewTab ? Cursors.Default : Cursors.SizeWE;
			return;
		}
		if (!dragging) {
			var drag = SystemInformation.DragSize;
			if (Math.Abs(e.X - downAt.X) < drag.Width && Math.Abs(e.Y - downAt.Y) < drag.Height) return;
			dragging = true;
		}
		dragAt = DragTo(grabbed, e.X, (ModifierKeys & Keys.Alt) == 0);
		dragOff = grabbed is Handle.Tab or Handle.NewTab && (e.Y < -Height / 2 || e.Y > Height * 3 / 2);
		Guide(dragOff ? int.MinValue : (int)Math.Round(X(dragAt)));
		Invalidate();
		Dragging?.Invoke(grabbed, dragOff ? null : dragAt);
	}

	protected override void OnMouseUp(MouseEventArgs e) {
		base.OnMouseUp(e);
		if (e.Button != MouseButtons.Left || grabbed == Handle.None) return;
		var (handle, from, at, off, moved) = (grabbed, grabbedFrom, dragAt, dragOff, dragging);
		EndDrag();
		if (handle == Handle.NewTab) { if (!off) TextClicked?.Invoke(moved ? at : from); }    // (a click: as in iteration 83)
		else if (moved && (off || Math.Abs(at - from) > 0.01f)) Dragged?.Invoke(handle, from, off ? null : at);
	}

	protected override void OnMouseCaptureChanged(EventArgs e) {
		base.OnMouseCaptureChanged(e);
		if (!Capture) EndDrag();    // (a dialog, or another window, took the mouse: nothing changes)
	}

	void EndDrag() {
		bool was = dragging;
		grabbed = Handle.None;
		dragging = dragOff = false;
		Guide(int.MinValue);
		if (was) { Invalidate(); Dragging?.Invoke(Handle.None, null); }
	}

	// Word's dotted line down the page under a dragged marker (drawn over the screen and drawn again to take it away).
	void Guide(int x) {
		if (x == guideX) return;
		if (guideX != int.MinValue) DrawGuide(guideX);
		guideX = x;
		if (guideX != int.MinValue) DrawGuide(guideX);
	}

	void DrawGuide(int x) {
		if (Parent == null || !IsHandleCreated) return;
		var top = PointToScreen(new Point(x, Height));
		var bottom = Parent.PointToScreen(new Point(0, Parent.ClientSize.Height));
		ControlPaint.DrawReversibleLine(top, new Point(top.X, bottom.Y), Color.Gray);
	}

	protected override void OnMouseDoubleClick(MouseEventArgs e) {
		base.OnMouseDoubleClick(e);
		if (e.Button != MouseButtons.Left) return;
		var part = HitTest(e.Location);
		if (part != Part.None) PartDoubleClicked?.Invoke(part);
	}

	protected override void OnPaint(PaintEventArgs e) {
		var g = e.Graphics;
		g.Clear(PrintLayoutPage.Surround);    // (around the paper: the theme's grey)
		if (frame.PixelsPerPoint <= 0 || frame.PaperRight <= frame.PaperLeft) return;
		int top = LogicalToDeviceUnits(3), bottom = Height - LogicalToDeviceUnits(5);
		using (var margin = new SolidBrush(Margin)) g.FillRectangle(margin, frame.PaperLeft, top, frame.PaperRight - frame.PaperLeft, bottom - top);
		using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, frame.TextLeft, top, Math.Max(0, frame.TextRight - frame.TextLeft), bottom - top);
		DrawScale(g, top, bottom);
		if (Preview() is { } p) {    // (being dragged: as the drag would leave it)
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
