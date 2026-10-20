// The Multiple Pages view and Print Preview for RTE Writer (rte_writer37.cs, through #load). PageOverview: the printed
// pages side by side, as many as fit across the window, read-only; PagePreviewForm puts it in a window of its own, with
// a toolbar (Print, the pages, zoom), sized to the screen's working area. The pages are the printer's own (DocumentPrinter through a PreviewPrintController:
// the same page breaks, header, footer and line numbers as printing), each drawn once at the size shown. Ctrl+wheel,
// or Zoom, shows them larger or smaller; double-clicking a page (or Enter) asks to edit the document there (PageChosen);
// Esc asks to go back (Closing).
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

sealed class PageOverview : ScrollableControl
{
	readonly List<(Image Picture, Size Paper)> pages = new();    // (Paper: 1/100 inch)
	readonly Dictionary<int, Bitmap> drawn = new();               // each page at the size shown
	List<int> starts = new();
	float scale = -1;      // page height as a share of the window's (1: a page's height fits), < 0: not set yet
	int current;           // the page with the focus (Enter edits it)
	List<Rectangle> places = new();

	/// Raised with a page (from 0) and where it starts in the text, to edit the document there.
	public event Action<int, int>? PageChosen;
	/// Raised by Esc: back to editing where the caret was.
	public event Action? Closing;

	public static readonly Color Surround = Color.FromArgb(0x80, 0x80, 0x80);
	const int Gap = 16, LabelHeight = 20;

	public PageOverview() {
		SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.Selectable | ControlStyles.ResizeRedraw, true);
		BackColor = Surround;
		AutoScroll = true;
		TabStop = true;
	}

	/// The number of pages shown.
	public int PageCount => pages.Count;

	/// The pages across (for the status bar and tests).
	public int Columns { get; private set; } = 1;

	/// The page with the focus (from 0), which Enter edits; setting it scrolls to it.
	public int CurrentPage {
		get => current;
		set {
			int page = Math.Clamp(value, 0, Math.Max(0, pages.Count - 1));
			ScrollTo(page);
			if (page == current) return;
			current = page;
			Invalidate();
			CurrentChanged?.Invoke();
		}
	}

	/// Raised when another page gets the focus.
	public event Action? CurrentChanged;

	/// Shows document's pages (printed through a preview, as Print Preview prints them); starts: where each page starts in
	/// the text. caretPage: the page the caret is on, which gets the focus and is scrolled to.
	public void Show(PrintDocument document, List<int> starts, int caretPage) {
		Clear();
		var controller = new PreviewPrintController { UseAntiAlias = true };
		var old = document.PrintController;
		document.PrintController = controller;
		try { document.Print(); } finally { document.PrintController = old; }
		foreach (var info in controller.GetPreviewPageInfo()) pages.Add((info.Image, info.PhysicalSize));
		this.starts = starts;
		current = Math.Clamp(caretPage, 0, Math.Max(0, pages.Count - 1));
		Arrange();
		ScrollTo(current);
		Invalidate();
		CurrentChanged?.Invoke();
	}

	/// Shows the pages larger (factor > 1) or smaller; 0: back to a page's height fitting the window.
	public void Zoom(float factor) {
		scale = factor == 0 ? 1 : Math.Clamp((scale < 0 ? 1 : scale) * factor, 0.1f, 4f);
		Arrange();
		ScrollTo(current);
		Invalidate();
	}

	public void Clear() {
		foreach (var (picture, _) in pages) picture.Dispose();
		pages.Clear();
		ClearDrawn();
		places.Clear();
	}

	void ClearDrawn() {
		foreach (var bitmap in drawn.Values) bitmap.Dispose();
		drawn.Clear();
	}

	protected override void Dispose(bool disposing) {
		if (disposing) Clear();
		base.Dispose(disposing);
	}

	// Each page's place: the pages one size (a page's height a share of the window's), in rows of as many as fit across,
	// centred; the rows scroll.
	void Arrange() {
		ClearDrawn();
		places = new List<Rectangle>();
		if (pages.Count == 0) { AutoScrollMinSize = Size.Empty; return; }
		int gap = LogicalToDeviceUnits(Gap), label = LogicalToDeviceUnits(LabelHeight);
		int bar = SystemInformation.VerticalScrollBarWidth;
		int width = Math.Max(1, ClientSize.Width + (VerticalScroll.Visible ? bar : 0) - bar), height = Math.Max(1, ClientSize.Height);
		var paper = pages[0].Paper;
		float s = scale < 0 ? 1 : scale;
		int pageHeight = Math.Max(LogicalToDeviceUnits(60), (int)((height - 2 * gap - label) * s));
		int pageWidth = Math.Max(1, (int)(pageHeight * (double)paper.Width / Math.Max(1, paper.Height)));
		Columns = Math.Clamp((width - gap) / (pageWidth + gap), 1, pages.Count);
		int rowWidth = Columns * pageWidth + (Columns - 1) * gap;
		int left = Math.Max(gap, (width - rowWidth) / 2);
		for (int i = 0; i < pages.Count; i++) {
			int row = i / Columns, column = i % Columns;
			// (a page on other paper keeps its own shape, at the row's height)
			var own = pages[i].Paper;
			int w = own == paper ? pageWidth : Math.Max(1, (int)(pageHeight * (double)own.Width / Math.Max(1, own.Height)));
			places.Add(new Rectangle(left + column * (pageWidth + gap), gap + row * (pageHeight + label + gap), Math.Min(w, pageWidth), pageHeight));
		}
		int rows = (pages.Count + Columns - 1) / Columns;
		AutoScrollMinSize = new Size(Math.Max(width, left + rowWidth + gap), gap + rows * (pageHeight + label + gap));
	}

	void ScrollTo(int page) {
		if (page < 0 || page >= places.Count) return;
		var place = places[page];
		int top = -AutoScrollPosition.Y, bottom = top + ClientSize.Height;
		if (place.Top - LogicalToDeviceUnits(Gap) < top || place.Bottom + LogicalToDeviceUnits(LabelHeight) > bottom)
			AutoScrollPosition = new Point(0, Math.Max(0, place.Top - LogicalToDeviceUnits(Gap)));
	}

	protected override void OnResize(EventArgs e) {
		base.OnResize(e);
		if (pages.Count > 0) { Arrange(); Invalidate(); }
	}

	// A page's picture at its size on screen, drawn once (the preview's pictures are at the printer's resolution).
	Bitmap Drawn(int i) {
		var place = places[i];
		if (drawn.TryGetValue(i, out var bitmap) && bitmap.Size == place.Size) return bitmap;
		bitmap?.Dispose();
		bitmap = new Bitmap(Math.Max(1, place.Width), Math.Max(1, place.Height));
		using (var g = Graphics.FromImage(bitmap)) {
			g.Clear(Color.White);
			g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
			g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
			g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
			g.DrawImage(pages[i].Picture, new Rectangle(Point.Empty, place.Size));
		}
		return drawn[i] = bitmap;
	}

	protected override void OnPaint(PaintEventArgs e) {
		var g = e.Graphics;
		g.Clear(BackColor);
		if (pages.Count == 0) return;
		g.TranslateTransform(AutoScrollPosition.X, AutoScrollPosition.Y);
		var visible = e.ClipRectangle;
		visible.Offset(-AutoScrollPosition.X, -AutoScrollPosition.Y);
		int label = LogicalToDeviceUnits(LabelHeight), shadow = LogicalToDeviceUnits(3);
		using var shade = new SolidBrush(Color.FromArgb(0x60, 0x60, 0x60));
		using var focus = new Pen(SystemColors.Highlight, LogicalToDeviceUnits(3));
		using var font = new Font("Segoe UI", 9f);
		for (int i = 0; i < places.Count; i++) {
			var place = places[i];
			if (!Rectangle.Inflate(place, shadow * 2, label).IntersectsWith(visible)) continue;
			g.FillRectangle(shade, new Rectangle(place.X + shadow, place.Y + shadow, place.Width, place.Height));
			g.DrawImage(Drawn(i), place.Location);
			if (i == current && Focused) g.DrawRectangle(focus, Rectangle.Inflate(place, shadow, shadow));
			TextRenderer.DrawText(g, (i + 1).ToString(), font, new Rectangle(place.X, place.Bottom + shadow, place.Width, label - shadow), Color.White,
				TextFormatFlags.HorizontalCenter | TextFormatFlags.Top);
		}
	}

	int PageAt(Point client) {
		var p = new Point(client.X - AutoScrollPosition.X, client.Y - AutoScrollPosition.Y);
		return places.FindIndex(r => r.Contains(p));
	}

	protected override void OnMouseDown(MouseEventArgs e) {
		base.OnMouseDown(e);
		Focus();
		int page = PageAt(e.Location);
		if (page >= 0) CurrentPage = page;
	}

	protected override void OnMouseDoubleClick(MouseEventArgs e) {
		base.OnMouseDoubleClick(e);
		int page = PageAt(e.Location);
		if (page >= 0) Choose(page);
	}

	/// Asks to edit the document at a page's start.
	public void Choose(int page) {
		if (page < 0 || page >= pages.Count) return;
		PageChosen?.Invoke(page, page < starts.Count ? starts[page] : 0);
	}

	protected override void OnMouseWheel(MouseEventArgs e) {
		if ((ModifierKeys & Keys.Control) != 0) { Zoom(e.Delta > 0 ? 1.1f : 1 / 1.1f); return; }
		base.OnMouseWheel(e);
		FollowScroll();
	}

	protected override void OnScroll(ScrollEventArgs se) {
		base.OnScroll(se);
		FollowScroll();
	}

	// Scrolled (the wheel, the scroll bar): the focus goes to the row of pages across the middle of the window (the first
	// in it), unless it is on that row already.
	void FollowScroll() {
		if (places.Count == 0) return;
		int middle = -AutoScrollPosition.Y + ClientSize.Height / 2, gap = LogicalToDeviceUnits(Gap + LabelHeight);
		int page = places.FindIndex(r => middle < r.Bottom + gap);
		if (page < 0) page = places.Count - 1;
		if (current < places.Count && places[current].Top == places[page].Top) return;
		if (page == current) return;
		current = page;
		Invalidate();
		CurrentChanged?.Invoke();
	}

	protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
	protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

	protected override bool IsInputKey(Keys keyData) => (keyData & Keys.KeyCode) is Keys.Left or Keys.Right or Keys.Up or Keys.Down || base.IsInputKey(keyData);

	protected override void OnKeyDown(KeyEventArgs e) {
		base.OnKeyDown(e);
		int next = e.KeyCode switch {
			Keys.Left => current - 1, Keys.Right => current + 1, Keys.Up => current - Columns, Keys.Down => current + Columns,
			Keys.Home => 0, Keys.End => pages.Count - 1, Keys.PageUp => current - Columns, Keys.PageDown => current + Columns, _ => -2
		};
		if (e.KeyCode == Keys.Enter) { Choose(current); e.Handled = true; return; }
		if (e.KeyCode == Keys.Escape) { Closing?.Invoke(); e.Handled = true; return; }
		if (next == -2 || pages.Count == 0) return;
		CurrentPage = next;
		e.Handled = true;
	}
}

// Print Preview: the printed pages in a window of their own, fitted to the screen's working area (above the taskbar).
// The toolbar: Print, the first, previous, next and last page with "Page 2 of 6", zoom out, in and whole page, Close.
// The keys are the page view's (arrows, Page Up / Down, Home / End, Ctrl+wheel), Ctrl+P prints and Esc closes;
// double-clicking a page closes the preview to edit the document there (ChosenStart).
sealed class PagePreviewForm : Form
{
	public readonly PageOverview Pages = new() { Dock = DockStyle.Fill };
	readonly ToolStripLabel position = new() { AutoSize = false, TextAlign = ContentAlignment.MiddleCenter };
	readonly ToolStripButton first, previous, next, last;

	/// Raised by Print (and Ctrl+P): the caller prints (its Print dialog), and the preview stays open.
	public event Action? PrintRequested;

	/// Where the page double-clicked starts in the text (null: closed without one).
	public int? ChosenStart { get; private set; }

	public PagePreviewForm(string title, Control owner) {
		Text = "Print Preview - " + title;
		ShowIcon = false;
		ShowInTaskbar = false;
		KeyPreview = true;
		StartPosition = FormStartPosition.Manual;
		Font = SystemFonts.MessageBoxFont ?? Font;
		// the screen's working area (without the taskbar) the owner is on: nearly its whole height, a page's width
		var area = Screen.FromControl(owner).WorkingArea;
		int height = (int)(area.Height * 0.94), width = Math.Min(area.Width, Math.Max(owner.LogicalToDeviceUnits(640), (int)(height * 0.8)));
		Bounds = new Rectangle(area.Left + (area.Width - width) / 2, area.Top + (area.Height - height) / 2, width, height);
		MinimumSize = new Size(owner.LogicalToDeviceUnits(420), owner.LogicalToDeviceUnits(320));

		var bar = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden, Padding = new Padding(owner.LogicalToDeviceUnits(4), 2, 4, 2), Stretch = true };
		ToolStripButton Button(string text, string tip, Action click) {
			var b = new ToolStripButton(text) { ToolTipText = tip, DisplayStyle = ToolStripItemDisplayStyle.Text };
			b.Click += (s, e) => { click(); Pages.Focus(); };
			bar.Items.Add(b);
			return b;
		}
		Button("&Print...", "Print the document (Ctrl+P)", () => PrintRequested?.Invoke());
		bar.Items.Add(new ToolStripSeparator());
		first = Button("|<", "First page (Home)", () => Pages.CurrentPage = 0);
		previous = Button("<", "Previous page (Left)", () => Pages.CurrentPage -= 1);
		position.Width = owner.LogicalToDeviceUnits(96);
		bar.Items.Add(position);
		next = Button(">", "Next page (Right)", () => Pages.CurrentPage += 1);
		last = Button(">|", "Last page (End)", () => Pages.CurrentPage = Pages.PageCount - 1);
		bar.Items.Add(new ToolStripSeparator());
		Button("Zoom &Out", "Smaller pages, more across (Ctrl+wheel)", () => Pages.Zoom(1 / 1.25f));
		Button("Zoom &In", "Larger pages (Ctrl+wheel)", () => Pages.Zoom(1.25f));
		Button("&Whole Page", "A page's height fits the window", () => Pages.Zoom(0));
		bar.Items.Add(new ToolStripSeparator());
		Button("&Close", "Close the preview (Esc)", Close);
		bar.Items.Add(new ToolStripLabel("Double-click a page to edit it there") { Alignment = ToolStripItemAlignment.Right, ForeColor = SystemColors.GrayText });

		Controls.Add(Pages);
		Controls.Add(bar);
		Pages.CurrentChanged += ShowPosition;
		Pages.Closing += Close;
		Pages.PageChosen += (page, start) => { ChosenStart = start; DialogResult = DialogResult.OK; Close(); };
		Shown += (s, e) => Pages.Focus();
	}

	void ShowPosition() {
		int count = Pages.PageCount, page = Pages.CurrentPage;
		position.Text = count == 0 ? "No pages" : $"Page {page + 1} of {count}";
		first.Enabled = previous.Enabled = page > 0;
		next.Enabled = last.Enabled = page < count - 1;
	}

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
		if (keyData == Keys.Escape) { Close(); return true; }
		if (keyData == (Keys.Control | Keys.P)) { PrintRequested?.Invoke(); return true; }
		return base.ProcessCmdKey(ref msg, keyData);
	}
}
