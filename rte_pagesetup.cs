// Iteration 82: AhkWord's own Page Setup (Layout > Page Setup), in place of Windows' dialog: the document's margins,
// orientation and paper, and the header's and footer's distances from the page's edges and the different first page
// (as Word's Page Setup has them; the Header and Footer dialog has them too), with a small picture of the page. OK sets
// them for this document (saved with it: RTF, Word); "Make these the default" for new documents too. The paper is the
// document's own now: the file's (RTF's \paperw \paperh, Word's w:pgSz) or Page Setup's, else the default printer's;
// printing asks the printer for that paper (Papers.ForPrinter).

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

/// Paper sizes, in 1/100 inch, width and height upright.
static class Papers
{
	public sealed record Paper(int Width, int Height);

	// (their names in Names, in this order)
	public static readonly Paper[] All = { new(850, 1100), new(850, 1400), new(725, 1050), new(1100, 1700), new(1169, 1654), new(827, 1169), new(583, 827), new(717, 1012) };

	static string[] Names() => new[] {
		T("paper.letter", "Letter"), T("paper.legal", "Legal"), T("paper.executive", "Executive"), T("paper.tabloid", "Tabloid"),
		T("paper.a3", "A3"), T("paper.a4", "A4"), T("paper.a5", "A5"), T("paper.b5", "B5 (JIS)")
	};

	/// A paper's name in the interface's language.
	public static string NameOf(Paper paper) => Names()[Array.IndexOf(All, paper)];

	/// Letter, or A4 where the region measures in centimetres.
	public static int[] RegionDefault => RegionInfo.CurrentRegion.IsMetric ? new[] { 827, 1169 } : new[] { 850, 1100 };

	/// The named paper of this size (within 1/25 inch, upright or across), or null.
	public static Paper? Find(int width, int height) {
		var (w, h) = (Math.Min(width, height), Math.Max(width, height));
		return All.FirstOrDefault(p => Math.Abs(p.Width - w) <= 4 && Math.Abs(p.Height - h) <= 4);
	}

	/// A paper's width and height upright, in 1/100 inch (twips: given in twips), or null when either is missing or not
	/// above 0.
	public static int[]? Upright(int? width, int? height, bool twips = true) {
		if (width is not int w || height is not int h || w <= 0 || h <= 0) return null;
		if (twips) (w, h) = ((int)Math.Round(w / 14.4), (int)Math.Round(h / 14.4));
		return new[] { Math.Min(w, h), Math.Max(w, h) };
	}

	/// The printer's own paper of this size (so that it takes the right tray and form), else a custom size.
	public static PaperSize ForPrinter(PrinterSettings printer, int width, int height) {
		try {
			foreach (PaperSize size in printer.PaperSizes)
				if (size.Kind != PaperKind.Custom && Math.Abs(Math.Min(size.Width, size.Height) - Math.Min(width, height)) <= 4
					&& Math.Abs(Math.Max(size.Width, size.Height) - Math.Max(width, height)) <= 4) return size;
		} catch (Exception) {
			// (no printer, or one that can't be asked: the size as it is)
		}
		return new PaperSize("Custom", width, height);
	}
}

/// What Page Setup gives: margins (1/100 inch: left, right, top, bottom), orientation, paper (1/100 inch, upright), the
/// header's and footer's distances from the page's edges (1/100 inch), a different first page, and whether these become
/// the default for new documents.
sealed record PageSetupChoice(int[] Margins, bool Landscape, int[] Paper, int HeaderDistance, int FooterDistance, bool DifferentFirst, bool MakeDefault);

sealed class PageSetupForm : Form
{
	readonly bool metric = RegionInfo.CurrentRegion.IsMetric;
	decimal ToUnits(int hundredths) => metric ? Math.Round(hundredths * 0.0254m, 2) : Math.Round(hundredths / 100m, 2);
	int FromUnits(decimal value) => (int)Math.Round(metric ? value / 0.0254m : value * 100);
	string Unit => metric ? "cm" : T("dlg.pageSetup.in", "in");

	NumericUpDown Box(string name) => new() { DecimalPlaces = 2, Increment = 0.1m, Width = LogicalToDeviceUnits(70), AccessibleName = name, Margin = new Padding(3, 3, 3, 3) };

	readonly NumericUpDown top, bottom, left, right, width, height, headerFrom, footerFrom;
	readonly RadioButton portrait = new() { Text = T("dlg.pageSetup.portrait", "Portrait"), AutoSize = true };
	readonly RadioButton landscape = new() { Text = T("dlg.pageSetup.landscape", "Landscape"), AutoSize = true };
	readonly ComboBox paper = new() { DropDownStyle = ComboBoxStyle.DropDownList, AccessibleName = T("dlg.pageSetup.paperName", "Paper size") };
	readonly CheckBox differentFirst = new() { Text = T("dlg.pageSetup.differentFirst", "Different first page (its own header and footer)"), AutoSize = true, Margin = new Padding(3, 6, 3, 3) };
	readonly CheckBox makeDefault = new() { Text = T("dlg.pageSetup.makeDefault", "Make these the default for new documents"), AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
	readonly PagePicture picture = new();
	bool filling;    // (the boxes being filled in by the dialog itself, not typed in)

	/// The page setup chosen (after OK).
	public PageSetupChoice Result { get; private set; } = null!;

	/// margins: 1/100 inch, left, right, top, bottom; paper: width and height upright; distances: 1/100 inch from the edges.
	public PageSetupForm(int[] margins, bool isLandscape, int[] paperSize, int headerDistance, int footerDistance, bool first) {
		Dialogs.Setup(this, T("dlg.pageSetup.title", "Page Setup"));
		top = Box(T("dlg.pageSetup.topName", "Top margin"));
		bottom = Box(T("dlg.pageSetup.bottomName", "Bottom margin"));
		left = Box(T("dlg.pageSetup.leftName", "Left margin"));
		right = Box(T("dlg.pageSetup.rightName", "Right margin"));
		width = Box(T("dlg.pageSetup.widthName", "Paper width"));
		height = Box(T("dlg.pageSetup.heightName", "Paper height"));
		headerFrom = Box(T("dlg.pageSetup.headerFromName", "Header from top"));
		footerFrom = Box(T("dlg.pageSetup.footerFromName", "Footer from bottom"));
		foreach (var box in new[] { top, bottom, left, right, headerFrom, footerFrom }) box.Maximum = ToUnits(1000);
		foreach (var box in new[] { width, height }) { box.Minimum = ToUnits(100); box.Maximum = ToUnits(4800); }

		Label Text(string text) => new() { Text = text, AutoSize = true, Margin = new Padding(3, 7, 3, 0) };
		TableLayoutPanel Grid() => new() { ColumnCount = 4, AutoSize = true, Margin = new Padding(0) };
		GroupBox Group(string title, Control content) {
			var group = new GroupBox { Text = title, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, Padding = new Padding(LogicalToDeviceUnits(6)) };
			content.Dock = DockStyle.Fill;
			group.Controls.Add(content);
			return group;
		}

		// margins: top and bottom, left and right, side by side as Word's
		var marginGrid = Grid();
		marginGrid.Controls.Add(Text(T("dlg.pageSetup.top", "Top:")), 0, 0);
		marginGrid.Controls.Add(top, 1, 0);
		marginGrid.Controls.Add(Text(T("dlg.pageSetup.bottom", "Bottom:")), 2, 0);
		marginGrid.Controls.Add(bottom, 3, 0);
		marginGrid.Controls.Add(Text(T("dlg.pageSetup.left", "Left:")), 0, 1);
		marginGrid.Controls.Add(left, 1, 1);
		marginGrid.Controls.Add(Text(T("dlg.pageSetup.right", "Right:")), 2, 1);
		marginGrid.Controls.Add(right, 3, 1);

		var orientationRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
		orientationRow.Controls.Add(portrait);
		orientationRow.Controls.Add(landscape);

		// the paper: a named size, or Custom with its own width and height (as the page is turned: Word's way)
		foreach (var p in Papers.All) paper.Items.Add(T("dlg.pageSetup.paperItem", "{0} ({1} × {2} {3})", Papers.NameOf(p), ToUnits(p.Width), ToUnits(p.Height), Unit));
		paper.Items.Add(T("dlg.pageSetup.custom", "Custom size"));
		Dialogs.FitItems(paper);
		var paperGrid = Grid();
		paperGrid.Controls.Add(paper, 0, 0);
		paperGrid.SetColumnSpan(paper, 4);
		paperGrid.Controls.Add(Text(T("dlg.pageSetup.width", "Width:")), 0, 1);
		paperGrid.Controls.Add(width, 1, 1);
		paperGrid.Controls.Add(Text(T("dlg.pageSetup.height", "Height:")), 2, 1);
		paperGrid.Controls.Add(height, 3, 1);

		var bandGrid = Grid();
		bandGrid.Controls.Add(Text(T("dlg.pageSetup.headerFrom", "Header from top:")), 0, 0);
		bandGrid.Controls.Add(headerFrom, 1, 0);
		bandGrid.Controls.Add(Text(T("dlg.pageSetup.footerFrom", "Footer from bottom:")), 2, 0);
		bandGrid.Controls.Add(footerFrom, 3, 0);
		bandGrid.Controls.Add(differentFirst, 0, 1);
		bandGrid.SetColumnSpan(differentFirst, 4);

		var groups = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Margin = new Padding(0) };
		var marginGroup = Group(T("dlg.pageSetup.margins", "Margins ({0})", Unit), marginGrid);
		var orientationGroup = Group(T("dlg.pageSetup.orientation", "Orientation"), orientationRow);
		var paperGroup = Group(T("dlg.pageSetup.paper", "Paper ({0})", Unit), paperGrid);
		var bandGroup = Group(T("dlg.pageSetup.headerFooter", "Header and footer ({0})", Unit), bandGrid);
		groups.Controls.AddRange(new Control[] { marginGroup, orientationGroup, paperGroup, bandGroup });
		groups.Layout += (s, e) => {    // (the groups as wide as the widest)
			int widest = groups.Controls.Cast<Control>().Max(c => c.PreferredSize.Width);
			foreach (Control c in groups.Controls) c.MinimumSize = new Size(widest, 0);
		};

		picture.Size = new Size(LogicalToDeviceUnits(150), LogicalToDeviceUnits(190));
		picture.Margin = new Padding(LogicalToDeviceUnits(12), LogicalToDeviceUnits(8), 0, 0);
		var note = new Label { Text = T("dlg.pageSetup.note", "For this document, saved with it. Sections with margins of their own keep them (Header and Footer)."), AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(150), 0), Margin = new Padding(LogicalToDeviceUnits(12), LogicalToDeviceUnits(8), 0, 0), ForeColor = SystemColors.GrayText };
		var side = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Margin = new Padding(0) };
		side.Controls.Add(picture);
		side.Controls.Add(note);

		var main = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
		main.Controls.Add(groups, 0, 0);
		main.Controls.Add(side, 1, 0);
		main.Controls.Add(makeDefault, 0, 1);
		main.SetColumnSpan(makeDefault, 2);
		var ok = Dialogs.Ok(this, closes: false);
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 8, ok, cancel);
		main.Controls.Add(buttons, 0, 2);
		main.SetColumnSpan(buttons, 2);
		Controls.Add(main);
		AcceptButton = ok;
		CancelButton = cancel;

		// the values
		filling = true;
		top.Value = Clamp(top, margins[2]);
		bottom.Value = Clamp(bottom, margins[3]);
		left.Value = Clamp(left, margins[0]);
		right.Value = Clamp(right, margins[1]);
		(isLandscape ? landscape : portrait).Checked = true;
		var (w, h) = isLandscape ? (paperSize[1], paperSize[0]) : (paperSize[0], paperSize[1]);
		width.Value = Clamp(width, w);
		height.Value = Clamp(height, h);
		headerFrom.Value = Clamp(headerFrom, headerDistance);
		footerFrom.Value = Clamp(footerFrom, footerDistance);
		differentFirst.Checked = first;
		ChooseNamedPaper();
		filling = false;

		paper.SelectedIndexChanged += (s, e) => {
			if (filling || paper.SelectedIndex < 0 || paper.SelectedIndex >= Papers.All.Length) { Picture(); return; }
			var p = Papers.All[paper.SelectedIndex];
			filling = true;
			var (pw, ph) = landscape.Checked ? (p.Height, p.Width) : (p.Width, p.Height);
			width.Value = Clamp(width, pw);
			height.Value = Clamp(height, ph);
			filling = false;
			Picture();
		};
		foreach (var box in new[] { width, height })
			box.ValueChanged += (s, e) => { if (filling) return; filling = true; ChooseNamedPaper(); filling = false; Picture(); };
		// (turning the page: its width and height change places)
		landscape.CheckedChanged += (s, e) => {
			if (filling) return;
			bool across = landscape.Checked;
			if (across == width.Value > height.Value || width.Value == height.Value) { Picture(); return; }
			filling = true;
			(width.Value, height.Value) = (height.Value, width.Value);
			filling = false;
			Picture();
		};
		foreach (var box in new[] { top, bottom, left, right, headerFrom, footerFrom }) box.ValueChanged += (s, e) => Picture();
		differentFirst.CheckedChanged += (s, e) => Picture();
		ok.Click += (s, e) => { if (Accept()) DialogResult = DialogResult.OK; };
		Picture();
	}

	decimal Clamp(NumericUpDown box, int hundredths) => Math.Clamp(ToUnits(hundredths), box.Minimum, box.Maximum);

	// the paper's name for the size in the boxes (Custom size when none has it)
	void ChooseNamedPaper() {
		var named = Papers.Find(FromUnits(width.Value), FromUnits(height.Value));
		paper.SelectedIndex = named == null ? Papers.All.Length : Array.IndexOf(Papers.All, named);
	}

	void Picture() => picture.Show(FromUnits(width.Value), FromUnits(height.Value), FromUnits(left.Value), FromUnits(right.Value), FromUnits(top.Value), FromUnits(bottom.Value), FromUnits(headerFrom.Value), FromUnits(footerFrom.Value));

	// OK: the text keeps at least half an inch across and down
	bool Accept() {
		int w = FromUnits(width.Value), h = FromUnits(height.Value);
		int across = w - FromUnits(left.Value) - FromUnits(right.Value), down = h - FromUnits(top.Value) - FromUnits(bottom.Value);
		if (across < 50 || down < 50) {
			MessageBox.Show(this, T("msg.pageSetupTooSmall", "The margins leave too little room for the text: at least {0} {1} across and down.", ToUnits(50), Unit), T("dlg.pageSetup.title", "Page Setup"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
			(across < 50 ? left : top).Focus();
			return false;
		}
		bool turned = landscape.Checked;
		int[] upright = turned ? new[] { h, w } : new[] { w, h };
		Result = new PageSetupChoice(new[] { FromUnits(left.Value), FromUnits(right.Value), FromUnits(top.Value), FromUnits(bottom.Value) }, turned, upright,
			FromUnits(headerFrom.Value), FromUnits(footerFrom.Value), differentFirst.Checked, makeDefault.Checked);
		return true;
	}

	/// The page as the dialog has it: the paper, its text area (dashed) and where the header and footer are.
	sealed class PagePicture : Control
	{
		int paperWidth = 850, paperHeight = 1100, left, right, top, bottom, header, footer;

		public PagePicture() {
			SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
			TabStop = false;
			AccessibleRole = AccessibleRole.Graphic;
			AccessibleName = T("dlg.pageSetup.pictureName", "The page");
		}

		public void Show(int width, int height, int left, int right, int top, int bottom, int header, int footer) {
			(paperWidth, paperHeight, this.left, this.right, this.top, this.bottom, this.header, this.footer) = (Math.Max(1, width), Math.Max(1, height), left, right, top, bottom, header, footer);
			Invalidate();
		}

		protected override void OnPaint(PaintEventArgs e) {
			var g = e.Graphics;
			g.Clear(BackColor);
			g.SmoothingMode = SmoothingMode.AntiAlias;
			int room = LogicalToDeviceUnits(8);
			float scale = Math.Min((Width - 2f * room) / paperWidth, (Height - 2f * room) / paperHeight);
			float w = paperWidth * scale, h = paperHeight * scale, x = (Width - w) / 2, y = (Height - h) / 2;
			int shadow = LogicalToDeviceUnits(3);
			using (var dark = new SolidBrush(Color.FromArgb(70, 0, 0, 0))) g.FillRectangle(dark, x + shadow, y + shadow, w, h);
			g.FillRectangle(Brushes.White, x, y, w, h);
			g.DrawRectangle(Pens.Gray, x, y, w, h);
			var text = RectangleF.FromLTRB(x + left * scale, y + top * scale, x + w - right * scale, y + h - bottom * scale);
			if (text.Width > 0 && text.Height > 0) {
				// the text's lines, in grey, and the text area's edge
				using var lines = new Pen(Color.FromArgb(0xC8, 0xC8, 0xC8), Math.Max(1, LogicalToDeviceUnits(2)));
				float step = Math.Max(LogicalToDeviceUnits(4), 0.18f * 100 * scale);
				for (float at = text.Top + step / 2; at < text.Bottom; at += step) g.DrawLine(lines, text.Left + 1, at, text.Right - 1, at);
				using var edge = new Pen(Color.FromArgb(0x60, 0x80, 0xC0)) { DashStyle = DashStyle.Dash };
				g.DrawRectangle(edge, text.X, text.Y, text.Width, text.Height);
			}
			using var band = new Pen(Color.FromArgb(0xE0, 0x90, 0x40), Math.Max(1, LogicalToDeviceUnits(2)));
			float headerY = y + header * scale, footerY = y + h - footer * scale;
			if (header < paperHeight / 2) g.DrawLine(band, text.Left, headerY, text.Left + Math.Max(4, text.Width / 3), headerY);
			if (footer < paperHeight / 2) g.DrawLine(band, text.Right - Math.Max(4, text.Width / 3), footerY, text.Right, footerY);
		}
	}
}
