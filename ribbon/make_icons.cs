// Draws AhkWord's ribbon icons as 32-bit BMPs (the format the Windows Ribbon compiler, uicc.exe, takes), in the
// sizes the ribbon picks from by screen DPI: 16 20 24 32 for small images, 32 40 48 64 for large ones.
// The icons are Microsoft's Fluent UI System Icons (MIT license, https://github.com/microsoft/fluentui-system-icons):
// the "Regular" font has each icon drawn for several pixel sizes, and the one drawn for the size wanted (or for an exact
// fraction of it) is used, so lines stay sharp. A few are drawn here. Run by build_ribbon.ahk, which downloads the font.
// args: [0] output folder, [1] folder with FluentSystemIcons-Regular.ttf and .json. Also writes _sheet.png, all icons
// at 32 pixels with their names, to look them over.
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Text.Json;

string outDir = args[0], fontDir = args[1];
Directory.CreateDirectory(outDir);
Icons.Load(Path.Combine(fontDir, "FluentSystemIcons-Regular.ttf"), Path.Combine(fontDir, "FluentSystemIcons-Regular.json"));

// ribbon image name -> Fluent icon name (without ic_fluent_ and _<size>_regular)
var fluent = new Dictionary<string, string> {
	["new"] = "document_add", ["open"] = "folder_open", ["save"] = "save", ["saveas"] = "save_edit", ["print"] = "print",
	["printpreview"] = "document_search", ["pagesetup"] = "document_settings", ["exit"] = "dismiss",
	["export"] = "arrow_export_ltr", ["exportpdf"] = "document_pdf", ["about"] = "info",
	["undo"] = "arrow_undo", ["redo"] = "arrow_redo", ["paste"] = "clipboard_paste", ["cut"] = "cut", ["copy"] = "copy",
	["selectall"] = "select_all_on", ["find"] = "search", ["replace"] = "arrow_swap",
	["font"] = "text_font", ["changecase"] = "text_change_case", ["underline"] = "text_underline",
	["bullets"] = "text_bullet_list_ltr", ["numbering"] = "text_number_list_ltr",
	["outdent"] = "text_indent_decrease_ltr", ["indent"] = "text_indent_increase_ltr", ["listoptions"] = "options",
	["alignleft"] = "text_align_left", ["aligncenter"] = "text_align_center", ["alignright"] = "text_align_right", ["justify"] = "text_align_justify",
	["pagebreak"] = "document_page_break", ["sectionbreak"] = "document_split_hint", ["table"] = "table", ["picture"] = "image", ["thumbnail"] = "image_multiple",
	["link"] = "link", ["crossreference"] = "number_symbol_square", ["updatereferences"] = "arrow_sync", ["footnote"] = "text_footnote", ["tableofcontents"] = "document_bullet_list", ["snippets"] = "text_expand", ["horizontalline"] = "divider_short", ["datetime"] = "calendar_clock", ["headerfooter"] = "document_header_footer",
	["nextspelling"] = "text_proofing_tools", ["dictionary"] = "book_add",
	["structured"] = "text_header_1", ["freeform"] = "text_font_size",
	["zoomin"] = "zoom_in", ["zoomout"] = "zoom_out", ["zoom100"] = "zoom_fit", ["html"] = "code", ["markdown"] = "markdown",
	["viewrich"] = "document_text", ["sidebyside"] = "layout_column_two",
	["tablesettings"] = "table_settings", ["linespacing"] = "text_line_spacing", ["pagewidth"] = "document_width", ["printlayout"] = "document_one_page", ["onepage"] = "page_fit", ["multiplepages"] = "document_multiple", ["close"] = "document_dismiss", ["newtemplate"] = "document_add", ["startpage"] = "home", ["newwindow"] = "window_new", ["stylesmenu"] = "text_edit_style", ["modifystyle"] = "edit", ["newstyle"] = "add", ["updatestyle"] = "arrow_sync", ["deletestyle"] = "delete", ["savestyles"] = "save", ["clearformatting"] = "text_clear_formatting", ["paragraphsettings"] = "text_paragraph", ["applystyles"] = "paint_brush_arrow_down", ["openwindow"] = "folder_open", ["inserttemplate"] = "document_text", ["fillfields"] = "textbox", ["borderall"] = "border_all", ["borderoutside"] = "border_outside",
	["borderinside"] = "border_inside", ["bordernone"] = "border_none", ["paintbucket"] = "paint_bucket",
	["macros"] = "play", ["record"] = "record", ["hotkeys"] = "keyboard", ["runscript"] = "script", ["macrosfolder"] = "folder_open", ["eventhooks"] = "flash", ["commandpalette"] = "search", ["api"] = "plug_connected", ["extensions"] = "puzzle_piece",
	["insertabove"] = "table_stack_above", ["insertbelow"] = "table_stack_below", ["insertleft"] = "table_stack_left", ["insertright"] = "table_stack_right",
	["deleterow"] = "table_delete_row", ["deletecolumn"] = "table_delete_column", ["deletetable"] = "table_dismiss",
};
var drawn = new Dictionary<string, Action<Graphics, float>> {
	["legal"] = Icons.Legal,
	["linenumbers"] = Icons.LineNumbers,
	["exportword"] = (g, s) => Icons.WithLetter(g, s, "W", Color.FromArgb(0x18, 0x5A, 0xBD)),
	["exportodt"] = (g, s) => Icons.WithLetter(g, s, "O", Color.FromArgb(0x10, 0x7C, 0x41)),
};

var all = fluent.Keys.Concat(drawn.Keys).OrderBy(n => n).ToList();
int count = 0;
foreach (int size in new[] { 16, 20, 24, 32, 40, 48, 64 })
	foreach (var name in all) {
		Icons.Save(Path.Combine(outDir, $"{name}_{size}.bmp"), size, (g, s) => {
			if (fluent.TryGetValue(name, out var icon)) Icons.Fluent(g, s, icon);
			else drawn[name](g, s);
		});
		count++;
	}
Icons.Sheet(Path.Combine(outDir, "_sheet.png"), all, name => (g, s) => {
	if (fluent.TryGetValue(name, out var icon)) Icons.Fluent(g, s, icon);
	else drawn[name](g, s);
});
AHK.Return($"{count} icons in {outDir} (Fluent UI System Icons)");

static class Icons
{
	static PrivateFontCollection fonts = null!;
	static FontFamily family = null!;
	static Dictionary<string, int> codes = null!;
	static readonly Color Ink = Color.FromArgb(0x30, 0x30, 0x30), Accent = Color.FromArgb(0x00, 0x67, 0xC0);

	public static void Load(string fontFile, string namesFile) {
		fonts = new PrivateFontCollection();
		fonts.AddFontFile(fontFile);
		family = fonts.Families[0];
		codes = JsonSerializer.Deserialize<Dictionary<string, int>>(File.ReadAllText(namesFile))!;
	}

	public static void Save(string path, int size, Action<Graphics, float> draw) {
		using var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
		using (var g = Graphics.FromImage(bmp)) {
			g.Clear(Color.Transparent);
			g.SmoothingMode = SmoothingMode.AntiAlias;
			g.PixelOffsetMode = PixelOffsetMode.HighQuality;
			g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
			draw(g, size);
		}
		WriteBmp(bmp, path);
	}

	// All icons at 32 pixels with their names, to look them over.
	public static void Sheet(string path, List<string> names, Func<string, Action<Graphics, float>> draw) {
		int cols = 8, cell = 110;
		using var bmp = new Bitmap(cols * cell, (names.Count + cols - 1) / cols * 80);
		using var g = Graphics.FromImage(bmp);
		g.Clear(Color.White);
		g.SmoothingMode = SmoothingMode.AntiAlias;
		g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
		using var label = new Font("Segoe UI", 10, GraphicsUnit.Pixel);
		using var format = new StringFormat { Alignment = StringAlignment.Center };
		for (int i = 0; i < names.Count; i++) {
			var state = g.Save();
			g.TranslateTransform(i % cols * cell + (cell - 32) / 2, i / cols * 80 + 8);
			draw(names[i])(g, 32);
			g.Restore(state);
			g.DrawString(names[i], label, Brushes.Gray, new RectangleF(i % cols * cell, i / cols * 80 + 46, cell, 20), format);
		}
		bmp.Save(path, ImageFormat.Png);
	}

	// A 32 bpp BI_RGB bitmap with straight (not premultiplied) alpha, bottom-up: what uicc.exe expects.
	static void WriteBmp(Bitmap bmp, string path) {
		int w = bmp.Width, h = bmp.Height, stride = w * 4;
		var data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
		var pixels = new byte[stride * h];
		for (int y = 0; y < h; y++)
			System.Runtime.InteropServices.Marshal.Copy(data.Scan0 + y * data.Stride, pixels, (h - 1 - y) * stride, stride);
		bmp.UnlockBits(data);
		using var file = new BinaryWriter(File.Create(path));
		file.Write((ushort)0x4D42); file.Write(14 + 40 + pixels.Length); file.Write(0); file.Write(14 + 40);
		file.Write(40); file.Write(w); file.Write(h); file.Write((ushort)1); file.Write((ushort)32); file.Write(0);
		file.Write(pixels.Length); file.Write(2835); file.Write(2835); file.Write(0); file.Write(0);
		file.Write(pixels);
	}

	// The design size to draw at s pixels: s itself, else the largest that divides s (drawn 2x, 3x ... so its lines
	// stay on whole pixels), else the smallest larger one, else the largest.
	static int DesignSize(string icon, int s) {
		var sizes = new[] { 12, 16, 20, 24, 28, 32, 48 }.Where(d => codes.ContainsKey($"ic_fluent_{icon}_{d}_regular")).ToList();
		if (sizes.Count == 0) throw new KeyNotFoundException($"Fluent icon \"{icon}\" not found.");
		if (sizes.Contains(s)) return s;
		int divisor = sizes.Where(d => s % d == 0).DefaultIfEmpty(0).Max();
		if (divisor > 0) return divisor;
		return sizes.Where(d => d > s).DefaultIfEmpty(sizes.Max()).Min();
	}

	// A Fluent icon filling the box (x, y, box x box), drawn as a path: its em box is the icon's square, with the baseline
	// at the bottom (the font's ascent is its em height).
	public static void Fluent(Graphics g, float s, string icon, Color? color = null, RectangleF? box = null) {
		var area = box ?? new RectangleF(0, 0, s, s);
		int design = DesignSize(icon, (int)Math.Round(area.Width));
		using var path = new GraphicsPath();
		path.AddString(char.ConvertFromUtf32(codes[$"ic_fluent_{icon}_{design}_regular"]), family, 0, area.Width, new PointF(0, 0), StringFormat.GenericTypographic);
		float baseline = area.Width * family.GetCellAscent(FontStyle.Regular) / family.GetEmHeight(FontStyle.Regular);
		using var move = new Matrix();
		move.Translate(area.X, area.Y + area.Width - baseline);
		path.Transform(move);
		using var brush = new SolidBrush(color ?? Ink);
		g.FillPath(brush, path);
	}

	static void Text(Graphics g, string text, float size, Color color, RectangleF box) {
		using var font = new Font("Segoe UI", size, FontStyle.Bold, GraphicsUnit.Pixel);
		using var brush = new SolidBrush(color);
		using var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
		g.DrawString(text, font, brush, box, format);
	}

	// A document with a letter on a colored square in its lower right corner (W = Word, O = OpenDocument).
	public static void WithLetter(Graphics g, float s, string letter, Color color) {
		Fluent(g, s, "document", null, new RectangleF(0, 0, s, s));
		var box = new RectangleF(s * 0.46f, s * 0.46f, s * 0.54f, s * 0.54f);
		using var back = new SolidBrush(color);
		g.FillRectangle(back, box);
		Text(g, letter, s * 0.42f, Color.White, box);
	}

	// "1.1" above two indented lines: legal numbering (Fluent has no such icon).
	public static void Legal(Graphics g, float s) {
		Text(g, "1.1", s * 0.42f, Accent, new RectangleF(0, 0, s, s * 0.5f));
		using var pen = new Pen(Ink, Math.Max(1f, s / 16f));
		g.DrawLine(pen, s * 0.1f, s * 0.66f, s * 0.9f, s * 0.66f);
		g.DrawLine(pen, s * 0.3f, s * 0.86f, s * 0.9f, s * 0.86f);
	}

	// Lines of text with their numbers in the margin, left of a thin rule: line numbers (Fluent's numbered list is the
	// Numbering button's).
	public static void LineNumbers(Graphics g, float s) {
		using var pen = new Pen(Ink, Math.Max(1f, s / 16f));
		using var rule = new Pen(Color.FromArgb(0xA0, 0xA0, 0xA0), Math.Max(1f, s / 32f));
		g.DrawLine(rule, s * 0.36f, s * 0.08f, s * 0.36f, s * 0.92f);
		for (int i = 0; i < 3; i++) {
			float y = s * (0.2f + i * 0.3f);
			Text(g, (i + 1).ToString(), s * 0.34f, Accent, new RectangleF(0, y - s * 0.17f, s * 0.34f, s * 0.34f));
			g.DrawLine(pen, s * 0.48f, y, s * (i == 2 ? 0.74f : 0.94f), y);
		}
	}
}
