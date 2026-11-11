// Draws AhkWord's ribbon icons as 32-bit BMPs (the format the Windows Ribbon compiler, uicc.exe, takes), in the
// sizes the ribbon picks from by screen DPI: 16 20 24 32 for small images, 32 40 48 64 for large ones.
// The icons are Microsoft's Fluent UI System Icons (MIT license, https://github.com/microsoft/fluentui-system-icons):
// the "Regular" font has each icon drawn for several pixel sizes, and the one drawn for the size wanted (or for an exact
// fraction of it) is used, so lines stay sharp. Most take a colour by kind of command, on their inside (the "Filled"
// font's shape, tinted, under the outline). A few are drawn here. Run by build_ribbon.ahk, which downloads the fonts.
// args: [0] output folder, [1] folder with FluentSystemIcons-Regular and -Filled .ttf and .json. Also writes _sheet.png, all icons
// at 32 pixels with their names, to look them over.
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Text.Json;

string outDir = args[0], fontDir = args[1];
if (args.Length > 2 && args[2] == "light") Icons.UseLightInk();    // (for the dark ribbon: RteRibbonDark.ribbon)
Directory.CreateDirectory(outDir);
Icons.Load(fontDir);

// ribbon image name -> Fluent icon name (without ic_fluent_ and _<size>_regular)
var fluent = new Dictionary<string, string> {
	["new"] = "document_add", ["open"] = "folder_open", ["save"] = "save", ["saveas"] = "save_edit", ["print"] = "print",
	["printpreview"] = "document_search", ["pagesetup"] = "document_settings", ["exit"] = "dismiss",
	["export"] = "arrow_export_ltr", ["exportpdf"] = "document_pdf", ["about"] = "info",
	["undo"] = "arrow_undo", ["redo"] = "arrow_redo", ["paste"] = "clipboard_paste", ["cut"] = "cut", ["copy"] = "copy", ["formatpainter"] = "paint_brush",
	["selectall"] = "select_all_on", ["find"] = "search", ["replace"] = "arrow_swap",
	["font"] = "text_font", ["changecase"] = "text_change_case", ["underline"] = "text_underline",
	["bullets"] = "text_bullet_list_ltr", ["numbering"] = "text_number_list_ltr",
	["outdent"] = "text_indent_decrease_ltr", ["indent"] = "text_indent_increase_ltr", ["listoptions"] = "options",
	["alignleft"] = "text_align_left", ["aligncenter"] = "text_align_center", ["alignright"] = "text_align_right", ["justify"] = "text_align_justify",
	["pagebreak"] = "document_page_break", ["sectionbreak"] = "document_split_hint", ["table"] = "table", ["picture"] = "image", ["thumbnail"] = "image_multiple",
	["link"] = "link", ["crossreference"] = "number_symbol_square", ["updatereferences"] = "arrow_sync", ["footnote"] = "text_footnote", ["tableofcontents"] = "document_bullet_list", ["snippets"] = "text_expand", ["horizontalline"] = "divider_short", ["datetime"] = "calendar_clock", ["headerfooter"] = "document_header_footer", ["columns"] = "text_column_two",
	["nextspelling"] = "text_proofing_tools", ["dictionary"] = "book_add", ["language"] = "local_language",
	["structured"] = "text_header_1", ["freeform"] = "text_font_size",
	["fontbold"] = "text_bold", ["fontitalic"] = "text_italic", ["fontunderline"] = "text_underline", ["fontstrike"] = "text_strikethrough",
	["fontsub"] = "text_subscript", ["fontsuper"] = "text_superscript", ["fontgrow"] = "font_increase", ["fontshrink"] = "font_decrease",
	["fontcolor"] = "text_color", ["fonthighlight"] = "highlight",
	["zoomin"] = "zoom_in", ["zoomout"] = "zoom_out", ["zoom100"] = "zoom_fit", ["html"] = "code", ["markdown"] = "markdown",
	["viewrich"] = "document_text", ["sidebyside"] = "layout_column_two",
	["tablesettings"] = "table_settings", ["linespacing"] = "text_line_spacing", ["pagewidth"] = "document_width", ["printlayout"] = "document_one_page", ["onepage"] = "page_fit", ["multiplepages"] = "document_multiple", ["close"] = "document_dismiss", ["newtemplate"] = "document_add", ["startpage"] = "home", ["newwindow"] = "window_new", ["stylesmenu"] = "text_edit_style", ["modifystyle"] = "edit", ["newstyle"] = "add", ["updatestyle"] = "arrow_sync", ["deletestyle"] = "delete", ["savestyles"] = "save", ["clearformatting"] = "text_clear_formatting", ["paragraphsettings"] = "text_paragraph", ["applystyles"] = "paint_brush_arrow_down", ["openwindow"] = "folder_open", ["inserttemplate"] = "document_text", ["fillfields"] = "textbox", ["borderall"] = "border_all", ["borderoutside"] = "border_outside",
	["borderinside"] = "border_inside", ["bordernone"] = "border_none", ["paintbucket"] = "paint_bucket",
	["macros"] = "play", ["record"] = "record", ["hotkeys"] = "keyboard", ["runscript"] = "script", ["macrosfolder"] = "folder_open", ["eventhooks"] = "flash", ["commandpalette"] = "search", ["api"] = "plug_connected", ["extensions"] = "puzzle_piece",
	["insertabove"] = "table_stack_above", ["insertbelow"] = "table_stack_below", ["insertleft"] = "table_stack_left", ["insertright"] = "table_stack_right",
	["deleterow"] = "table_delete_row", ["deletecolumn"] = "table_delete_column", ["deletetable"] = "table_dismiss",
	["deletecells"] = "table_cell_cross", ["mergecells"] = "table_cells_merge", ["splitcells"] = "table_cells_split",    // (iteration 92)
};
// The colour a Fluent icon's inside takes (its Filled shape, tinted, under the outline), by kind of command; icons not
// named here (the text formatting ones) stay plain.
var families = new Dictionary<Color, string> {
	[Color.FromArgb(0x2B, 0x7C, 0xD3)] = "new newtemplate newwindow print printpreview pagesetup export viewrich printlayout onepage multiplepages pagewidth inserttemplate link tableofcontents headerfooter pagebreak sectionbreak sidebyside startpage about crossreference footnote",
	[Color.FromArgb(0x87, 0x64, 0xB8)] = "save saveas savestyles",
	[Color.FromArgb(0xE8, 0xA3, 0x17)] = "open openwindow macrosfolder paste eventhooks snippets",
	[Color.FromArgb(0xD1, 0x34, 0x38)] = "exit close exportpdf deletestyle deleterow deletecolumn deletetable deletecells record datetime",
	[Color.FromArgb(0x3A, 0x9A, 0x3A)] = "picture thumbnail table tablesettings insertabove insertbelow insertleft insertright mergecells splitcells macros runscript extensions api fillfields",
	[Color.FromArgb(0xCA, 0x50, 0x10)] = "stylesmenu applystyles formatpainter paintbucket modifystyle newstyle updatestyle dictionary nextspelling",
	[Color.FromArgb(0x03, 0x83, 0x87)] = "find replace commandpalette language zoomin zoomout zoom100 updatereferences hotkeys html markdown",
};
var tint = families.SelectMany(f => f.Value.Split(' ').Select(n => (n, f.Key))).ToDictionary(p => p.n, p => p.Key);
// arrows, whose Filled shape is only a thicker line (a halo under the outline): their outline takes the colour instead
var colorLine = new HashSet<string> { "replace", "updatereferences", "updatestyle", "zoom100" };
var drawn = new Dictionary<string, Action<Graphics, float>> {
	["legal"] = Icons.Legal,
	["linenumbers"] = Icons.LineNumbers,
	["themecolor"] = Icons.Palette,
	["exportword"] = (g, s) => Icons.WithLetter(g, s, "W", Color.FromArgb(0x18, 0x5A, 0xBD)),
	["exportodt"] = (g, s) => Icons.WithLetter(g, s, "O", Color.FromArgb(0x10, 0x7C, 0x41)),
};

var all = fluent.Keys.Concat(drawn.Keys).OrderBy(n => n).ToList();
int count = 0;
Action<Graphics, float> Draw(string name) => (g, s) => {
	if (!fluent.TryGetValue(name, out var icon)) drawn[name](g, s);
	else if (tint.TryGetValue(name, out var color)) Icons.TwoTone(g, s, icon, color, colorLine.Contains(name));
	else Icons.Fluent(g, s, icon);
};
foreach (int size in new[] { 16, 20, 24, 32, 40, 48, 64 })
	foreach (var name in all) {
		Icons.Save(Path.Combine(outDir, $"{name}_{size}.bmp"), size, Draw(name));
		count++;
	}
Icons.Sheet(Path.Combine(outDir, "_sheet.png"), all, Draw);
AHK.Return($"{count} icons in {outDir} (Fluent UI System Icons)");

static class Icons
{
	static PrivateFontCollection fonts = null!, filledFonts = null!;
	static FontFamily family = null!, filledFamily = null!;
	static Dictionary<string, int> codes = null!, filledCodes = null!;
	static Color Ink = Color.FromArgb(0x30, 0x30, 0x30), Accent = Color.FromArgb(0x00, 0x67, 0xC0);
	static bool light;

	public static void UseLightInk() => (Ink, Accent, light) = (Color.FromArgb(0xEC, 0xEC, 0xEC), Color.FromArgb(0x6C, 0xC4, 0xF4), true);

	// The Regular font (the outlines) and the Filled one (the same icons as solid shapes, for the tinted insides).
	public static void Load(string fontDir) {
		(fonts, family, codes) = LoadFont(fontDir, "Regular");
		(filledFonts, filledFamily, filledCodes) = LoadFont(fontDir, "Filled");
	}

	static (PrivateFontCollection, FontFamily, Dictionary<string, int>) LoadFont(string fontDir, string style) {
		var collection = new PrivateFontCollection();
		collection.AddFontFile(Path.Combine(fontDir, $"FluentSystemIcons-{style}.ttf"));
		var names = JsonSerializer.Deserialize<Dictionary<string, int>>(File.ReadAllText(Path.Combine(fontDir, $"FluentSystemIcons-{style}.json")))!;
		return (collection, collection.Families[0], names);
	}

	// A Fluent icon with its inside in a tint of color: the Filled shape (light on the light ribbon, a deeper shade on
	// the dark one) under the Regular outline. An icon with no Filled shape at that size is drawn plain. line: the outline
	// itself in color (lighter on the dark ribbon), with no inside.
	public static void TwoTone(Graphics g, float s, string icon, Color color, bool line = false) {
		if (line) {
			Fluent(g, s, icon, light ? Mix(color, Color.White, 0.35f) : color);
			return;
		}
		int design = DesignSize(icon, (int)Math.Round(s));
		if (filledCodes.TryGetValue($"ic_fluent_{icon}_{design}_filled", out int code))
			Glyph(g, filledFamily, code, new RectangleF(0, 0, s, s), light ? Mix(color, Color.Black, 0.3f) : Mix(color, Color.White, 0.55f));
		Fluent(g, s, icon);
	}

	static Color Mix(Color a, Color b, float t) =>
		Color.FromArgb((int)(a.R + (b.R - a.R) * t), (int)(a.G + (b.G - a.G) * t), (int)(a.B + (b.B - a.B) * t));

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
		g.Clear(light ? Color.FromArgb(0x2B, 0x2B, 0x2B) : Color.White);
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
		Glyph(g, family, codes[$"ic_fluent_{icon}_{design}_regular"], area, color ?? Ink);
	}

	static void Glyph(Graphics g, FontFamily font, int code, RectangleF area, Color color) {
		using var path = GlyphPath(font, code, area);
		using var brush = new SolidBrush(color);
		g.FillPath(brush, path);
	}

	static GraphicsPath GlyphPath(FontFamily font, int code, RectangleF area) {
		var path = new GraphicsPath();
		path.AddString(char.ConvertFromUtf32(code), font, 0, area.Width, new PointF(0, 0), StringFormat.GenericTypographic);
		float baseline = area.Width * font.GetCellAscent(FontStyle.Regular) / font.GetEmHeight(FontStyle.Regular);
		using var move = new Matrix();
		move.Translate(area.X, area.Y + area.Width - baseline);
		path.Transform(move);
		return path;
	}

	// Fluent's palette on a pale wood tint, its paint dots each in a colour (red, amber, green, blue, purple, from the left):
	// View > Theme. The dots are the outline's small closed figures; the rest (the rim, outside and inside) stays ink.
	public static void Palette(Graphics g, float s) {
		int design = DesignSize("color", (int)Math.Round(s));
		var area = new RectangleF(0, 0, s, s);
		var wood = light ? Color.FromArgb(0x8A, 0x6A, 0x48) : Color.FromArgb(0xF4, 0xE4, 0xC8);
		Glyph(g, filledFamily, filledCodes[$"ic_fluent_color_{design}_filled"], area, wood);
		using var outline = GlyphPath(family, codes[$"ic_fluent_color_{design}_regular"], area);
		using var rim = new GraphicsPath(FillMode.Alternate);
		var dots = new List<GraphicsPath>();
		using (var figures = new GraphicsPathIterator(outline)) {
			using var figure = new GraphicsPath();
			while (figures.NextSubpath(figure, out _) > 0) {
				var bounds = figure.GetBounds();
				if (bounds.Width < s * 0.3f && bounds.Height < s * 0.3f) dots.Add((GraphicsPath)figure.Clone());
				else rim.AddPath(figure, false);
				figure.Reset();
			}
		}
		using (var back = new SolidBrush(wood))    // (the Filled shape's own holes for the dots are not quite under them)
			foreach (var dot in dots) {
				var r = dot.GetBounds();
				r.Inflate(s * 0.04f, s * 0.04f);
				g.FillEllipse(back, r);
			}
		using (var ink = new SolidBrush(Ink)) g.FillPath(ink, rim);
		var paints = new[] { Color.FromArgb(0xE0, 0x40, 0x40), Color.FromArgb(0xF0, 0xB4, 0x20), Color.FromArgb(0x40, 0xA8, 0x40), Color.FromArgb(0x30, 0x80, 0xE0), Color.FromArgb(0x90, 0x50, 0xC8) };
		int i = 0;
		foreach (var dot in dots.OrderBy(d => d.GetBounds().X)) {
			using var paint = new SolidBrush(paints[i++ % paints.Length]);
			g.FillPath(paint, dot);
			dot.Dispose();
		}
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
