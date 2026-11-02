// The window's themes (View > Theme; Settings.Theme; api.Theme): the ribbon's colors and the title bar. The document page
// stays white in every theme.
//
// What the Windows Ribbon does with colors (seen on Windows 11, iteration 70): in its light mode the background is ignored,
// which leaves a near-white body, and only the highlight and the File button take a color; in its dark mode (DarkModeRibbon)
// the background is honored, darkened, under a black tab row. So every theme with a color is a dark ribbon, with light icons.

using System.Drawing;
using System.Runtime.InteropServices;
using RibbonLib;

sealed record AppTheme(string Name, bool Dark, Color Background, Color Highlight, Color FileButton, Color Caption)
{
	// In dark mode this "text color" doesn't touch the labels, which stay white. It is the text in the edit boxes, whose
	// backgrounds stay white: the lower it is the darker their text, until just under A9 that text turns white and is gone.
	static readonly Color BoxText = Color.FromArgb(0xB0, 0xB0, 0xB0);

	static Color Rgb(int rgb) => Color.FromArgb(rgb >> 16, (rgb >> 8) & 255, rgb & 255);

	// (the caption: a color the ribbon's black title reads on, see PaintTitleBar)
	static AppTheme Hue(string name, int background, int highlight, int caption) => new(name, true, Rgb(background), Rgb(highlight), Rgb(highlight), Rgb(caption));

	public static readonly AppTheme[] All = {
		new("Light", false, Rgb(0xF0F0F0), Color.White, Rgb(0x0F6B6B), Color.White),
		Hue("Dark", 0x303030, 0x505050, 0x9A9A9A),
		Hue("Teal", 0x0F5C5C, 0x1A7A7A, 0x3C9A9A),
		Hue("Navy", 0x1B3A6B, 0x2F5FA8, 0x7FA3D6),
		Hue("Forest", 0x1F5A33, 0x2F8A50, 0x6FB487),
		Hue("Plum", 0x4F2A6B, 0x7A45A0, 0xAB8CC8),
		Hue("Burgundy", 0x6B1F2E, 0xA03348, 0xCC8593),
		Hue("Bronze", 0x6B4A1A, 0xA07328, 0xC8A262),
	};

	public static AppTheme? Find(string? name) {
		if (string.Equals(name, "Color", StringComparison.OrdinalIgnoreCase)) name = "Teal";    // (its first name)
		return All.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
	}

	public static AppTheme Named(string? name) => Find(name) ?? All[0];

	// A little square of the theme's color, for its line in View > Theme.
	public Bitmap Swatch(int size) {
		var picture = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
		using var g = Graphics.FromImage(picture);
		g.Clear(Dark ? Highlight : Color.White);
		g.DrawRectangle(Pens.Gray, 0, 0, size - 1, size - 1);
		return picture;
	}

	public void Paint(Ribbon ribbon) {
		ribbon.DarkModeRibbon = Dark;
		ribbon.SetColors(Background, Highlight, Dark ? BoxText : Rgb(0x202020));
		ribbon.ApplicationButtonColor = FileButton;
	}

	[DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
	const int DWMWA_CAPTION_COLOR = 35;    // (Windows 11; ignored before)

	// The ribbon writes the title in black whatever the theme, so the bar is a color black reads on, even for the dark ones.
	public void PaintTitleBar(IntPtr window) {
		int color = ColorTranslator.ToWin32(Caption);
		DwmSetWindowAttribute(window, DWMWA_CAPTION_COLOR, ref color, sizeof(int));
	}
}
