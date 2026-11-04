// The window's themes (View > Theme; Settings.Theme; api.Theme): the ribbon's colors and the title bar. The document page
// stays white in every theme.
//
// What the Windows Ribbon does with colors (seen on Windows 11, iteration 70): in its light mode the background is ignored,
// which leaves a near-white body, and only the highlight and the File button take a color; in its dark mode (DarkModeRibbon)
// the background is honored, darkened, under a black tab row. So every theme with a color is a dark ribbon, with light icons.

using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

	// The darker themes darken every window of the program, dialogs included, through the system colors WinForms hands out. It
	// holds for the whole process, so, like the ribbon, it is set once, before the first window. Light leaves it all alone.
	// It must come before any control is made, not just before the first window: a control's fields start WinForms' window for
	// system events, and then SetColorMode signals that window and waits for it, which could hang the start for good, with no
	// window at all (seen in AhkWord.ahk on its main thread, never on the tests' own thread).
	public void SetColorMode() {
		if (!Dark) return;
		DarkWindows = true;
#pragma warning disable WFO5001
		Application.SetColorMode(SystemColorMode.Dark);
#pragma warning restore WFO5001
		PrintLayoutPage.Surround = Surround;
	}

	public static bool DarkWindows { get; private set; }

	// What shows a document stays as it prints, in every theme: the page, its strips, the previews of a paragraph or a style.
	public static Color Paper => DarkWindows ? Color.White : SystemColors.Window;
	public static Color Ink => DarkWindows ? Color.Black : SystemColors.WindowText;
	public static Color FaintInk => DarkWindows ? Color.Gray : SystemColors.GrayText;
	public static T OnPaper<T>(T control) where T : Control { if (DarkWindows) { control.BackColor = Paper; control.ForeColor = Ink; } return control; }

	public static void DarkLink(LinkLabel link) { if (DarkWindows) link.LinkColor = link.ActiveLinkColor = link.VisitedLinkColor = Rgb(0x8AB4F8); }    // (the usual blue can't be read there)

	public static Color Warning => DarkWindows ? Rgb(0xFF9A8A) : Color.Firebrick;    // (Firebrick is too dark to read on the darker themes)

	static Color Shade(Color color, double by) => Color.FromArgb((int)(color.R * by), (int)(color.G * by), (int)(color.B * by));

	// The window around the page, in the darker themes.
	public Color Panel => Shade(Background, 0.72);      // the outline, the source view's title
	public Color Surround => Shade(Background, 0.5);    // around the page

	[DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
	const int DWMWA_CAPTION_COLOR = 35;    // (Windows 11; ignored before)

	// The ribbon writes the title in black whatever the theme, so the bar is a color black reads on, even for the dark ones.
	public void PaintTitleBar(IntPtr window) {
		int color = ColorTranslator.ToWin32(Caption);
		DwmSetWindowAttribute(window, DWMWA_CAPTION_COLOR, ref color, sizeof(int));
	}
}

// A strip in its own BackColor: the usual renderer paints its own gradient over it.
sealed class FlatStripRenderer : ToolStripProfessionalRenderer
{
	protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) {
		using var back = new SolidBrush(e.BackColor);
		e.Graphics.FillRectangle(back, e.AffectedBounds);
	}

	protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }
}

// The dialogs' common look (iteration 74): the same margin round the contents (scaled for the screen), and buttons at
// the right of one width at least, OK before Cancel.
static class Dialogs
{
	const int Margin = 10, ButtonWidth = 80;    // (logical pixels)

	/// A dialog: fixed and sized to its contents, or sizable (its size then set by the caller); no minimize and maximize
	/// boxes, not in the taskbar, centered on its owner.
	public static void Setup(Form form, string title, bool sizable = false) {
		form.Text = title;
		form.FormBorderStyle = sizable ? FormBorderStyle.Sizable : FormBorderStyle.FixedDialog;
		form.MinimizeBox = form.MaximizeBox = form.ShowInTaskbar = false;
		form.StartPosition = FormStartPosition.CenterParent;
		// (the font stays WinForms' own: Windows' message box font, on some screens, looked softer than it, iteration 73)
		if (!sizable) {
			form.AutoSize = true;
			form.AutoSizeMode = AutoSizeMode.GrowAndShrink;
		}
		form.Padding = new Padding(form.LogicalToDeviceUnits(Margin));
	}

	public static Button Button(Control scale, string text, DialogResult result = DialogResult.None) =>
		new() { Text = text, DialogResult = result, AutoSize = true, MinimumSize = new Size(scale.LogicalToDeviceUnits(ButtonWidth), 0) };

	/// OK (closes: it ends the dialog by itself; otherwise its Click handler decides).
	public static Button Ok(Control scale, bool closes = true) => Button(scale, T("dlg.ok", "OK"), closes ? DialogResult.OK : DialogResult.None);

	/// Cancel, or another text for the button that ends the dialog with Cancel (Close, Discard).
	public static Button Cancel(Control scale, string? text = null) => Button(scale, text ?? T("dlg.cancel", "Cancel"), DialogResult.Cancel);

	/// Widens a box (never narrows it) so that the longest of texts fits, in font: a translation can be longer than the
	/// English the box's width was chosen for. A combo box: its items, and room for its arrow.
	public static void FitWidth(Control box, Font font, IEnumerable<string> texts) {
		int extra = box.LogicalToDeviceUnits(10) + (box is ComboBox ? SystemInformation.VerticalScrollBarWidth : 0);
		int need = texts.Select(t => TextRenderer.MeasureText(t, font).Width).DefaultIfEmpty(0).Max() + extra;
		if (need > box.Width) box.Width = need;
	}

	public static void FitWidth(ComboBox box, Font font) => FitWidth(box, font, box.Items.Cast<object>().Select(o => o.ToString() ?? ""));

	/// The buttons at the right, in this order from the left, with a gap above them (logical pixels).
	public static FlowLayoutPanel Row(Control scale, int gapAbove, params Button[] buttons) {
		var row = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, scale.LogicalToDeviceUnits(gapAbove), 0, 0) };
		row.Controls.AddRange(buttons.Reverse().ToArray<Control>());
		return row;
	}
}
