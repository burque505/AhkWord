// Print Layout for RTE Writer (rte_writer29.cs, through #load): lines break where they will on the printed page, as
// in Word's Print Layout and LibreOffice, not at the window's edge (which made Word and OpenDocument files open with very
// long lines). RichEdit lays its lines out for a target device (EM_SETTARGETDEVICE, as WordPad's "wrap to ruler"):
// the printer's information context and the text width of Page Setup's paper and margins, so the editor's lines are
// the printed ones (and so are its line numbers). The page (its margins and the text) is shown centred on grey; a
// window narrower than the page scrolls sideways. Without a printer, the screen lays the lines out at the same width.
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EllipticBit.RichEditorNET;

sealed class PrintLayoutTarget : IDisposable
{
	[DllImport("gdi32.dll", CharSet = CharSet.Unicode)] static extern IntPtr CreateIC(string? driver, string? device, string? port, IntPtr devMode);
	[DllImport("gdi32.dll")] static extern bool DeleteDC(IntPtr hdc);
	[DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	const int EM_SETTARGETDEVICE = 0x0448;

	/// An information context: never drawn on, only measured with (RichEdit keeps it while it is the target).
	public IntPtr Hdc { get; private set; }
	/// The printer's (false: the screen's, when there is no printer).
	public bool IsPrinter { get; }
	public string Device { get; }

	public PrintLayoutTarget(string? printerName) {
		if (!string.IsNullOrEmpty(printerName)) Hdc = CreateIC("WINSPOOL", printerName, null, IntPtr.Zero);
		IsPrinter = Hdc != IntPtr.Zero;
		if (!IsPrinter) Hdc = CreateIC("DISPLAY", null, null, IntPtr.Zero);
		Device = IsPrinter ? printerName! : "the screen";
	}

	/// The editor's lines laid out for this device, width twips wide.
	public void Apply(RichEditBox editor, int width) {
		if (editor.IsHandleCreated && Hdc != IntPtr.Zero) SendMessage(editor.Handle, EM_SETTARGETDEVICE, Hdc, (IntPtr)width);
	}

	/// The editor back to the window's width (wrap) or no wrapping at all.
	public static void Release(RichEditBox editor, bool wrap) {
		if (editor.IsHandleCreated) SendMessage(editor.Handle, EM_SETTARGETDEVICE, IntPtr.Zero, (IntPtr)(wrap ? 0 : 1));
	}

	public void Dispose() {
		if (Hdc != IntPtr.Zero) { DeleteDC(Hdc); Hdc = IntPtr.Zero; }
	}
}

static class PrintLayoutPage
{
	public static readonly Color Surround = Color.FromArgb(0xE6, 0xE6, 0xE6);    // the grey around the page, as in Word

	/// The paper's width and height in pixels at a zoom of 1.
	public static (double Width, double Height) Paper(PageLayout layout, int dpi) => (layout.Width / 1440.0 * dpi, layout.Height / 1440.0 * dpi);

	/// The page's parts in pixels for the zoom and DPI: its left margin, text width and right margin.
	public static (int Left, int Text, int Right) Pixels(PageLayout layout, float zoom, int dpi) {
		int Px(int twips) => (int)Math.Round(twips / 1440.0 * dpi * zoom);
		return (Px(layout.Left), Px(layout.TextWidth), Px(layout.Right));
	}
}
