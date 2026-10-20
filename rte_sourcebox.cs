// The source editor of the Markdown and HTML views, for RTE Writer (rte_writer38.cs, through #load): Scintilla, through
// the Scintilla5.NET package (MIT license; Scintilla and Lexilla are under their own permissive license), with the
// Markdown and HTML lexers' colours and line numbers. It has the members of a text box that the views use (Text,
// SelectionStart, SelectionLength, SelectedText, Select, Find, WordWrap ...), with the same meaning: character
// positions (as in .NET strings) and "\n" line ends, so the views work with it as with the RichTextBox it replaces.
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ScintillaNET;

sealed class SourceBox : Scintilla
{
	// Keys that RichEdit's plain-text box formatted with (kept out of the source editor; the views give some of them
	// Markdown or HTML meanings).
	public static readonly HashSet<Keys> FormattingKeys = new() {
		Keys.Control | Keys.L, Keys.Control | Keys.E, Keys.Control | Keys.R, Keys.Control | Keys.J, Keys.Control | Keys.D1,
		Keys.Control | Keys.D2, Keys.Control | Keys.D5, Keys.Control | Keys.Shift | Keys.A, Keys.Control | Keys.Shift | Keys.L,
		Keys.Control | Keys.Oemplus, Keys.Control | Keys.Shift | Keys.Oemplus, Keys.Control | Keys.B, Keys.Control | Keys.I, Keys.Control | Keys.U,
	};

	public enum Languages { None, Markdown, Html }

	// Csharp-in-AHK uses packages where NuGet keeps them, so Scintilla.dll and Lexilla.dll are not next to the program:
	// Scintilla5.NET is told where they are in the package (before the first Scintilla is made).
	static SourceBox() {
		if (!string.IsNullOrEmpty(ScintillaNativeLibrary.SatelliteDirectory)) return;
		string lib = Path.GetDirectoryName(typeof(Scintilla).Assembly.Location) ?? "";
		string platform = RuntimeInformation.ProcessArchitecture switch { System.Runtime.InteropServices.Architecture.Arm64 => "win-arm64", System.Runtime.InteropServices.Architecture.X86 => "win-x86", _ => "win-x64" };
		string native = Path.GetFullPath(Path.Combine(lib, "..", "..", "runtimes", platform, "native"));
		if (File.Exists(Path.Combine(native, "Scintilla.dll"))) ScintillaNativeLibrary.SatelliteDirectory = native;
	}

	public SourceBox() {
		BorderStyle = ScintillaNET.BorderStyle.None;
		EolMode = Eol.Lf;
		WrapMode = WrapMode.Word;
		WrapIndentMode = WrapIndentMode.Same;
		TabWidth = 4;
		Styles[Style.Default].Font = "Consolas";
		Styles[Style.Default].SizeF = 10.5f;
		StyleClearAll();
		Styles[Style.LineNumber].ForeColor = Color.FromArgb(0x90, 0x90, 0x90);
		Styles[Style.LineNumber].BackColor = Color.FromArgb(0xF5, 0xF5, 0xF5);
		Margins[1].Width = 0;    // (no symbol margin)
		UpdateUI += (s, e) => {
			if ((e.Change & (UpdateChange.Selection | UpdateChange.Content)) != 0) SelectionChanged?.Invoke(this, EventArgs.Empty);
		};
		Indicators[HeadingIndicator].Style = IndicatorStyle.TextFore;
		Indicators[HeadingIndicator].ForeColor = HeadingColour;
		// the lines an edit touched get their heading colour again, once Scintilla has finished with the change (marked
		// while it reports it, a line already drawn was not drawn again)
		Insert += (s, e) => HeadingsChanged(LineOf(e.Position), LineOf(e.Position) + Math.Max(0, e.LinesAdded));
		Delete += (s, e) => HeadingsChanged(LineOf(e.Position), LineOf(e.Position));
	}

	int changedFirst = -1, changedLast = -1;

	void HeadingsChanged(int first, int last) {
		bool queued = changedFirst >= 0;
		changedFirst = queued ? Math.Min(changedFirst, first) : first;
		changedLast = queued ? Math.Max(changedLast, last) : last;
		if (queued) return;
		if (IsHandleCreated) BeginInvoke(MarkChangedHeadings);
		else MarkChangedHeadings();
	}

	void MarkChangedHeadings() {
		int first = changedFirst, last = changedLast;
		changedFirst = changedLast = -1;
		if (first >= 0 && !IsDisposed) MarkHeadings(first, last);
	}

	// Lexilla's Markdown lexer colours a heading's "#" marks only: the rest of a heading line gets their colour from an
	// indicator (which lexing leaves alone).
	const int HeadingIndicator = 8;    // (0 to 7 are the lexers')
	static readonly Color HeadingColour = Color.FromArgb(0x1F, 0x4E, 0x8C);
	static readonly System.Text.RegularExpressions.Regex HeadingLine = new(@"^ {0,3}#{1,6}(?=[ \t]|$)");

	void MarkHeadings(int first, int last) {
		if (Lines.Count == 0) return;
		last = Math.Min(last, Lines.Count - 1);
		if (first > last) return;
		IndicatorCurrent = HeadingIndicator;
		int start = Lines[first].Position, end = Lines[last].EndPosition;
		IndicatorClearRange(start, end - start);
		if (language != Languages.Markdown) return;
		for (int line = first; line <= last; line++) {
			string text = Lines[line].Text;
			if (HeadingLine.Match(text) is not { Success: true } marks) continue;
			int length = text.TrimEnd('\r', '\n').Length - marks.Length;
			if (length > 0) IndicatorFillRange(Lines[line].Position + marks.Length, length);
		}
	}

	protected override void OnHandleCreated(EventArgs e) {
		base.OnHandleCreated(e);
		Margins.Left = LogicalToDeviceUnits(6);
		Margins.Right = LogicalToDeviceUnits(6);
		FitLineNumbers();
	}

	protected override void OnTextChanged(EventArgs e) {
		base.OnTextChanged(e);
		FitLineNumbers();
	}

	// ---- the language: the lexer and its colours

	Languages language;

	public Languages Language {
		get => language;
		set {
			if (value == language) return;
			language = value;
			StyleClearAll();    // (every style as the default style, before the language's colours)
			Styles[Style.LineNumber].ForeColor = Color.FromArgb(0x90, 0x90, 0x90);
			Styles[Style.LineNumber].BackColor = Color.FromArgb(0xF5, 0xF5, 0xF5);
			LexerName = value switch { Languages.Markdown => "markdown", Languages.Html => "hypertext", _ => "" };
			if (value == Languages.Markdown) MarkdownColours();
			else if (value == Languages.Html) HtmlColours();
			Colorize(0, TextLength);
			MarkHeadings(0, Lines.Count - 1);
		}
	}

	void Colour(int style, Color fore, bool bold = false, bool italic = false, Color? back = null) {
		Styles[style].ForeColor = fore;
		Styles[style].Bold = bold;
		Styles[style].Italic = italic;
		if (back is Color b) Styles[style].BackColor = b;
	}

	// Lexilla's Markdown styles (SCE_MARKDOWN_*)
	void MarkdownColours() {
		for (int style = 6; style <= 11; style++) Colour(style, HeadingColour, bold: true);   // HEADER1 ... HEADER6 (the "#" marks)
		Colour(2, Color.Black, bold: true); Colour(3, Color.Black, bold: true);               // STRONG1, STRONG2
		Colour(4, Color.Black, italic: true); Colour(5, Color.Black, italic: true);           // EM1, EM2
		Colour(13, Color.FromArgb(0xA0, 0x52, 0x2D)); Colour(14, Color.FromArgb(0xA0, 0x52, 0x2D));    // ULIST_ITEM, OLIST_ITEM
		Colour(15, Color.FromArgb(0x5A, 0x7D, 0x5A));                                          // BLOCKQUOTE
		Colour(16, Color.FromArgb(0x80, 0x80, 0x80));                                          // STRIKEOUT
		Colour(17, Color.FromArgb(0x80, 0x80, 0x80));                                          // HRULE
		Colour(18, Color.FromArgb(0x00, 0x56, 0xB3));                                          // LINK
		var code = Color.FromArgb(0xF3, 0xF3, 0xF3);
		for (int style = 19; style <= 21; style++) Colour(style, Color.FromArgb(0xA3, 0x15, 0x15), back: code);    // CODE, CODE2, CODEBK
	}

	// Lexilla's HTML styles (SCE_H_*), coloured as Visual Studio does
	void HtmlColours() {
		var element = Color.FromArgb(0x80, 0x00, 0x00);
		Colour(1, element); Colour(11, element);                                               // TAG, TAGEND
		Colour(2, Color.FromArgb(0xC0, 0x00, 0x00));                                           // TAGUNKNOWN
		Colour(3, Color.FromArgb(0xE5, 0x00, 0x00)); Colour(4, Color.FromArgb(0xE5, 0x00, 0x00));    // ATTRIBUTE, ATTRIBUTEUNKNOWN
		var value = Color.FromArgb(0x00, 0x00, 0xFF);
		Colour(5, value); Colour(6, value); Colour(7, value); Colour(19, value);               // NUMBER, DOUBLESTRING, SINGLESTRING, VALUE
		Colour(8, Color.FromArgb(0x00, 0x00, 0xA0));                                           // OTHER (inside tags)
		Colour(9, Color.FromArgb(0x00, 0x80, 0x00), italic: true);                              // COMMENT
		Colour(10, Color.FromArgb(0x80, 0x80, 0x80));                                          // ENTITY
		Colour(12, element); Colour(13, element);                                              // XMLSTART, XMLEND
		Colour(21, Color.FromArgb(0x80, 0x80, 0x80));                                          // SGML (<!DOCTYPE>)
	}

	// ---- line numbers

	bool lineNumbers = true;

	/// Whether line numbers are shown left of the text.
	public bool LineNumbers {
		get => lineNumbers;
		set { lineNumbers = value; FitLineNumbers(); }
	}

	int numberDigits = -1;

	// The margin fits the largest line number (at least three digits); worked out again when that gains a digit.
	void FitLineNumbers() {
		if (!IsHandleCreated) return;
		int digits = lineNumbers ? Math.Max(3, Lines.Count.ToString().Length) : 0;
		if (digits == numberDigits) return;
		numberDigits = digits;
		Margins[0].Type = MarginType.Number;
		Margins[0].Width = digits == 0 ? 0 : TextWidth(Style.LineNumber, new string('9', digits)) + LogicalToDeviceUnits(10);
	}

	// ---- the members of a text box that the views use (character positions, "\n" line ends)

	public override string Text {
		get => base.Text;
		set => base.Text = (value ?? "").Replace("\r\n", "\n");
	}

	/// The start of the selection (the caret, if nothing is selected).
	public new int SelectionStart => base.SelectionStart;

	public int SelectionLength => SelectionEnd - SelectionStart;

	/// The selected text; setting it replaces the selection (one undo step) and puts the caret after it.
	public new string SelectedText {
		get => base.SelectedText;
		set => ReplaceSelection((value ?? "").Replace("\r\n", "\n"));
	}

	/// Selects length characters from start, with the caret at the end (as TextBoxBase.Select).
	public void Select(int start, int length) {
		start = Math.Clamp(start, 0, TextLength);
		SetSelection(Math.Clamp(start + length, 0, TextLength), start);
	}

	public bool WordWrap {
		get => WrapMode != WrapMode.None;
		set => WrapMode = value ? WrapMode.Word : WrapMode.None;
	}

	public void ScrollToCaret() => ScrollCaret();

	public void ClearUndo() => EmptyUndoBuffer();

	/// Pastes the clipboard's text (Scintilla pastes text only, with its line ends made "\n").
	public void Paste(DataFormats.Format format) => Paste();

	/// After the selection or the text changed (when Scintilla updates the display).
	public event EventHandler? SelectionChanged;

	/// Finds text between start and end (match case / whole word as RichTextBox.Find), selects it and returns its position,
	/// or -1.
	public int Find(string text, int start, int end, RichTextBoxFinds options) {
		if (text == "") return -1;
		TargetStart = Math.Clamp(start, 0, TextLength);
		TargetEnd = Math.Clamp(end, 0, TextLength);
		SearchFlags = ((options & RichTextBoxFinds.MatchCase) != 0 ? SearchFlags.MatchCase : SearchFlags.None)
			| ((options & RichTextBoxFinds.WholeWord) != 0 ? SearchFlags.WholeWord : SearchFlags.None);
		int at = SearchInTarget(text);
		if (at >= 0) SetSelection(TargetEnd, TargetStart);
		return at;
	}

	/// The line (from 0) with the character at a position, and the position where a line starts.
	public int LineOf(int position) => LineFromPosition(Math.Clamp(position, 0, TextLength));
	public int LineStart(int line) => Lines[line].Position;
}
