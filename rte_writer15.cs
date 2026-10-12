#:package EllipticBit.RichEditorNET
#load "rte_markdown5.cs"
#load "rte_htmllists2.cs"
#load "rte_tables.cs"
#load "rte_office6.cs"
#load "rte_legal.cs"
#load "rte_pagebreaks.cs"
#load "rte_headerfooter.cs"
// RTE Writer: a word processor built on EllipticBit.RichEditorNET, hosted in AutoHotkey by Csharp-in-AHK.
// Run it with rte_writer15.ahk. Iteration 15 adds to iteration 14:
//   - Pictures' alt text can be seen and changed: resting the mouse on a picture shows it (or "no alt text"); right-click
//     a picture > Alt text... changes it (RichEditorNET can't change it, so the picture is replaced by the same picture
//     with the new alt text, in one undo step); Insert > Image asks for it (the file name was used without asking).
// Iteration 14 added to iteration 13:
//   - The document's header and footer are shown in strips above and below the text (BandStrip), filled in as on the
//     first page ({page} is 1, {pages} the number of pages it prints on, worked out once typing pauses); clicking one
//     opens File > Header and Footer. View > Header and footer shows or hides the strips.
//   - Pictures keep their alt text through RTF (rte_office6.cs): the RTF RTE Writer writes (files, Word, LibreOffice)
//     has it as each picture's description, and pictures read from RTF get it back.
// Iteration 13 added to iteration 12:
//   - Headers and footers per document (rte_headerfooter.cs): File > Header and Footer sets the document's own (and,
//     if asked, the default for new documents, which documents without one of their own get). They are printed as
//     before, and saved: as real headers and footers in RTF, and so in Word and OpenDocument files (made by Word,
//     LibreOffice or the built-in converter), with fields for {page} {pages} {date} {time}, tab stops for the left,
//     center and right parts, and the paper and margins of Page Setup; as a <meta> in HTML and a comment in Markdown.
//     Headers and footers made or changed in Word or LibreOffice are read back.
// Iteration 12 added to iteration 11:
//   - Manual page breaks (rte_pagebreaks.cs): Insert > Page break (Ctrl+Enter) puts one at the caret, as a paragraph of
//     its own, drawn as a dashed line with "Page break" (RichEdit shows nothing there). Printing starts a new page
//     after it. RTF keeps it as \page, and so Word and LibreOffice files; HTML and Markdown (rte_markdown5.cs) have it
//     as <div style="page-break-after: always; break-after: page"></div>, which browsers obey when printing, and the
//     built-in .docx converter as a Word page break (rte_office4.cs). Page breaks in documents from elsewhere (Word's
//     "\page" in the middle of a paragraph, HTML page-break styles, form feeds in text) become such paragraphs.
// Iteration 11 added to iteration 10:
//   - Exports through Word and LibreOffice (.docx, .odt) give legal lists real list numbering, so they renumber in Word
//     and LibreOffice when edited (rte_office3.cs: the RTF they get has the legal items as an RTF list).
// Iteration 10 added to iteration 9:
//   - Legal numbering (rte_legal.cs): Paragraph > Legal numbering (Ctrl+Shift+G, or the 1.1 button) numbers items 1.,
//     1.1, 1.1.1 by their list level; Tab / Shift+Tab, Enter, Backspace and Delete renumber them in the same undo step;
//     Restart / Continue numbering work on them; Paragraph > List levels sets how the numbers look (1. 1.1 / 1. 1.1. /
//     1 1.1). RichEdit cannot show compound numbers, so they are text in front of each item, kept up to date. They are
//     saved as <ol class="rte-legal"> in HTML (rte_htmllists2.cs), as [type=legal] lists in Markdown
//     (rte_markdown4.cs), as Word legal numbering by the built-in .docx converter (rte_office2.cs), and as the numbers
//     themselves in RTF (and so by Word and LibreOffice). Paragraphs that start with typed numbers such as "2.1 " take
//     them as their level when they are made legal, and opened documents with them become legal lists.
// Iteration 9 added to iteration 8:
//   - Tables (rte_tables.cs): Table > Insert table, insert and delete rows and columns, delete the table; Tab and
//     Shift+Tab move between cells; right-click in a table for the table commands; Ctrl+Shift+Insert inserts a row
//     below, Ctrl+Shift+Delete deletes the row, and Delete or Backspace deletes rows that are selected whole. They are
//     real RichEdit tables, and are saved as <table> in HTML and as GitHub tables in Markdown (rte_markdown3.cs), and
//     read back from both.
//   - Rich Text Format: open and save .rtf, which keeps everything the editor has.
//   - Word and OpenDocument (rte_office.cs): open .docx, .doc and .odt; save as .docx or .odt; File > Export as Word /
//     Export as OpenDocument (the current file stays as it is). Done through an invisible Word (only when Word isn't
//     running), LibreOffice (headless, with its own profile), the built-in Open XML converter (.docx only, NuGet
//     packages) or pandoc, whichever is available.
// Iteration 8 added to iteration 7:
//   - Markdown keeps alignment, colors, highlight, fonts, list styles and image sizes, as HTML and M-ArkDown
//     extensions (rte_markdown2.cs).
//   - Paragraph > List levels: the numbering style, punctuation and start number of each list level 1-9, used for
//     numbered lists and when nesting with Tab, and applied to the current list on request.
//   - Paragraph > Restart numbering / Continue numbering.
//   - HTML keeps list start numbers, restarts and punctuation (rte_htmllists.cs), which RichEditorNET's HTML drops.
// Iteration 7 replaced RichEditorNET's Markdown support with rte_markdown.cs:
//   - Opening and saving .md files, View > Markdown source and Rte.Markdown / GetMarkdown / SetMarkdown use real
//     Markdown: blank lines between paragraphs, list numbers, code, quotes, rules, tables and images (saved as files in
//     a "<name>_images" folder next to the document).
// Iteration 6 added macro power to iteration 5:
//   - Tools > Record macro (Ctrl+Shift+R): typing, cursor keys and the formatting commands of the menus and toolbar are
//     recorded as RteClient2 calls; stopping offers to save them as an AHK script in rte_macros, with a hotkey.
//     Mouse clicks are not recorded (positions differ from document to document).
//   - Tools > Macro hotkeys: a hotkey for any macro, active while the editor window is active.
// Iteration 5 added writing tools to iteration 4:
//   - Spelling: right-click a misspelled word (red underline) for suggestions, "delete repeated word" and "add to
//     dictionary"; F7 goes to the next spelling error. Uses the Windows spell checker, like the underlines.
//   - Outline: a panel listing the headings (View > Outline, Ctrl+Shift+O); click one to go there; right-click to
//     promote or demote it. It follows the caret and updates as you type.
// Iteration 4 added editable print headers and footers to iteration 3:
//   - File > Header and Footer: show or hide each, with left, center and right parts; the fields {title} {file} {page}
//     {pages} {date} {time} are filled in on each page. Headers and footers are drawn in the page margins.
// Iteration 3 added comfort features to iteration 2:
//   - File > Open Recent; the window size and position, view settings, last folder and page setup are remembered in
//     %APPDATA%\RTE Writer\settings.json (or the file named by the RTE_WRITER_SETTINGS environment variable).
//   - Drag and drop: drop a document to open it, or images to insert them where they are dropped.
//   - Page Setup, Print Preview, Print, and Export as PDF (through the "Microsoft Print to PDF" printer).
// Iteration 2 added events and fire-and-forget calls to iteration 1:
//   - Events: AHK scripts subscribe with Rte.On(Event, Callback) (RteClient2.ahk). The editor posts each event to the
//     subscriber's window as the registered message "RTE_WRITER_EVENT" (wParam = event number, lParam = event data),
//     so it never waits for a script. TextChanged and SelectionChanged are sent once typing / moving pauses briefly.
//   - Rte.Post(Name, args...) queues a call and returns at once; queued calls run in order, before any later call.
//     Errors in posted calls go to the status bar, Rte.LastAsyncError and the AsyncError event.
// Iteration 1:
//   - Documents: HTML (.html/.htm), Markdown (.md, GitHub flavor) and plain text (.txt).
//   - Formatting toolbar and menus; Tab / Shift+Tab nest list items, with automatic numbering styles per level.
//   - Automation: the public members of EditorApi, plus the main RichEditBox members, can be called by name from AHK
//     (see RteClient.ahk):
//       * from other AHK processes, through a hidden window "RTE Writer automation server <pid>" (WM_COPYDATA);
//       * from the hosting AHK process, through the delegate in AppContext.GetData("rte.dispatch").
//     Request: verb, member name and arguments separated by U+001F; verb is "call", "get" or "set".
//     Reply: "0" + result as text, or "1" + error message.
//     Tools > Macros runs the AHK scripts in the rte_macros folder (with RTE_WRITER_HWND set to this editor's server).
// args: [0] AutoHotkey exe (A_AhkPath), [1] folder of rte_writer.ahk, [2] optional file to open.
using System.Collections;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.Formatting;
using EllipticBit.RichEditorNET.TextObjectModel2;

// Runs on AutoHotkey's (STA) thread. ShowDialog runs the WinForms message loop until the window is closed;
// AHK hotkeys and timers keep running meanwhile and can call the editor (see RteClient.ahk).
Application.EnableVisualStyles();
using var form = new WriterForm(
	args.Length > 0 ? args[0] : "",
	args.Length > 1 ? args[1] : Environment.CurrentDirectory,
	args.Length > 2 && args[2] != "" ? args[2] : null);
form.ShowDialog();
AHK.Return(form.FileName ?? "");

class WriterForm : Form
{
	public const string AppName = "RTE Writer";
	public readonly WriterEditBox Editor = new() { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 11f), BorderStyle = BorderStyle.None };
	public string? FileName { get; private set; }
	public bool AutoLevelStyles = true;

	readonly string ahkPath, appDir;
	readonly EditorApi api;
	readonly AutomationServer server;
	readonly ToolStripStatusLabel statusText = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
	readonly ToolStripStatusLabel positionLabel = new(), wordsLabel = new(), modeLabel = new(), automationLabel = new();
	readonly System.Windows.Forms.Timer wordCountTimer = new() { Interval = 400 };
	readonly Dictionary<string, ToolStripButton> toggles = new();
	readonly Dictionary<ParagraphAlignment, ToolStripButton> alignButtons = new();
	ToolStripComboBox blockCombo = null!, fontCombo = null!, sizeCombo = null!;
	ToolStripSplitButton fontColorButton = null!, highlightButton = null!;
	ToolStripButton bulletButton = null!, numberButton = null!;
	ToolStripMenuItem structuredItem = null!, freeFormItem = null!, wordWrapItem = null!, spellCheckItem = null!, autoLevelItem = null!, bandsItem = null!;
	ToolStripMenuItem macrosMenu = null!;
	Color fontColor = Color.FromArgb(0xC0, 0, 0), highlightColor = Color.Yellow;
	FindReplaceForm? findForm;
	bool updatingUi, discardOnClose, lastModified;

	// Events for AHK scripts, and calls queued by Rte.Post
	internal readonly EventHub Events = new();

	// Iteration 3: settings, recent files, drag and drop
	internal readonly Settings Settings = Settings.Load();

	// Iteration 5: spelling and outline
	internal readonly Spelling? Spelling = Spelling.Create();

	// Iteration 6: macro recording and hotkeys
	readonly MacroRecorder recorder = new();
	readonly ToolStripStatusLabel recordingLabel = new() { ForeColor = Color.Firebrick, Visible = false, IsLink = false, ToolTipText = "Recording a macro - click to stop" };
	ToolStripMenuItem recordItem = null!;
	readonly SplitContainer outlineSplit = new() { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1 };
	readonly TreeView outlineTree = new() { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, HideSelection = false, FullRowSelect = true, ShowLines = false, ShowPlusMinus = true };
	readonly Label outlineHint = new() { Dock = DockStyle.Fill, ForeColor = SystemColors.GrayText, Padding = new Padding(6), Visible = false,
		Text = "No headings yet.\n\nGive paragraphs a heading style (Paragraph > Block style, Ctrl+1 to Ctrl+6) and they appear here." };
	readonly System.Windows.Forms.Timer outlineTimer = new() { Interval = 400 };
	List<(int Level, int Start, string Text)> outline = new();
	ToolStripMenuItem outlineItem = null!;
	bool updatingOutline, outlineBuilt;
	ToolStripMenuItem recentMenu = null!;
	FileDropTarget? dropTarget;
	readonly System.Windows.Forms.Timer textEventTimer = new() { Interval = 200 }, selectionEventTimer = new() { Interval = 120 };
	readonly Queue<string[]> postedCalls = new();

	// Iteration 10: legal numbering
	ToolStripButton legalButton = null!;
	readonly System.Windows.Forms.Timer legalTimer = new() { Interval = 300 };
	bool skipLegalCheck, guardingCaret, guardPending;
	int lastCaret = -1;

	public WriterForm(string ahkPath, string appDir, string? startFile) {
		this.ahkPath = ahkPath;
		this.appDir = appDir;
		Bands = DefaultBands();
		Text = AppName;
		StartPosition = FormStartPosition.CenterScreen;
		Size = new Size(LogicalToDeviceUnits(1150), LogicalToDeviceUnits(820));
		Editor.EnableHtmlFontSizing = Settings.Structured;    // structured documents (headings) by default; View menu switches
		Editor.WordWrap = Settings.WordWrap;
		Editor.EnableSpellCheck = Settings.SpellCheck;
		AutoLevelStyles = Settings.AutoLevelStyles;
		ApplyWindowSettings();

		api = new EditorApi(this);
		var page = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Window,
			Padding = new Padding(LogicalToDeviceUnits(28), LogicalToDeviceUnits(14), LogicalToDeviceUnits(16), 0) };
		page.Controls.Add(Editor);
		page.Controls.Add(headerStrip);    // (added after the editor, so docked before it)
		page.Controls.Add(footerStrip);
		headerStrip.Click += (s, e) => Run(HeaderFooterCommand);
		footerStrip.Click += (s, e) => Run(HeaderFooterCommand);
		Editor.TextChanged += (s, e) => { bandPagesStale = true; lastEdit = DateTime.Now; };
		bandTimer.Tick += (s, e) => UpdateBands();
		bandTimer.Start();
		var outlineTitle = new Label { Text = "Outline", Dock = DockStyle.Top, Height = LogicalToDeviceUnits(26), Padding = new Padding(6, 6, 0, 0),
			Font = new Font(Font, FontStyle.Bold), ForeColor = SystemColors.GrayText };
		outlineSplit.Panel1.BackColor = outlineTree.BackColor = outlineHint.BackColor = SystemColors.Control;
		outlineSplit.Panel1.Controls.Add(outlineTree);
		outlineSplit.Panel1.Controls.Add(outlineHint);
		outlineSplit.Panel1.Controls.Add(outlineTitle);
		outlineSplit.Panel2.Controls.Add(page);
		outlineSplit.Panel1Collapsed = !Settings.ShowOutline;
		var status = new StatusStrip();
		status.Items.AddRange(new ToolStripItem[] { recordingLabel, statusText, positionLabel, wordsLabel, modeLabel, automationLabel });
		recordingLabel.Font = new Font(status.Font, FontStyle.Bold);
		recordingLabel.Click += (s, e) => Run(ToggleRecording);
		var toolbar = BuildToolbar();
		var menu = BuildMenu();
		Controls.Add(outlineSplit);    // docking order: the last control added is docked first
		Controls.Add(toolbar);
		Controls.Add(menu);
		Controls.Add(status);
		MainMenuStrip = menu;

		// Automation: in-process (AppContext) and cross-process (hidden window) entry points.
		AppContext.SetData("rte.dispatch", (Func<string[], string>)DispatchFromAutomation);
		server = new AutomationServer(DispatchFromAutomation);
		automationLabel.Text = "Automation: ready";
		automationLabel.ToolTipText = $"AHK scripts can control this editor (RteClient2.ahk). Server window 0x{server.Handle.ToInt64():X}";

		Editor.SelectionChanged += (s, e) => {
			// the caret guard looks once the command that moved the caret has finished (commands move it on the way)
			if (!guardPending && IsHandleCreated) { guardPending = true; BeginInvoke(() => { guardPending = false; GuardLegalCaret(); }); }
			UpdateUi();
			Restart(selectionEventTimer);
		};
		Editor.TextChanged += (s, e) => { wordCountTimer.Stop(); wordCountTimer.Start(); UpdateTitle(); Restart(textEventTimer); Restart(outlineTimer); Restart(legalTimer); };
		legalTimer.Tick += (s, e) => { legalTimer.Stop(); CheckLegalNumbers(); };
		LegalLists.Format = Enum.TryParse<LegalLists.Periods>(Settings.LegalFormat, true, out var legalFormat) ? legalFormat : LegalLists.Periods.Top;
		Editor.ContextMenuHook = ShowContextMenu;
		Editor.OpenLink = OpenLinkCommand;
		Editor.AltTextOf = position => api.PictureAltText(position);
		Editor.KeyDown += (s, e) => { if (recorder.Recording) RecordKeyDown(e); };
		Editor.KeyPress += (s, e) => {
			if (recorder.Recording && (e.KeyChar >= ' ' || e.KeyChar == '\t')) { recorder.Type(e.KeyChar); UpdateRecordingLabel(); }
		};
		Editor.MouseDown += (s, e) => {
			if (recorder.Recording && e.Button == MouseButtons.Left)
				SetStatus("Mouse clicks are not recorded: move with the keyboard (arrows, Home/End, Ctrl to jump, Shift to select).");
		};
		outlineTimer.Tick += (s, e) => { outlineTimer.Stop(); RefreshOutline(); };
		outlineTree.NodeMouseClick += (s, e) => { if (e.Button == MouseButtons.Left && e.Node.Tag is int start) Run(() => GoToPosition(start)); };
		outlineTree.KeyDown += (s, e) => {
			if (e.KeyCode == Keys.Enter && outlineTree.SelectedNode?.Tag is int start) { e.SuppressKeyPress = true; Run(() => GoToPosition(start)); }
		};
		outlineTree.NodeMouseClick += (s, e) => { if (e.Button == MouseButtons.Right) ShowOutlineMenu(e.Node, outlineTree.PointToScreen(e.Location)); };
		textEventTimer.Tick += (s, e) => { textEventTimer.Stop(); Events.Raise(EditorEvent.TextChanged, Editor.TextLength); };
		selectionEventTimer.Tick += (s, e) => {
			selectionEventTimer.Stop();
			HighlightOutline();
			Events.Raise(EditorEvent.SelectionChanged, (uint)Editor.SelectionStart | ((long)Editor.SelectionLength << 32));
		};
		Editor.InsertHyperlinkClicked += (s, e) => InsertHyperlinkDialog();
		Editor.InsertImageClicked += (s, e) => InsertImageDialog();
		Editor.InsertLinkedThumbnailClicked += (s, e) => InsertThumbnailDialog();
		wordCountTimer.Tick += (s, e) => { wordCountTimer.Stop(); UpdateWordCount(); };

		// Drag and drop: on the text, our drop target sits in front of RichEdit's (which would embed a dropped file as an
		// icon); elsewhere in the window, the form handles the drop.
		Editor.HandleCreated += (s, e) => Editor.BeginInvoke(RegisterDropTarget);
		AllowDrop = true;
		DragEnter += (s, e) => e.Effect = e.Data?.GetData(DataFormats.FileDrop) is string[] files && CanDrop(files) ? DragDropEffects.Copy : DragDropEffects.None;
		DragDrop += (s, e) => { if (e.Data?.GetData(DataFormats.FileDrop) is string[] files) DropFiles(files, null); };

		Shown += (s, e) => {
			try { Editor.ZoomFactor = Settings.Zoom; } catch (ArgumentException) { }
			outlineSplit.SplitterDistance = Math.Clamp(LogicalToDeviceUnits(Settings.OutlineWidth), LogicalToDeviceUnits(100), Math.Max(LogicalToDeviceUnits(100), outlineSplit.Width / 2));
			RefreshOutline();
			if (startFile != null) Run(() => api.Open(startFile));
			Editor.Modified = false;
			Editor.Focus();
			UpdateUi();
			UpdateTitle();
			UpdateWordCount();
			SetStatus(startFile != null ? "" : "New document. Tools > Macros runs AHK scripts; Help > Automation API lists what they can call.");
		};
		FormClosing += OnFormClosing;
		FormClosed += (s, e) => {
			Events.Raise(EditorEvent.Closed);
			AppContext.SetData("rte.dispatch", null);
			server.Dispose();
			findForm?.Dispose();
		};
	}

	// ---------------------------------------------------------------- Menus

	MenuStrip BuildMenu() {
		var menu = new MenuStrip();
		menu.Items.Add(Menu("&File",
			Item("&New", () => { if (ConfirmDiscard()) api.New(); }, Keys.Control | Keys.N),
			Item("&Open...", OpenDialog, Keys.Control | Keys.O),
			RecentMenu(),
			Item("&Save", () => SaveCommand(), Keys.Control | Keys.S),
			Item("Save &As...", () => SaveAsDialog(), Keys.Control | Keys.Shift | Keys.S),
			new ToolStripSeparator(),
			Item("Page Set&up...", PageSetupCommand),
			Item("&Header and Footer...", HeaderFooterCommand),
			Item("Print Pre&view...", PrintPreviewCommand),
			Item("&Print...", PrintCommand, Keys.Control | Keys.P),
			Item("&Export as PDF...", ExportPdfDialog),
			Item("Export as &Word (.docx)...", () => ExportOfficeDialog("docx")),
			Item("Export as Open&Document (.odt)...", () => ExportOfficeDialog("odt")),
			new ToolStripSeparator(),
			Item("E&xit", Close)));

		menu.Items.Add(Menu("&Edit",
			Item("&Undo", () => { skipLegalCheck = true; Editor.Undo(); }, Keys.Control | Keys.Z, M("Undo")),
			Item("&Redo", () => { skipLegalCheck = true; Editor.Redo(); }, Keys.Control | Keys.Y, M("Redo")),
			new ToolStripSeparator(),
			Item("Cu&t", Editor.Cut, Keys.Control | Keys.X, M("Cut")),
			Item("&Copy", Editor.Copy, Keys.Control | Keys.C, M("Copy")),
			Item("&Paste", Editor.Paste, Keys.Control | Keys.V, M("Paste")),
			new ToolStripSeparator(),
			Item("Select &All", Editor.SelectAll, Keys.Control | Keys.A, M("SelectAll")),
			new ToolStripSeparator(),
			Item("&Find...", () => ShowFind(false), Keys.Control | Keys.F),
			Item("Find &Next", () => FindNext(), Keys.F3),
			Item("&Replace...", () => ShowFind(true), Keys.Control | Keys.H)));

		var underline = Menu("Underline st&yle");
		foreach (var style in Enum.GetValues<UnderlineStyle>())
			underline.DropDownItems.Add(Item(style.ToString(), () => Editor.ToggleUnderline(style), Keys.None, M("ToggleUnderline", style)));
		menu.Items.Add(Menu("F&ormat",
			Item("&Bold", Editor.ToggleBold, Keys.Control | Keys.B, M("ToggleBold")),
			Item("&Italic", Editor.ToggleItalic, Keys.Control | Keys.I, M("ToggleItalic")),
			Item("&Underline", () => Editor.ToggleUnderline(), Keys.Control | Keys.U, M("ToggleUnderline")),
			underline,
			Item("&Strikethrough", Editor.ToggleStrikeThrough, Keys.None, M("ToggleStrikeThrough")),
			Item("Su&perscript", api.ToggleSuperscript, Keys.Control | Keys.Shift | Keys.Oemplus, M("ToggleSuperscript")),
			Item("Subs&cript", api.ToggleSubscript, Keys.Control | Keys.Oemplus, M("ToggleSubscript")),
			new ToolStripSeparator(),
			Item("&Font...", FontDialogCommand),
			Item("Font c&olor...", () => { if (ChooseColor(fontColor) is Color c) SetFontColor(c); }),
			Item("&Highlight color...", () => { if (ChooseColor(highlightColor) is Color c) SetHighlight(c); }),
			Item("Increase font size", () => { if (FreeFormForSizes()) api.IncreaseFontSize(); }, Keys.Control | Keys.OemCloseBrackets, M("IncreaseFontSize")),
			Item("Decrease font size", () => { if (FreeFormForSizes()) api.DecreaseFontSize(); }, Keys.Control | Keys.OemOpenBrackets, M("DecreaseFontSize")),
			new ToolStripSeparator(),
			Item("UPPER CASE", () => api.ChangeCase("upper"), Keys.None, M("ChangeCase", "upper")),
			Item("lower case", () => api.ChangeCase("lower"), Keys.None, M("ChangeCase", "lower")),
			Item("Title Case", () => api.ChangeCase("title"), Keys.None, M("ChangeCase", "title"))));

		var block = Menu("&Block style");
		foreach (var style in Enum.GetValues<BlockStyle>()) {
			Keys keys = style switch {
				BlockStyle.Paragraph => Keys.Control | Keys.D0,
				BlockStyle.Preformatted => Keys.Control | Keys.D7,
				_ => Keys.Control | (Keys.D0 + (int)style)
			};
			block.DropDownItems.Add(Item(BlockName(style), () => api.SetBlockStyle(style), keys, M("SetBlockStyle", style)));
		}
		var align = Menu("&Alignment");
		foreach (var (alignment, keys) in new[] { (ParagraphAlignment.Left, Keys.L), (ParagraphAlignment.Center, Keys.E), (ParagraphAlignment.Right, Keys.R), (ParagraphAlignment.Justify, Keys.J) })
			align.DropDownItems.Add(Item(alignment.ToString(), () => Editor.SetAlignment(alignment), Keys.Control | keys, M("SetAlignment", alignment)));
		var listStyles = Menu("List st&yle");
		foreach (var style in Enum.GetValues<ListStyle>())
			listStyles.DropDownItems.Add(Item(style.ToString(), () => api.ToggleList(style), Keys.None, M("ToggleList", style)));
		autoLevelItem = new ToolStripMenuItem("Automatic numbering style per level (Paragraph > List levels)") { CheckOnClick = true, Checked = AutoLevelStyles };
		autoLevelItem.CheckedChanged += (s, e) => AutoLevelStyles = autoLevelItem.Checked;
		var increase = Item("Increase list level", () => api.Nest(1), Keys.None, M("Nest", 1));
		increase.ShortcutKeyDisplayString = "Tab";
		var decrease = Item("Decrease list level", () => api.Nest(-1), Keys.None, M("Nest", -1));
		decrease.ShortcutKeyDisplayString = "Shift+Tab";
		menu.Items.Add(Menu("&Paragraph",
			block,
			align,
			new ToolStripSeparator(),
			Item("&Bulleted list", () => api.ToggleList(ListStyle.Bullet), Keys.Control | Keys.Shift | Keys.L, M("ToggleList", ListStyle.Bullet)),
			Item("&Numbered list", () => api.NumberedList(), Keys.Control | Keys.Shift | Keys.N, M("NumberedList")),
			listStyles,
			increase,
			decrease,
			autoLevelItem,
			Item("List le&vels...", ListLevelsCommand),
			Item("Le&gal numbering (1.1.1)", () => api.LegalNumbering(), Keys.Control | Keys.Shift | Keys.G, M("LegalNumbering")),
			Item("Convert &typed numbers to lists", () => {
				int count = api.ConvertTypedNumbers();
				SetStatus(count == 0 ? "No typed list numbers found (paragraphs that start with \"1.\", \"a)\", \"(i)\", \"•\" ... and a tab or space, two or more in a row)."
					: $"{count} paragraph{(count == 1 ? "" : "s")} became list items.");
			}, Keys.None, M("ConvertTypedNumbers")),
			Item("&Restart numbering...", RestartNumberingCommand),
			Item("&Continue numbering", () => api.ContinueNumbering(), Keys.None, M("ContinueNumbering")),
			new ToolStripSeparator(),
			Item("&Indent paragraph", () => Editor.Indent(), Keys.Control | Keys.M, M("Indent")),
			Item("&Outdent paragraph", () => Editor.Outdent(), Keys.Control | Keys.Shift | Keys.M, M("Outdent"))));

		menu.Items.Add(Menu("&Insert",
			Item("&Hyperlink...", InsertHyperlinkDialog, Keys.Control | Keys.K),
			Item("&Image...", InsertImageDialog),
			Item("&Linked thumbnail...", InsertThumbnailDialog),
			new ToolStripSeparator(),
			PageBreakItem(Item("Page &break", () => api.InsertPageBreak(), Keys.None, M("InsertPageBreak"))),
			new ToolStripSeparator(),
			Item("&Date and time", () => { string now = DateTime.Now.ToString("f"); api.InsertText(now); Record(M("InsertText", now)); })));

		var tableCommands = new[] {
			Item("Insert row &above", () => api.InsertRow(false), Keys.None, M("InsertRow", false)),
			Item("Insert row &below", () => api.InsertRow(true), Keys.None, M("InsertRow", true)),
			Item("Insert column &left", () => api.InsertColumn(false), Keys.None, M("InsertColumn", false)),
			Item("Insert column &right", () => api.InsertColumn(true), Keys.None, M("InsertColumn", true)),
			Item("&Delete row", () => api.DeleteRow(), Keys.None, M("DeleteRow")),
			Item("Delete colu&mn", () => api.DeleteColumn(), Keys.None, M("DeleteColumn")),
			Item("Delete &table", () => api.DeleteTable(), Keys.None, M("DeleteTable")),
		};
		tableCommands[1].ShortcutKeyDisplayString = "Ctrl+Shift+Ins";    // handled in ProcessCmdKey, only in tables
		tableCommands[4].ShortcutKeyDisplayString = "Ctrl+Shift+Del";
		var tableMenu = Menu("T&able", Item("&Insert table...", InsertTableDialog), new ToolStripSeparator());
		tableMenu.DropDownItems.AddRange(tableCommands[..4]);
		tableMenu.DropDownItems.Add(new ToolStripSeparator());
		tableMenu.DropDownItems.AddRange(tableCommands[4..]);
		tableMenu.DropDownOpening += (s, e) => {
			bool inTable = api.InTable();
			foreach (var item in tableCommands) item.Enabled = inTable;
		};
		menu.Items.Add(tableMenu);

		structuredItem = Item("&Structured document (headings, block styles)", () => SetStructured(true));
		freeFormItem = Item("&Free-form fonts (any font size)", () => SetStructured(false));
		wordWrapItem = Item("&Word wrap", () => Editor.WordWrap = !Editor.WordWrap);
		spellCheckItem = Item("S&pell check", () => Editor.EnableSpellCheck = !Editor.EnableSpellCheck);
		menu.Items.Add(Menu("&View",
			structuredItem,
			freeFormItem,
			new ToolStripSeparator(),
			wordWrapItem,
			spellCheckItem,
			new ToolStripSeparator(),
			Item("Zoom &in", () => Editor.ZoomFactor = Math.Min(5f, Editor.ZoomFactor + 0.1f)),
			Item("Zoom &out", () => Editor.ZoomFactor = Math.Max(0.5f, Editor.ZoomFactor - 0.1f)),
			Item("Zoom &100%", () => Editor.ZoomFactor = 1f),
			new ToolStripSeparator(),
			bandsItem = Item("Header and &footer", () => api.HeaderFooterVisible = !Settings.ShowBands),
			new ToolStripSeparator(),
			Item("&HTML source", () => ShowText("HTML source", api.GetHtml())),
			Item("&Markdown source", () => ShowText("Markdown source (images shown as image1.png ...)", api.GetMarkdown(false))),
			new ToolStripSeparator(),
			outlineItem = Item("Out&line", () => api.OutlineVisible = outlineSplit.Panel1Collapsed, Keys.Control | Keys.Shift | Keys.O)));

		macrosMenu = Menu("&Macros");
		macrosMenu.DropDownItems.Add(new ToolStripMenuItem("(none)"));    // filled when opened
		macrosMenu.DropDownOpening += (s, e) => FillMacrosMenu();
		recordItem = Item("&Record macro", ToggleRecording, Keys.Control | Keys.Shift | Keys.R);
		menu.Items.Add(Menu("&Tools",
			macrosMenu,
			recordItem,
			Item("Macro &hotkeys...", MacroHotkeysCommand),
			new ToolStripSeparator(),
			Item("Next &spelling error", NextSpellingErrorCommand, Keys.F7),
			Item("&Add selected word to dictionary", () => {
				string word = Editor.SelectedText.Trim();
				if (word == "" || word.Any(char.IsWhiteSpace)) throw new InvalidOperationException("Select a single word first.");
				api.AddToDictionary(word);
				SetStatus($"\"{word}\" was added to the Windows dictionary.");
			}),
			new ToolStripSeparator(),
			Item("&Run AHK script...", RunScriptDialog),
			Item("&Open macros folder", OpenMacrosFolder)));

		menu.Items.Add(Menu("&Help",
			Item("&Automation API", () => ShowText("Automation API (RteClient2.ahk: Rte.<Name>(...))", api.Help())),
			Item("A&bout", () => MessageBox.Show(this, $"{AppName} (iteration 15: pictures' alt text)\n\n{api.Version}\n\nA word processor built on RichEditorNET, hosted in AutoHotkey with Csharp-in-AHK.\n\nConverters for .docx and .odt:\n{OfficeConverters.Describe()}",
				AppName, MessageBoxButtons.OK, MessageBoxIcon.Information))));

		menu.MenuActivate += (s, e) => {
			structuredItem.Checked = Editor.EnableHtmlFontSizing;
			freeFormItem.Checked = !Editor.EnableHtmlFontSizing;
			wordWrapItem.Checked = Editor.WordWrap;
			bandsItem.Checked = Settings.ShowBands;
			spellCheckItem.Checked = Editor.EnableSpellCheck;
			outlineItem.Checked = !outlineSplit.Panel1Collapsed;
		};
		return menu;
	}

	// macro: the call recorded while a macro is being recorded (made with M), or null for commands that are not recorded.
	ToolStripMenuItem Item(string text, Action action, Keys keys = Keys.None, string? macro = null) =>
		new(text, null, (s, e) => Run(action, macro)) { ShortcutKeys = keys };

	static string M(string member, params object?[] args) => MacroRecorder.Call(member, args);

	static ToolStripMenuItem Menu(string text, params ToolStripItem[] items) {
		var item = new ToolStripMenuItem(text);
		item.DropDownItems.AddRange(items);
		return item;
	}

	static string BlockName(BlockStyle style) => style switch {
		BlockStyle.Paragraph => "Normal text",
		BlockStyle.Preformatted => "Preformatted",
		_ => "Heading " + (int)style
	};

	// ---------------------------------------------------------------- Toolbar

	ToolStrip BuildToolbar() {
		int icon = LogicalToDeviceUnits(16);
		var bar = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden, ImageScalingSize = new Size(icon, icon), Padding = new Padding(4, 2, 4, 2) };
		bar.Items.Add(TextButton("New", "New (Ctrl+N)", () => { if (ConfirmDiscard()) api.New(); }));
		bar.Items.Add(TextButton("Open", "Open (Ctrl+O)", OpenDialog));
		bar.Items.Add(TextButton("Save", "Save (Ctrl+S)", () => SaveCommand()));
		bar.Items.Add(TextButton("Print", "Print (Ctrl+P)", PrintCommand));
		bar.Items.Add(TextButton("PDF", "Export as PDF", ExportPdfDialog));
		bar.Items.Add(new ToolStripSeparator());

		blockCombo = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = false, Width = LogicalToDeviceUnits(120), ToolTipText = "Block style (Ctrl+0 to Ctrl+7)" };
		foreach (var style in Enum.GetValues<BlockStyle>()) blockCombo.Items.Add(BlockName(style));
		blockCombo.SelectedIndexChanged += (s, e) => {
			if (updatingUi || blockCombo.SelectedIndex < 0) return;
			Run(() => api.SetBlockStyle((BlockStyle)blockCombo.SelectedIndex), M("SetBlockStyle", (BlockStyle)blockCombo.SelectedIndex));
			Editor.Focus();
		};
		fontCombo = new ToolStripComboBox { AutoSize = false, Width = LogicalToDeviceUnits(150), ToolTipText = "Font",
			AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems };
		using (var fonts = new InstalledFontCollection())
			fontCombo.Items.AddRange(fonts.Families.Select(f => (object)f.Name).ToArray());
		fontCombo.SelectedIndexChanged += (s, e) => { if (!updatingUi && fontCombo.SelectedIndex >= 0) ApplyFontName(); };
		fontCombo.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ApplyFontName(); } };
		sizeCombo = new ToolStripComboBox { AutoSize = false, Width = LogicalToDeviceUnits(52), ToolTipText = "Font size (free-form mode)" };
		foreach (var size in new[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72 }) sizeCombo.Items.Add(size.ToString());
		sizeCombo.SelectedIndexChanged += (s, e) => { if (!updatingUi && sizeCombo.SelectedIndex >= 0) ApplyFontSize(); };
		sizeCombo.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ApplyFontSize(); } };
		bar.Items.AddRange(new ToolStripItem[] { blockCombo, fontCombo, sizeCombo, new ToolStripSeparator() });

		var baseFont = bar.Font;
		Toggle(bar, "bold", "B", new Font(baseFont, FontStyle.Bold), "Bold (Ctrl+B)", Editor.ToggleBold, M("ToggleBold"));
		Toggle(bar, "italic", "I", new Font(baseFont, FontStyle.Italic), "Italic (Ctrl+I)", Editor.ToggleItalic, M("ToggleItalic"));
		Toggle(bar, "underline", "U", new Font(baseFont, FontStyle.Underline), "Underline (Ctrl+U)", () => Editor.ToggleUnderline(), M("ToggleUnderline"));
		Toggle(bar, "strike", "S", new Font(baseFont, FontStyle.Strikeout), "Strikethrough", Editor.ToggleStrikeThrough, M("ToggleStrikeThrough"));
		Toggle(bar, "super", "x²", baseFont, "Superscript (Ctrl+Shift+=)", api.ToggleSuperscript, M("ToggleSuperscript"));
		Toggle(bar, "sub", "x₂", baseFont, "Subscript (Ctrl+=)", api.ToggleSubscript, M("ToggleSubscript"));
		bar.Items.Add(new ToolStripSeparator());

		fontColorButton = ColorButton("A", "Font color", fontColor, c => SetFontColor(c), api.ClearFontColor,
			("Dark red", Color.FromArgb(0xC0, 0, 0)), ("Red", Color.Red), ("Orange", Color.FromArgb(0xE0, 0x70, 0)), ("Green", Color.FromArgb(0, 0x80, 0)),
			("Blue", Color.FromArgb(0, 0x50, 0xC0)), ("Purple", Color.FromArgb(0x70, 0x30, 0xA0)), ("Gray", Color.Gray), ("Black", Color.Black));
		highlightButton = ColorButton("ab", "Highlight", highlightColor, c => SetHighlight(c), api.ClearHighlight,
			("Yellow", Color.Yellow), ("Bright green", Color.Lime), ("Turquoise", Color.Cyan), ("Pink", Color.Pink),
			("Light orange", Color.FromArgb(0xFF, 0xD0, 0x90)), ("Light gray", Color.LightGray));
		bar.Items.AddRange(new ToolStripItem[] { fontColorButton, highlightButton, new ToolStripSeparator() });

		foreach (var alignment in Enum.GetValues<ParagraphAlignment>()) {
			var button = IconButton(AlignIcon(alignment), $"Align {alignment.ToString().ToLowerInvariant()}", () => Editor.SetAlignment(alignment), M("SetAlignment", alignment));
			alignButtons[alignment] = button;
			bar.Items.Add(button);
		}
		bar.Items.Add(new ToolStripSeparator());
		bulletButton = IconButton(ListIcon(false), "Bulleted list (Ctrl+Shift+L)", () => api.ToggleList(ListStyle.Bullet), M("ToggleList", ListStyle.Bullet));
		numberButton = IconButton(ListIcon(true), "Numbered list (Ctrl+Shift+N)", () => api.NumberedList(), M("NumberedList"));
		legalButton = TextButton("1.1", "Legal numbering, 1.1.1 (Ctrl+Shift+G)", () => api.LegalNumbering(), M("LegalNumbering"));
		bar.Items.Add(bulletButton);
		bar.Items.Add(numberButton);
		bar.Items.Add(legalButton);
		bar.Items.Add(IconButton(IndentIcon(false), "Decrease indent / list level (Shift+Tab in a list)", () => { if (api.InList()) { api.Nest(-1); Record(M("Nest", -1)); } else { Editor.Outdent(); Record(M("Outdent")); } }));
		bar.Items.Add(IconButton(IndentIcon(true), "Increase indent / list level (Tab in a list)", () => { if (api.InList()) { api.Nest(1); Record(M("Nest", 1)); } else { Editor.Indent(); Record(M("Indent")); } }));
		bar.Items.Add(new ToolStripSeparator());
		bar.Items.Add(TextButton("Link", "Insert hyperlink (Ctrl+K)", InsertHyperlinkDialog));
		bar.Items.Add(TextButton("Image", "Insert image", InsertImageDialog));
		bar.Items.Add(TextButton("Table", "Insert table", InsertTableDialog));
		return bar;
	}

	ToolStripButton TextButton(string text, string tip, Action action, string? macro = null) =>
		new(text, null, (s, e) => Run(action, macro)) { DisplayStyle = ToolStripItemDisplayStyle.Text, ToolTipText = tip };

	ToolStripButton IconButton(Image image, string tip, Action action, string? macro = null) =>
		new(null, image, (s, e) => Run(action, macro)) { DisplayStyle = ToolStripItemDisplayStyle.Image, ToolTipText = tip };

	void Toggle(ToolStrip bar, string key, string text, Font font, string tip, Action action, string? macro = null) {
		var button = new ToolStripButton(text, null, (s, e) => Run(action, macro)) {
			DisplayStyle = ToolStripItemDisplayStyle.Text, Font = font, ToolTipText = tip, AutoSize = false,
			Width = LogicalToDeviceUnits(26), Height = LogicalToDeviceUnits(24)
		};
		toggles[key] = button;
		bar.Items.Add(button);
	}

	ToolStripSplitButton ColorButton(string letter, string tip, Color initial, Action<Color> apply, Action clear, params (string Name, Color Color)[] palette) {
		var button = new ToolStripSplitButton { DisplayStyle = ToolStripItemDisplayStyle.Image, Image = ColorIcon(letter, initial), ToolTipText = tip };
		button.ButtonClick += (s, e) => Run(() => apply(button == fontColorButton ? fontColor : highlightColor));
		foreach (var (name, color) in palette)
			button.DropDownItems.Add(new ToolStripMenuItem(name, Swatch(color), (s, e) => Run(() => apply(color))));
		button.DropDownItems.Add(new ToolStripSeparator());
		button.DropDownItems.Add(new ToolStripMenuItem(letter == "A" ? "Automatic" : "No highlight", null, (s, e) => Run(clear, M(letter == "A" ? "ClearFontColor" : "ClearHighlight"))));
		button.DropDownItems.Add(new ToolStripMenuItem("More colors...", null, (s, e) => {
			if (ChooseColor(button == fontColorButton ? fontColor : highlightColor) is Color c) Run(() => apply(c));
		}));
		return button;
	}

	void SetFontColor(Color c) {
		Editor.SetFontColor(c);
		Record(M("SetFontColor", c));
		fontColor = c;
		fontColorButton.Image = ColorIcon("A", c);
	}

	void SetHighlight(Color c) {
		Editor.SetBackgroundColor(c);
		Record(M("SetBackgroundColor", c));
		highlightColor = c;
		highlightButton.Image = ColorIcon("ab", c);
	}

	void ApplyFontName() {
		string name = fontCombo.Text.Trim();
		if (name == "") return;
		Run(() => Editor.SetFontName(name), M("SetFontName", name));
		Editor.Focus();
	}

	void ApplyFontSize() {
		if (!float.TryParse(sizeCombo.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out float size)) return;
		if (FreeFormForSizes()) Run(() => api.SetFontSize(size), M("SetFontSize", size));
		Editor.Focus();
	}

	// Font sizes need free-form mode (structured documents have block styles instead): asks to switch to it.
	bool FreeFormForSizes() {
		if (Editor.EnableFontSize) return true;
		var answer = MessageBox.Show(this, "Font sizes can be changed in free-form mode only; structured documents (headings, block styles) set sizes by block style.\n\nSwitch to free-form fonts (View > Free-form fonts)?",
			AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
		if (answer != DialogResult.Yes) { UpdateUi(); return false; }
		SetStructured(false);
		return true;
	}

	// ---------------------------------------------------------------- Icons (drawn on a 16x16 grid, scaled for DPI)

	Bitmap Icon(Action<Graphics> draw) {
		int size = LogicalToDeviceUnits(16);
		var bitmap = new Bitmap(size, size);
		using var g = Graphics.FromImage(bitmap);
		g.SmoothingMode = SmoothingMode.AntiAlias;
		g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
		g.ScaleTransform(size / 16f, size / 16f);
		draw(g);
		return bitmap;
	}

	static readonly Color IconColor = Color.FromArgb(50, 50, 50);

	Bitmap AlignIcon(ParagraphAlignment alignment) => Icon(g => {
		using var pen = new Pen(IconColor, 1.5f);
		float[] widths = alignment == ParagraphAlignment.Justify ? new[] { 14f, 14f, 14f, 14f } : new[] { 14f, 9f, 14f, 9f };
		for (int i = 0; i < 4; i++) {
			float w = widths[i], y = 2.5f + i * 3.6f;
			float x = alignment switch { ParagraphAlignment.Center => (16 - w) / 2, ParagraphAlignment.Right => 15 - w, _ => 1 };
			g.DrawLine(pen, x, y, x + w, y);
		}
	});

	Bitmap ListIcon(bool numbered) => Icon(g => {
		using var pen = new Pen(IconColor, 1.5f);
		using var brush = new SolidBrush(IconColor);
		using var font = new Font("Segoe UI", 5.5f, FontStyle.Bold, GraphicsUnit.Pixel);
		for (int i = 0; i < 3; i++) {
			float y = 3f + i * 5f;
			if (numbered) g.DrawString((i + 1).ToString(), font, brush, -0.5f, y - 3.8f);
			else g.FillEllipse(brush, 1.2f, y - 1.4f, 2.8f, 2.8f);
			g.DrawLine(pen, 6, y, 15, y);
		}
	});

	Bitmap IndentIcon(bool increase) => Icon(g => {
		using var pen = new Pen(IconColor, 1.5f);
		using var brush = new SolidBrush(Color.FromArgb(0, 0x60, 0xC0));
		g.DrawLine(pen, 1, 2.5f, 15, 2.5f);
		g.DrawLine(pen, 7, 6.2f, 15, 6.2f);
		g.DrawLine(pen, 7, 9.8f, 15, 9.8f);
		g.DrawLine(pen, 1, 13.5f, 15, 13.5f);
		g.FillPolygon(brush, increase
			? new[] { new PointF(1, 5), new PointF(1, 11), new PointF(5, 8) }
			: new[] { new PointF(5, 5), new PointF(5, 11), new PointF(1, 8) });
	});

	Bitmap ColorIcon(string letter, Color color) => Icon(g => {
		using var font = new Font("Segoe UI", letter.Length > 1 ? 8.5f : 10.5f, FontStyle.Bold, GraphicsUnit.Pixel);
		using var text = new SolidBrush(IconColor);
		var size = g.MeasureString(letter, font);
		g.DrawString(letter, font, text, (16 - size.Width) / 2, -1.5f);
		using var bar = new SolidBrush(color);
		g.FillRectangle(bar, 1, 12, 14, 3.5f);
	});

	Bitmap Swatch(Color color) => Icon(g => {
		using var brush = new SolidBrush(color);
		g.FillRectangle(brush, 1.5f, 1.5f, 13, 13);
		g.DrawRectangle(Pens.Gray, 1.5f, 1.5f, 13, 13);
	});

	// ---------------------------------------------------------------- Commands

	// Runs a UI command. Errors (e.g. a feature that is unavailable in the current mode) go to the status bar.
	void Run(Action action, string? macro = null) {
		try {
			action();
			if (macro != null) Record(macro);
		} catch (Exception ex) {
			if (ex is TargetInvocationException { InnerException: { } inner }) ex = inner;
			SetStatus("⚠ " + ex.Message);
			System.Media.SystemSounds.Beep.Play();
		}
		UpdateUi();
		UpdateTitle();
		Restart(outlineTimer);
	}

	public void SetStatus(string text) => statusText.Text = text;

	static void Restart(System.Windows.Forms.Timer timer) {
		timer.Stop();
		timer.Start();
	}

	internal void SetFileName(string? path) {
		FileName = path;
		UpdateTitle();
	}

	void SetStructured(bool structured) {
		Editor.EnableHtmlFontSizing = structured;
		SetStatus(structured ? "Structured document: block styles (Ctrl+0..7) instead of font sizes." : "Free-form fonts: any font size; block styles are off.");
	}

	bool ConfirmDiscard() {
		if (!Editor.Modified) return true;
		var answer = MessageBox.Show(this, $"Save changes to {DocumentName}?", AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
		return answer == DialogResult.No || (answer == DialogResult.Yes && SaveCommand());
	}

	string DocumentName => FileName != null ? Path.GetFileName(FileName) : "Untitled";

	// File types for the Open, Save As and Export dialogs. Word and OpenDocument have one entry per converter as well as
	// the automatic one (Via ""), so that a particular converter can be chosen.
	sealed record FileType(string Label, string Patterns, string Via = "");

	static readonly FileType[] OpenTypes = {
		new("Documents", "*.html;*.htm;*.md;*.markdown;*.rtf;*.docx;*.doc;*.odt;*.txt"),
		new("HTML", "*.html;*.htm"), new("Markdown", "*.md;*.markdown"), new("Rich Text Format", "*.rtf"),
		new("Word", "*.docx;*.doc"),
		new("Word, built-in converter", "*.docx", OfficeConverters.OpenXml), new("Word, converted by Word", "*.docx;*.doc", OfficeConverters.Word),
		new("Word, converted by LibreOffice", "*.docx;*.doc", OfficeConverters.LibreOffice), new("Word, converted by pandoc", "*.docx", OfficeConverters.Pandoc),
		new("OpenDocument text", "*.odt"),
		new("OpenDocument, converted by LibreOffice", "*.odt", OfficeConverters.LibreOffice), new("OpenDocument, converted by Word", "*.odt", OfficeConverters.Word),
		new("OpenDocument, converted by pandoc", "*.odt", OfficeConverters.Pandoc),
		new("Text", "*.txt"), new("All files", "*.*")
	};

	static readonly FileType[] WordTypes = {
		new("Word", "*.docx"), new("Word, built-in converter", "*.docx", OfficeConverters.OpenXml),
		new("Word, converted by Word", "*.docx", OfficeConverters.Word), new("Word, converted by LibreOffice", "*.docx", OfficeConverters.LibreOffice),
		new("Word, converted by pandoc", "*.docx", OfficeConverters.Pandoc)
	};

	static readonly FileType[] OdtTypes = {
		new("OpenDocument text", "*.odt"), new("OpenDocument, converted by LibreOffice", "*.odt", OfficeConverters.LibreOffice),
		new("OpenDocument, converted by Word", "*.odt", OfficeConverters.Word), new("OpenDocument, converted by pandoc", "*.odt", OfficeConverters.Pandoc)
	};

	static readonly FileType[] SaveTypes = new[] { new FileType("HTML", "*.html"), new FileType("Markdown", "*.md"), new FileType("Rich Text Format", "*.rtf") }
		.Concat(WordTypes).Concat(OdtTypes).Append(new FileType("Text", "*.txt")).ToArray();

	static string Filter(IEnumerable<FileType> types) => string.Join("|", types.Select(t => $"{t.Label} ({t.Patterns.Replace(";*.htm;", ";").Replace(";*.markdown", "")})|{t.Patterns}"));

	// The dialog's FilterIndex (from 1) for a file name and converter: the first type with its extension and converter.
	static int FilterIndex(FileType[] types, string? fileName, string via) {
		string extension = Path.GetExtension(fileName ?? "").ToLowerInvariant();
		int index = Array.FindIndex(types, t => t.Patterns.Split(';').Contains("*" + extension) && t.Via == via);
		if (index < 0) index = Array.FindIndex(types, t => t.Patterns.Split(';').Contains("*" + extension));
		return Math.Max(0, index) + 1;
	}

	static string ViaOf(FileType[] types, int filterIndex) => filterIndex >= 1 && filterIndex <= types.Length ? types[filterIndex - 1].Via : "";

	void OpenDialog() {
		if (!ConfirmDiscard()) return;
		using var dlg = new OpenFileDialog { Title = "Open", Filter = Filter(OpenTypes), InitialDirectory = StartFolder };
		if (dlg.ShowDialog(this) == DialogResult.OK) {
			Cursor = Cursors.WaitCursor;
			try { api.Open(dlg.FileName, ViaOf(OpenTypes, dlg.FilterIndex)); }
			finally { Cursor = Cursors.Default; }
			SetStatus("Opened " + dlg.FileName + (api.LastConverter != "" ? " (converted by " + api.LastConverter + ")" : ""));
		}
	}

	bool SaveCommand() {
		if (FileName == null) return SaveAsDialog();
		try {
			api.Save();
			SetStatus("Saved " + FileName);
			return true;
		} catch (Exception ex) {
			MessageBox.Show(this, ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
			return false;
		}
	}

	bool SaveAsDialog() {
		using var dlg = new SaveFileDialog {
			Title = "Save As", Filter = Filter(SaveTypes), FilterIndex = FilterIndex(SaveTypes, FileName ?? "Untitled.html", api.FileConverter),
			FileName = FileName != null ? Path.GetFileName(FileName) : "Untitled.html",
			InitialDirectory = FileName != null ? Path.GetDirectoryName(FileName) : StartFolder
		};
		if (dlg.ShowDialog(this) != DialogResult.OK) return false;
		Cursor = Cursors.WaitCursor;
		try {
			api.SaveAs(dlg.FileName, ViaOf(SaveTypes, dlg.FilterIndex));
			SetStatus("Saved " + dlg.FileName + (api.LastConverter != "" ? " (converted by " + api.LastConverter + ")" : ""));
			return true;
		} catch (Exception ex) {
			MessageBox.Show(this, ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
			return false;
		} finally {
			Cursor = Cursors.Default;
		}
	}

	string StartFolder => Settings.LastFolder is { } folder && Directory.Exists(folder) ? folder : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

	// ---------------------------------------------------------------- Iteration 3: settings and recent files

	void ApplyWindowSettings() {
		if (Settings.Window is { Length: 4 } w && w[2] >= 400 && w[3] >= 300) {
			var bounds = new Rectangle(w[0], w[1], w[2], w[3]);
			if (Screen.AllScreens.Any(screen => screen.WorkingArea.IntersectsWith(bounds))) {
				StartPosition = FormStartPosition.Manual;
				Bounds = bounds;
			}
		}
		if (Settings.Maximized) WindowState = FormWindowState.Maximized;
	}

	void SaveSettings() {
		var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
		Settings.Window = new[] { bounds.X, bounds.Y, bounds.Width, bounds.Height };
		Settings.Maximized = WindowState == FormWindowState.Maximized;
		Settings.Structured = Editor.EnableHtmlFontSizing;
		Settings.WordWrap = Editor.WordWrap;
		Settings.SpellCheck = Editor.EnableSpellCheck;
		Settings.AutoLevelStyles = AutoLevelStyles;
		Settings.ShowOutline = !outlineSplit.Panel1Collapsed;
		if (!outlineSplit.Panel1Collapsed) Settings.OutlineWidth = (int)Math.Round(outlineSplit.SplitterDistance * 96.0 / DeviceDpi);
		try { Settings.Zoom = Editor.ZoomFactor; } catch (ObjectDisposedException) { }
		Settings.Save();
	}

	// Called whenever a document is opened or saved.
	internal void AddRecent(string path) {
		Settings.RecentFiles.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
		Settings.RecentFiles.Insert(0, path);
		if (Settings.RecentFiles.Count > Settings.MaxRecentFiles) Settings.RecentFiles.RemoveRange(Settings.MaxRecentFiles, Settings.RecentFiles.Count - Settings.MaxRecentFiles);
		Settings.LastFolder = Path.GetDirectoryName(path);
		Settings.Save();
	}

	ToolStripMenuItem RecentMenu() {
		recentMenu = new ToolStripMenuItem("Open &Recent");
		recentMenu.DropDownItems.Add(new ToolStripMenuItem("(none)"));    // filled when opened
		recentMenu.DropDownOpening += (s, e) => FillRecentMenu();
		return recentMenu;
	}

	void FillRecentMenu() {
		recentMenu.DropDownItems.Clear();
		int number = 0;
		foreach (var path in Settings.RecentFiles.ToList()) {
			number++;
			string label = $"{(number < 10 ? "&" + number : "1&0")}  {Path.GetFileName(path)}";
			recentMenu.DropDownItems.Add(new ToolStripMenuItem(label, null, (s, e) => Run(() => OpenRecent(path))) { ToolTipText = path });
		}
		if (number == 0) recentMenu.DropDownItems.Add(new ToolStripMenuItem("(no recent files)") { Enabled = false });
		recentMenu.DropDownItems.Add(new ToolStripSeparator());
		recentMenu.DropDownItems.Add(new ToolStripMenuItem("&Clear list", null, (s, e) => { Settings.RecentFiles.Clear(); Settings.Save(); }) { Enabled = number > 0 });
	}

	void OpenRecent(string path) {
		if (!File.Exists(path)) {
			Settings.RecentFiles.Remove(path);
			Settings.Save();
			throw new FileNotFoundException($"{path} no longer exists, so it was removed from the list.");
		}
		if (!ConfirmDiscard()) return;
		api.Open(path);
		SetStatus("Opened " + path);
	}

	// ---------------------------------------------------------------- Iteration 3: drag and drop

	static readonly string[] DocumentExtensions = { ".html", ".htm", ".md", ".markdown", ".rtf", ".docx", ".doc", ".odt", ".txt" };
	static readonly string[] ImageExtensions = { ".png", ".jpg", ".jpeg", ".gif" };
	static bool IsDocument(string path) => DocumentExtensions.Contains(Path.GetExtension(path).ToLowerInvariant());
	static bool IsImage(string path) => ImageExtensions.Contains(Path.GetExtension(path).ToLowerInvariant());
	static bool CanDrop(string[] files) => files.Any(f => IsDocument(f) || IsImage(f));

	internal string DropTargetStatus { get; private set; } = "not registered";

	void RegisterDropTarget() {
		if (!Editor.IsHandleCreated) return;
		dropTarget = FileDropTarget.Register(Editor.Handle, CanDrop, DropFiles, out string status);
		DropTargetStatus = status;
	}

	// A dropped document is opened; dropped images are inserted at the drop point (screen coordinates), or at the caret.
	// This runs after the drop has completed, so a "save changes?" question does not hold up the program dragged from.
	internal void DropFiles(string[] files, Point? screenPoint) {
		BeginInvoke(() => Run(() => {
			if (files.FirstOrDefault(IsDocument) is { } document) {
				if (!ConfirmDiscard()) return;
				api.Open(document);
				SetStatus("Opened " + document);
				return;
			}
			if (screenPoint is { } point) Editor.Select(Editor.GetCharIndexFromPosition(Editor.PointToClient(point)), 0);
			int count = 0;
			foreach (var image in files.Where(IsImage)) {
				api.InsertImageFile(image, Path.GetFileNameWithoutExtension(image));
				count++;
			}
			SetStatus($"Inserted {count} image{(count == 1 ? "" : "s")}.");
		}));
	}

	// ---------------------------------------------------------------- Iteration 3: printing and PDF

	internal DocumentPrinter CreatePrinter(string? printerName = null) => new(Editor, DocumentName, FileName ?? "", Settings, Bands, printerName);

	/// The document's header and footer.
	internal PageBands Bands = new();

	// ---- the header and footer shown above and below the text (iteration 14)

	readonly BandStrip headerStrip = new(true), footerStrip = new(false);
	readonly System.Windows.Forms.Timer bandTimer = new() { Interval = 500 };
	int? bandPages;                 // the number of pages, for {pages}; null: not known
	bool bandPagesStale = true;     // the text changed since it was counted
	DateTime lastEdit = DateTime.MinValue;
	string bandsShown = "";

	// Shows the strips as the document's header and footer are now (twice a second: cheap unless something changed). The
	// number of pages is counted again (with the printer's layout) once typing has paused for a second.
	void UpdateBands() {
		bool visible = Settings.ShowBands;
		var bands = Bands;
		bool needPages = visible && ((bands.ShowHeader && PageBands.Parts(bands.Header).Any(p => p.Contains("{pages}", StringComparison.OrdinalIgnoreCase)))
			|| (bands.ShowFooter && PageBands.Parts(bands.Footer).Any(p => p.Contains("{pages}", StringComparison.OrdinalIgnoreCase))));
		if (needPages && bandPagesStale && (DateTime.Now - lastEdit).TotalMilliseconds > 1000) {
			bandPagesStale = false;
			try { using var printer = CreatePrinter(); bandPages = printer.PageCount(); }
			catch (Exception) { bandPages = null; }
		}
		string Fill(string text) => Regex.Replace(text, @"\{(title|file|page|pages|date|time)\}", m => m.Groups[1].Value.ToLowerInvariant() switch {
			"title" => DocumentName,
			"file" => FileName ?? DocumentName,
			"page" => "1",
			"pages" => bandPages?.ToString(CultureInfo.CurrentCulture) ?? "?",
			"date" => DateTime.Now.ToShortDateString(),
			_ => DateTime.Now.ToShortTimeString()
		}, RegexOptions.IgnoreCase);
		var header = PageBands.Parts(bands.Header).Select(Fill).ToArray();
		var footer = PageBands.Parts(bands.Footer).Select(Fill).ToArray();
		string shown = $"{visible}|{bands.ShowHeader}|{string.Join("\t", header)}|{bands.ShowFooter}|{string.Join("\t", footer)}";
		if (shown == bandsShown) return;
		bandsShown = shown;
		headerStrip.Display(visible && bands.ShowHeader, header);
		footerStrip.Display(visible && bands.ShowFooter, footer);
	}    // (the default, once the settings are loaded: in the constructor)

	/// The header and footer new documents get (File > Header and Footer can make a document's the default).
	internal PageBands DefaultBands() => new() {
		ShowHeader = Settings.ShowHeader, Header = Settings.Parts(Settings.Header), ShowFooter = Settings.ShowFooter, Footer = Settings.Parts(Settings.Footer)
	};

	internal void MakeDefault(PageBands bands) {
		(Settings.ShowHeader, Settings.Header, Settings.ShowFooter, Settings.Footer) = (bands.ShowHeader, Settings.Parts(bands.Header), bands.ShowFooter, Settings.Parts(bands.Footer));
		Settings.Save();
	}

	/// The page that other programs lay the document out on (twips): the default printer's paper (Letter or A4 if there
	/// is no printer) and Page Setup's margins and orientation.
	internal PageLayout PageLayout() {
		int width = 850, height = 1100;    // 1/100 inch
		try {
			using var printer = CreatePrinter();
			var paper = printer.Document.DefaultPageSettings.PaperSize;
			(width, height) = (paper.Width, paper.Height);
		} catch (Exception) {
			if (RegionInfo.CurrentRegion.IsMetric) (width, height) = (827, 1169);
		}
		if (Settings.Landscape) (width, height) = (height, width);
		int[] m = Settings.Margins is { Length: 4 } margins ? margins : new[] { 100, 100, 100, 100 };
		static int Twips(int hundredths) => (int)Math.Round(hundredths * 14.4);
		return new PageLayout(Twips(width), Twips(height), Twips(m[0]), Twips(m[1]), Twips(m[2]), Twips(m[3]), Settings.Landscape);
	}

	void HeaderFooterCommand() {
		using var dlg = new HeaderFooterDialog(Bands, DefaultBands());
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		var bands = dlg.Result();
		if (!bands.SameAs(Bands)) {
			Bands = bands;
			Editor.Modified = true;    // (saved with the document)
		}
		if (dlg.MakeDefault) MakeDefault(bands);
		SetStatus((dlg.MakeDefault ? "Header and footer set for this document and new ones." : "Header and footer set for this document.")
			+ " File > Print Preview shows them; Word and OpenDocument files get them too.");
	}

	void PageSetupCommand() {
		using var printer = CreatePrinter();
		using var dlg = new PageSetupDialog { Document = printer.Document, EnableMetric = true };
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		var page = printer.Document.DefaultPageSettings;
		Settings.Margins = new[] { page.Margins.Left, page.Margins.Right, page.Margins.Top, page.Margins.Bottom };
		Settings.Landscape = page.Landscape;
		Settings.Save();
	}

	void PrintPreviewCommand() {
		using var printer = CreatePrinter();
		using var dlg = new PrintPreviewDialog { Document = printer.Document, ShowIcon = false,
			Size = new Size(LogicalToDeviceUnits(900), LogicalToDeviceUnits(1000)), StartPosition = FormStartPosition.CenterParent };
		dlg.ShowDialog(this);
	}

	void PrintCommand() {
		using var printer = CreatePrinter();
		using var dlg = new PrintDialog { Document = printer.Document, UseEXDialog = true };
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		printer.Document.Print();
		SetStatus("Sent to " + printer.Document.PrinterSettings.PrinterName);
	}

	void ExportPdfDialog() {
		using var dlg = new SaveFileDialog {
			Title = "Export as PDF", Filter = "PDF (*.pdf)|*.pdf",
			FileName = Path.ChangeExtension(DocumentName, ".pdf"),
			InitialDirectory = FileName != null ? Path.GetDirectoryName(FileName) : StartFolder
		};
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Cursor = Cursors.WaitCursor;
		try {
			SetStatus("Exported " + api.ExportPdf(dlg.FileName));
		} finally {
			Cursor = Cursors.Default;
		}
	}

	// ---------------------------------------------------------------- Links

	// Ctrl+click on a link opens it: web and mail addresses at once, anything else (a file or a program could be
	// behind it) after asking.
	void OpenLinkCommand(string url) {
		bool web = Regex.IsMatch(url, @"^(https?|mailto):", RegexOptions.IgnoreCase);
		if (!web && MessageBox.Show(this, $"Open this link?\n\n{url}", AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
		try {
			System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true })?.Dispose();
			SetStatus("Opened " + url);
		} catch (Exception ex) {
			SetStatus("⚠ Could not open " + url + ": " + ex.Message);
		}
	}

	// ---------------------------------------------------------------- Iteration 10: legal numbering

	// The caret never stays inside a legal item's number (typing there would break it): it goes after the number, or,
	// moving left from the start of the text, to the end of the paragraph before.
	void GuardLegalCaret() {
		if (guardingCaret || !Editor.IsHandleCreated) return;
		var selection = Editor.TextDocument.Selection2;
		int caret = selection.Start;
		if (selection.End != caret) { lastCaret = -1; return; }
		var para = Editor.TextDocument.Range2(caret, caret);
		para.Expand(tomConstants.tomParagraph);
		if (LegalLists.IsLegal(para.Para.ListType)) {
			int textStart = para.Start + LegalLists.Prefix.Match(para.Text ?? "").Length;
			if (caret < textStart) {
				int target = lastCaret == textStart && para.Start > 0 ? para.Start - 1 : textStart;
				guardingCaret = true;
				try { selection.SetRange(target, target); } finally { guardingCaret = false; }
				caret = target;
			}
		}
		lastCaret = caret;
	}

	// After changes the commands don't cover (pasting, dropping, RichEdit's own editing), legal numbers are put right, as
	// an undo step of their own. Not after Undo / Redo, which restore numbers along with the text.
	void CheckLegalNumbers() {
		if (skipLegalCheck) { skipLegalCheck = false; return; }
		try {
			if (LegalLists.AnyLegal(Editor.TextDocument)) LegalLists.Renumber(Editor);
		} catch (Exception ex) { SetStatus("⚠ " + ex.Message); }
	}

	// ---------------------------------------------------------------- Iteration 9: Word, OpenDocument and tables

	// Writes a copy of the document as .docx or .odt; the current file stays the one being edited.
	void ExportOfficeDialog(string format) {
		var types = format == "docx" ? WordTypes : OdtTypes;
		using var dlg = new SaveFileDialog {
			Title = format == "docx" ? "Export as Word" : "Export as OpenDocument",
			Filter = Filter(types),
			FileName = Path.ChangeExtension(DocumentName, "." + format),
			InitialDirectory = FileName != null ? Path.GetDirectoryName(FileName) : StartFolder
		};
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Cursor = Cursors.WaitCursor;
		SetStatus("Converting...");
		try {
			string via = ViaOf(types, dlg.FilterIndex);
			string path = format == "docx" ? api.ExportDocx(dlg.FileName, via) : api.ExportOdt(dlg.FileName, via);
			SetStatus($"Exported {path} (converted by {api.LastConverter})");
		} finally {
			Cursor = Cursors.Default;
		}
	}

	void InsertTableDialog() {
		using var dlg = new InsertTableDialog(Settings.TableRows, Settings.TableColumns, Settings.TableHeader);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		(Settings.TableRows, Settings.TableColumns, Settings.TableHeader) = (dlg.Rows, dlg.Columns, dlg.Header);
		Settings.Save();
		Run(() => api.InsertTable(dlg.Rows, dlg.Columns, dlg.Header), M("InsertTable", dlg.Rows, dlg.Columns, dlg.Header));
	}

	// ---------------------------------------------------------------- Iteration 5: spelling

	// The right-click menu: spelling suggestions on a misspelled word, and the table commands in a table (which then act
	// on the cell clicked). Returns false (the library then shows its formatting toolbar) when neither applies.
	bool ShowContextMenu(Point screenPoint, int position) {
		// a picture: selected (as in Word), with Alt text... in the menu
		int picture = Editor.PictureAt(screenPoint);
		var selected = Editor.TextDocument.Selection2;
		if (picture < 0 && selected.End == selected.Start + 1 && selected.Text == "\uFFFC") picture = selected.Start;
		if (picture >= 0) {
			selected.SetRange(picture, picture + 1);
			var pictureMenu = new ContextMenuStrip { ShowImageMargin = false };
			pictureMenu.Items.Add(new ToolStripMenuItem("Cu&t", null, (s, e) => Run(Editor.Cut, M("Cut"))));
			pictureMenu.Items.Add(new ToolStripMenuItem("&Copy", null, (s, e) => Run(Editor.Copy, M("Copy"))));
			pictureMenu.Items.Add(new ToolStripMenuItem("&Paste", null, (s, e) => Run(Editor.Paste, M("Paste"))));
			pictureMenu.Items.Add(new ToolStripSeparator());
			pictureMenu.Items.Add(new ToolStripMenuItem("&Alt text...", null, (s, e) => Run(() => AltTextCommand(picture))));
			pictureMenu.Closed += (s, e) => BeginInvoke(pictureMenu.Dispose);
			pictureMenu.Show(screenPoint);
			return true;
		}
		bool inTable = RtfTables.At(Editor, position) != null;
		var error = Editor.EnableSpellCheck && Spelling != null ? api.SpellingErrorAt(position) : null;
		if (!inTable) return error != null && ShowSpellingMenu(screenPoint, error);
		var tom = Editor.TextDocument.Selection2;    // (WinForms' SelectionLength is short across table rows)
		// a click outside the selection moves the caret there, as in Word; in the cell the caret is in, it stays
		var clicked = RtfTables.At(Editor, position);
		var current = RtfTables.At(Editor, tom.Start);
		bool sameCell = clicked != null && current != null && clicked.Value.Table.Start == current.Value.Table.Start
			&& clicked.Value.Row == current.Value.Row && clicked.Value.Cell == current.Value.Cell;
		if ((position < tom.Start || position > tom.End) && !sameCell) tom.SetRange(position, position);
		var menu = new ContextMenuStrip { ShowImageMargin = false };
		if (error != null) {
			AddSpellingItems(menu, error);
			menu.Items.Add(new ToolStripSeparator());
		}
		bool selection = tom.End > tom.Start;
		menu.Items.Add(new ToolStripMenuItem("Cu&t", null, (s, e) => Run(Editor.Cut, M("Cut"))) { Enabled = selection });
		menu.Items.Add(new ToolStripMenuItem("&Copy", null, (s, e) => Run(Editor.Copy, M("Copy"))) { Enabled = selection });
		menu.Items.Add(new ToolStripMenuItem("&Paste", null, (s, e) => Run(Editor.Paste, M("Paste"))));
		menu.Items.Add(new ToolStripSeparator());
		void Add(string text, Action action, string macro, string keys = "") =>
			menu.Items.Add(new ToolStripMenuItem(text, null, (s, e) => Run(action, macro)) { ShortcutKeyDisplayString = keys });
		Add("Insert row &above", () => api.InsertRow(false), M("InsertRow", false));
		Add("Insert row &below", () => api.InsertRow(true), M("InsertRow", true), "Ctrl+Shift+Ins");
		Add("Insert column &left", () => api.InsertColumn(false), M("InsertColumn", false));
		Add("Insert column &right", () => api.InsertColumn(true), M("InsertColumn", true));
		menu.Items.Add(new ToolStripSeparator());
		int selectedRows = api.SelectedRows();
		if (selectedRows > 1) Add($"&Delete {selectedRows} rows", () => api.DeleteSelectedRows(), M("DeleteSelectedRows"), "Del");
		else Add("&Delete row", () => api.DeleteRow(), M("DeleteRow"), "Ctrl+Shift+Del");
		Add("Delete colu&mn", () => api.DeleteColumn(), M("DeleteColumn"));
		Add("Delete &table", () => api.DeleteTable(), M("DeleteTable"));
		menu.Closed += (s, e) => BeginInvoke(menu.Dispose);
		menu.Show(screenPoint);
		return true;
	}

	// The right-click menu on a misspelled word: suggestions, "delete repeated word" and "add to dictionary".
	bool ShowSpellingMenu(Point screenPoint, int position) =>
		Editor.EnableSpellCheck && Spelling != null && api.SpellingErrorAt(position) is { } error && ShowSpellingMenu(screenPoint, error);

	bool ShowSpellingMenu(Point screenPoint, Spelling.Error error) {
		var menu = new ContextMenuStrip { ShowImageMargin = false };
		AddSpellingItems(menu, error);
		menu.Closed += (s, e) => BeginInvoke(menu.Dispose);
		menu.Show(screenPoint);
		return true;
	}

	void AddSpellingItems(ContextMenuStrip menu, Spelling.Error error) {
		var bold = new Font(menu.Font, FontStyle.Bold);
		menu.Disposed += (s, e) => bold.Dispose();
		if (error.Action == Spelling.DeleteAction) {
			menu.Items.Add(new ToolStripMenuItem($"Delete repeated word \"{error.Word}\"", null, (s, e) => Run(() => api.DeleteRepeatedWord(error))) { Font = bold });
		} else {
			var suggestions = error.Action == Spelling.ReplaceAction && error.Replacement != "" ? new List<string> { error.Replacement } : Spelling.Suggest(error.Word);
			foreach (var suggestion in suggestions)
				menu.Items.Add(new ToolStripMenuItem(suggestion, null, (s, e) => Run(() => api.ReplaceRange(error.Start, error.Length, suggestion))) { Font = bold });
			if (suggestions.Count == 0) menu.Items.Add(new ToolStripMenuItem("(no suggestions)") { Enabled = false });
			menu.Items.Add(new ToolStripSeparator());
			menu.Items.Add(new ToolStripMenuItem($"Add \"{error.Word}\" to dictionary", null, (s, e) => Run(() => {
				api.AddToDictionary(error.Word);
				SetStatus($"\"{error.Word}\" was added to the Windows dictionary.");
			})));
		}
	}

	void NextSpellingErrorCommand() {
		if (Spelling == null) throw new InvalidOperationException("The Windows spell checker is not available.");
		if (api.NextSpellingError() == "") {
			SetStatus("No spelling errors found.");
			return;
		}
		// Open the suggestions just below the word.
		var point = Editor.GetPositionFromCharIndex(Editor.SelectionStart);
		point.Offset(0, Editor.Font.Height + LogicalToDeviceUnits(4));
		ShowSpellingMenu(Editor.PointToScreen(point), Editor.SelectionStart);
	}

	// ---------------------------------------------------------------- Iteration 5: outline

	internal void SetOutlineVisible(bool visible) {
		outlineSplit.Panel1Collapsed = !visible;
		if (visible) RefreshOutline();
	}

	internal bool OutlineVisible => !outlineSplit.Panel1Collapsed;

	internal List<(int Level, int Start, string Text)> CurrentOutline() {
		RefreshOutline();
		return outline;
	}

	// Rebuilds the tree when the headings have changed.
	void RefreshOutline() {
		if (!Editor.IsHandleCreated) return;
		List<(int Level, int Start, string Text)> headings;
		try {
			headings = api.Headings();
		} catch (COMException) {
			return;    // the document is being replaced; the next change refreshes again
		}
		if (outlineBuilt && headings.SequenceEqual(outline)) {
			HighlightOutline();
			return;
		}
		outline = headings;
		updatingOutline = true;
		outlineTree.BeginUpdate();
		try {
			outlineTree.Nodes.Clear();
			var parents = new TreeNode?[7];
			foreach (var (level, start, text) in headings) {
				var node = new TreeNode(text.Length > 80 ? text[..80] + "…" : text) { Tag = start, ToolTipText = text };
				TreeNode? parent = null;
				for (int l = level - 1; l >= 1 && parent == null; l--) parent = parents[l];
				(parent?.Nodes ?? outlineTree.Nodes).Add(node);
				parents[level] = node;
				for (int l = level + 1; l <= 6; l++) parents[l] = null;
			}
			outlineTree.ExpandAll();
		} finally {
			outlineTree.EndUpdate();
			updatingOutline = false;
		}
		outlineBuilt = true;
		outlineHint.Visible = headings.Count == 0;
		outlineTree.Visible = headings.Count > 0;
		HighlightOutline();
	}

	// Selects the heading of the section the caret is in.
	void HighlightOutline() {
		if (updatingOutline || outlineSplit.Panel1Collapsed || outline.Count == 0) return;
		int caret = Editor.SelectionStart;
		TreeNode? best = null;
		foreach (var node in AllNodes(outlineTree.Nodes))
			if (node.Tag is int start && start <= caret && (best == null || start >= (int)best.Tag!)) best = node;
		if (best != null && outlineTree.SelectedNode != best) outlineTree.SelectedNode = best;
		if (best == null) outlineTree.SelectedNode = null;
	}

	static IEnumerable<TreeNode> AllNodes(TreeNodeCollection nodes) {
		foreach (TreeNode node in nodes) {
			yield return node;
			foreach (var child in AllNodes(node.Nodes)) yield return child;
		}
	}

	// Puts the caret at a position and scrolls so that its line is at the top of the editor.
	[DllImport("user32")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	internal void GoToPosition(int position) {
		position = Math.Clamp(position, 0, Editor.TextLength);
		Editor.Select(position, 0);
		const int EM_GETFIRSTVISIBLELINE = 0x00CE, EM_LINESCROLL = 0x00B6;
		int first = (int)SendMessage(Editor.Handle, EM_GETFIRSTVISIBLELINE, IntPtr.Zero, IntPtr.Zero);
		int line = Editor.GetLineFromCharIndex(position);
		SendMessage(Editor.Handle, EM_LINESCROLL, IntPtr.Zero, (IntPtr)(line - first));
		Editor.Focus();
	}

	void ShowOutlineMenu(TreeNode node, Point screenPoint) {
		if (node.Tag is not int start) return;
		var level = outline.FirstOrDefault(h => h.Start == start).Level;
		var menu = new ContextMenuStrip { ShowImageMargin = false };
		menu.Items.Add(new ToolStripMenuItem("Go to heading", null, (s, e) => Run(() => GoToPosition(start))));
		menu.Items.Add(new ToolStripSeparator());
		menu.Items.Add(new ToolStripMenuItem($"Promote (to Heading {Math.Max(1, level - 1)})", null, (s, e) => Run(() => api.SetHeadingLevel(start, level - 1))) { Enabled = level > 1 });
		menu.Items.Add(new ToolStripMenuItem($"Demote (to Heading {Math.Min(6, level + 1)})", null, (s, e) => Run(() => api.SetHeadingLevel(start, level + 1))) { Enabled = level < 6 });
		menu.Items.Add(new ToolStripMenuItem("Make normal text", null, (s, e) => Run(() => api.SetHeadingLevel(start, 0))));
		menu.Closed += (s, e) => BeginInvoke(menu.Dispose);
		menu.Show(screenPoint);
	}

	// ---------------------------------------------------------------- Iteration 8: list levels

	void ListLevelsCommand() {
		using var dlg = new ListLevelsDialog(Settings.ListLevels, api.InList(), LegalLists.Format);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Settings.ListLevels = dlg.Levels;
		Settings.Save();
		if (dlg.LegalFormat != LegalLists.Format) {
			api.LegalFormat = dlg.LegalFormat.ToString();
			Record(M("LegalFormat", dlg.LegalFormat.ToString()));
		}
		if (dlg.ApplyToCurrentList) {
			api.ApplyListLevels();
			Record(M("ApplyListLevels"));
			SetStatus("List levels saved and applied to the current list.");
		} else {
			SetStatus("List levels saved. They are used for new numbered lists and when nesting with Tab.");
		}
	}

	void RestartNumberingCommand() {
		if (!api.InList()) throw new InvalidOperationException("Put the caret in a numbered list item first.");
		if (Prompt("Restart numbering", ("Start at", "1")) is not { } values) return;
		if (!int.TryParse(values[0], out int start) || start < 0) throw new ArgumentException("The start number must be 0 or more.");
		Run(() => api.RestartNumbering(start), M("RestartNumbering", start));
	}

	void OnFormClosing(object? sender, FormClosingEventArgs e) {
		if (!discardOnClose && !ConfirmDiscard()) e.Cancel = true;
		discardOnClose = false;
		if (!e.Cancel) SaveSettings();    // here, while the editor still reports its zoom (not after the window has closed)
	}

	// Called by automation: closes after the reply has been sent, so the calling script does not wait for a dialog.
	internal void CloseLater(bool discardChanges) {
		discardOnClose = discardChanges;
		BeginInvoke(Close);
	}

	void FontDialogCommand() {
		var font = api.SelectionRange().Font;
		using var dlg = new FontDialog { ShowEffects = false, ShowColor = false };
		try { dlg.Font = new Font(string.IsNullOrEmpty(font.Name) ? Editor.Font.Name : font.Name, font.Size > 0 ? font.Size : Editor.Font.Size); } catch (ArgumentException) { }
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Editor.SetFontName(dlg.Font.Name);
		Record(M("SetFontName", dlg.Font.Name));
		if (Editor.EnableFontSize) {
			api.SetFontSize(dlg.Font.SizeInPoints);
			Record(M("SetFontSize", dlg.Font.SizeInPoints));
		}
		else SetStatus("Font name applied. Font sizes need free-form mode (View > Free-form fonts).");
	}

	Color? ChooseColor(Color initial) {
		using var dlg = new ColorDialog { FullOpen = true, Color = initial };
		return dlg.ShowDialog(this) == DialogResult.OK ? dlg.Color : null;
	}

	void InsertHyperlinkDialog() {
		string selected = Editor.SelectedText.Trim();
		if (Prompt("Insert hyperlink", ("Address", "https://"), ("Text to display", selected != "" ? selected : "")) is not { } values) return;
		string url = values[0], display = values[1] != "" ? values[1] : values[0];
		Run(() => Editor.InsertHyperlink(url, display), M("InsertHyperlink", url, display));
	}

	void InsertImageDialog() {
		if (ChooseImage() is not { } path) return;
		if (Prompt("Insert image", ("Alt text: what the picture shows, for screen readers (empty for a decorative picture)", Path.GetFileNameWithoutExtension(path))) is not { } values) return;
		Run(() => api.InsertImageFile(path, values[0]), M("InsertImageFile", path, values[0]));
	}

	// Right-click a picture > Alt text...
	void AltTextCommand(int position) {
		string current = api.PictureAltText(position);
		if (Prompt("Alt text", ("Alt text: what the picture shows, for screen readers (empty for a decorative picture)", current)) is not { } values) return;
		if (values[0] == current) return;
		Run(() => api.SetPictureAltText(position, values[0]), M("SetPictureAltText", position, values[0]));
	}

	void InsertThumbnailDialog() {
		if (ChooseImage() is not { } path) return;
		if (Prompt("Insert linked thumbnail", ("Link to the full-size image", new Uri(path).AbsoluteUri), ("Alt text", Path.GetFileNameWithoutExtension(path))) is not { } values) return;
		Run(() => api.InsertLinkedThumbnailFile(values[0], path, values[1]), M("InsertLinkedThumbnailFile", values[0], path, values[1]));
	}

	string? ChooseImage() {
		using var dlg = new OpenFileDialog { Title = "Insert image", Filter = "Images (*.png;*.jpg;*.jpeg;*.gif)|*.png;*.jpg;*.jpeg;*.gif" };
		return dlg.ShowDialog(this) == DialogResult.OK ? dlg.FileName : null;
	}

	// ---------------------------------------------------------------- Find and replace

	void ShowFind(bool replace) {
		findForm ??= new FindReplaceForm(this);
		findForm.ShowFor(replace, Editor.SelectionLength is > 0 and < 100 && !Editor.SelectedText.Contains('\n') ? Editor.SelectedText : null);
	}

	internal bool FindNext() {
		if (findForm == null || findForm.FindText == "") {
			ShowFind(false);
			return false;
		}
		int index = api.Find(findForm.FindText, findForm.MatchCase, findForm.WholeWord);
		Record(M("Find", findForm.FindText, findForm.MatchCase, findForm.WholeWord));
		SetStatus(index < 0 ? $"\"{findForm.FindText}\" was not found." : "");
		return index >= 0;
	}

	internal void ReplaceNext() {
		if (findForm == null) return;
		var comparison = findForm.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
		if (Editor.SelectionLength > 0 && string.Equals(Editor.SelectedText, findForm.FindText, comparison))
			Run(() => Editor.SelectedText = findForm.ReplaceText, M("InsertText", findForm.ReplaceText));
		FindNext();
	}

	internal void ReplaceAll() {
		if (findForm == null || findForm.FindText == "") return;
		int count = 0;
		Run(() => count = api.Replace(findForm.FindText, findForm.ReplaceText, findForm.MatchCase, findForm.WholeWord),
			M("Replace", findForm.FindText, findForm.ReplaceText, findForm.MatchCase, findForm.WholeWord));
		SetStatus($"{count} replacement{(count == 1 ? "" : "s")} made.");
	}

	static ToolStripMenuItem PageBreakItem(ToolStripMenuItem item) {
		item.ShortcutKeyDisplayString = "Ctrl+Enter";    // handled in ProcessCmdKey (a menu shortcut can't be Ctrl+Enter)
		return item;
	}

	// ---------------------------------------------------------------- Keyboard: Tab / Shift+Tab in lists

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
		// Ctrl+Enter: a page break
		if (keyData == (Keys.Control | Keys.Enter) && Editor.IsHandleCreated && msg.HWnd == Editor.Handle) {
			Run(() => api.InsertPageBreak(), M("InsertPageBreak"));
			return true;
		}
		// Tables: row shortcuts, and Delete / Backspace on whole rows (RichEdit would only empty their cells)
		var rowKeys = Keys.Control | Keys.Shift;
		if (Editor.IsHandleCreated && msg.HWnd == Editor.Handle && (keyData == (rowKeys | Keys.Insert) || keyData == (rowKeys | Keys.Delete) || keyData is Keys.Delete or Keys.Back)) {
			if (keyData == (rowKeys | Keys.Insert) && api.InTable()) { Run(() => api.InsertRow(true), M("InsertRow", true)); return true; }
			if (keyData == (rowKeys | Keys.Delete) && api.InTable()) { Run(() => api.DeleteRow(), M("DeleteRow")); return true; }
			if (keyData is Keys.Delete or Keys.Back && api.SelectedRows() > 0) { Run(() => api.DeleteSelectedRows(), M("DeleteSelectedRows")); return true; }
		}
		// Legal lists: Enter, Backspace and Delete renumber in the same undo step (RichEdit's own would leave the numbers,
		// which are text, behind)
		if (Editor.IsHandleCreated && msg.HWnd == Editor.Handle && keyData is Keys.Enter or Keys.Back or Keys.Delete && api.TouchesLegalList()) {
			bool handled = false;
			Run(() => handled = keyData switch { Keys.Enter => api.LegalEnter(), Keys.Back => api.LegalBackspace(), _ => api.LegalDelete() },
				M("Key", keyData switch { Keys.Enter => "Enter", Keys.Back => "Backspace", _ => "Delete" }));
			if (handled) return true;
		}
		if ((keyData == Keys.Tab || keyData == (Keys.Shift | Keys.Tab)) && Editor.IsHandleCreated && msg.HWnd == Editor.Handle && api.InList() && !api.InTable()) {
			int delta = keyData == Keys.Tab ? 1 : -1;
			Run(() => api.Nest(delta), M("Nest", delta));
			return true;
		}
		// Macro hotkeys (Tools > Macro hotkeys) come before the menu shortcuts.
		if (MacroForKeys(keyData) is { } macro) {
			Run(() => RunMacroFromUi(macro));
			return true;
		}
		return base.ProcessCmdKey(ref msg, keyData);
	}

	// ---------------------------------------------------------------- Macros (AHK scripts run as separate processes)

	string MacrosDir => Path.Combine(appDir, "rte_macros");

	void FillMacrosMenu() {
		macrosMenu.DropDownItems.Clear();
		var files = Directory.Exists(MacrosDir) ? Directory.GetFiles(MacrosDir, "*.ahk").OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray() : Array.Empty<string>();
		foreach (var file in files)
			macrosMenu.DropDownItems.Add(new ToolStripMenuItem(Path.GetFileNameWithoutExtension(file), null, (s, e) => Run(() => RunMacroFromUi(file))) {
				ToolTipText = file, ShortcutKeyDisplayString = HotkeyOf(Path.GetFileName(file)) is Keys keys ? HotkeyBox.Describe(keys) : null
			});
		if (files.Length == 0)
			macrosMenu.DropDownItems.Add(new ToolStripMenuItem("(no .ahk files in rte_macros)") { Enabled = false });
		macrosMenu.DropDownItems.Add(new ToolStripSeparator());
		macrosMenu.DropDownItems.Add(new ToolStripMenuItem(recorder.Recording ? "Stop recording" : "Record macro...", null, (s, e) => Run(ToggleRecording)) { ShortcutKeyDisplayString = "Ctrl+Shift+R" });
		macrosMenu.DropDownItems.Add(new ToolStripMenuItem("Macro hotkeys...", null, (s, e) => Run(MacroHotkeysCommand)));
	}

	void RunScriptDialog() {
		using var dlg = new OpenFileDialog { Title = "Run AHK script", Filter = "AutoHotkey scripts (*.ahk)|*.ahk", InitialDirectory = Directory.Exists(MacrosDir) ? MacrosDir : appDir };
		if (dlg.ShowDialog(this) == DialogResult.OK) Run(() => RunMacroFromUi(dlg.FileName));
	}

	internal void RunScript(string path) {
		if (ahkPath == "" || !File.Exists(ahkPath)) throw new InvalidOperationException("AutoHotkey was not found; start RTE Writer with rte_writer.ahk.");
		var start = new System.Diagnostics.ProcessStartInfo(ahkPath) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(path)! };
		start.ArgumentList.Add(path);
		start.Environment["RTE_WRITER_HWND"] = server.Handle.ToInt64().ToString(CultureInfo.InvariantCulture);
		System.Diagnostics.Process.Start(start);
		SetStatus("Running " + Path.GetFileName(path));
	}

	// ---------------------------------------------------------------- Iteration 6: recording macros

	internal bool IsRecording => recorder.Recording;

	// Records a call while a macro is being recorded (from menus, toolbar and dialogs; automation calls are not recorded).
	void Record(string call) {
		if (!recorder.Recording) return;
		recorder.Add(call);
		UpdateRecordingLabel();
	}

	// Runs a macro chosen in the UI (Macros menu, hotkey, Run AHK script). While recording, this is recorded as a
	// RunWait of that script: a replay waits for it to finish before going on, so its output lands in the right place
	// (macros are separate processes and would otherwise run alongside each other).
	void RunMacroFromUi(string path) {
		RunScript(path);
		if (!recorder.Recording) return;
		static string Quoted(string s) => s.Replace("`", "``").Replace("'", "`'");
		string name = Path.GetFileNameWithoutExtension(path);
		string script = string.Equals(Path.GetDirectoryName(Path.GetFullPath(path)), Path.GetFullPath(MacrosDir), StringComparison.OrdinalIgnoreCase)
			? $"' A_ScriptDir '\\{Quoted(Path.GetFileName(path))}"    // recorded macros are saved in rte_macros too
			: Quoted(Path.GetFullPath(path));
		recorder.AddLine($"RunWait('\"' A_AhkPath '\" \"{script}\"')    ; runs the \"{name}\" macro and waits for it to finish");
		UpdateRecordingLabel();
		// Macros that react to events (Rte.On) keep running until the editor closes, so a replay would wait for ever.
		bool keepsRunning = File.ReadAllText(path).Contains("Rte.On(", StringComparison.OrdinalIgnoreCase);
		SetStatus(keepsRunning
			? $"⚠ Recorded: run the \"{name}\" macro - but it keeps running, so a replay would wait for it forever. Remove its RunWait line when you save."
			: $"Recorded: run the \"{name}\" macro. A replay waits for it to finish before going on.");
	}

	static readonly Dictionary<Keys, string> RecordedKeys = new() {
		[Keys.Left] = "Left", [Keys.Right] = "Right", [Keys.Up] = "Up", [Keys.Down] = "Down",
		[Keys.Home] = "Home", [Keys.End] = "End", [Keys.Back] = "Backspace", [Keys.Delete] = "Delete"
	};

	// Keys that reach the editor (menu shortcuts and Tab in lists are recorded as commands, before this).
	void RecordKeyDown(KeyEventArgs e) {
		if (e.KeyCode == Keys.Enter && !e.Control && !e.Alt) {
			recorder.Type('\n');
		} else if (RecordedKeys.TryGetValue(e.KeyCode, out var name) && !e.Alt) {
			// Backspace right after typing just takes back the last typed character.
			if (e.KeyCode == Keys.Back && !e.Control && Editor.SelectionLength == 0 && recorder.Backspace()) { UpdateRecordingLabel(); return; }
			recorder.AddKey((e.Control ? "Ctrl+" : "") + (e.Shift ? "Shift+" : "") + name);
		} else {
			return;
		}
		UpdateRecordingLabel();
	}

	void UpdateRecordingLabel() {
		recordingLabel.Visible = recorder.Recording;
		recordingLabel.Text = $"● Recording ({recorder.Steps} step{(recorder.Steps == 1 ? "" : "s")})";
		recordItem.Text = recorder.Recording ? "Stop &recording" : "&Record macro";
	}

	internal void StartRecording() {
		recorder.Start();
		UpdateRecordingLabel();
		SetStatus("Recording: type and use the menus and toolbar. Mouse clicks are not recorded. Ctrl+Shift+R stops.");
	}

	internal string StopRecording() {
		var lines = recorder.Stop();
		UpdateRecordingLabel();
		return lines.Count == 0 ? "" : MacroScript(lines);
	}

	void ToggleRecording() {
		if (!recorder.Recording) {
			StartRecording();
			return;
		}
		string script = StopRecording();
		if (script == "") {
			SetStatus("Nothing was recorded.");
			return;
		}
		using var dlg = new SaveMacroDialog(UniqueMacroName(), script, MenuShortcuts(), this);
		if (dlg.ShowDialog(this) != DialogResult.OK) {
			SetStatus("The recording was discarded.");
			return;
		}
		Directory.CreateDirectory(MacrosDir);
		string path = Path.Combine(MacrosDir, dlg.MacroName + ".ahk");
		if (File.Exists(path) && MessageBox.Show(this, $"Replace the existing macro \"{dlg.MacroName}\"?", AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) {
			SetStatus("The recording was not saved.");
			return;
		}
		File.WriteAllText(path, dlg.Script.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(false));
		SetMacroHotkey(Path.GetFileName(path), dlg.Hotkey);
		SetStatus($"Saved macro \"{dlg.MacroName}\"" + (dlg.Hotkey != Keys.None ? $" ({HotkeyBox.Describe(dlg.Hotkey)})" : "") + ". It is in Tools > Macros.");
	}

	static string MacroScript(List<string> lines) =>
		"#Requires AutoHotkey v2.0\r\n#Include \"..\\RteClient2.ahk\"\r\n"
		+ $"; Recorded in RTE Writer on {DateTime.Now:g}. Each line is an RTE Writer call (Help > Automation API lists them).\r\n"
		+ string.Join("\r\n", lines) + "\r\n";

	string UniqueMacroName() {
		for (int i = 1; ; i++) {
			string name = $"Recorded macro {i}";
			if (!File.Exists(Path.Combine(MacrosDir, name + ".ahk"))) return name;
		}
	}

	// ---------------------------------------------------------------- Iteration 6: macro hotkeys

	internal Keys? HotkeyOf(string macroFile) =>
		Settings.MacroHotkeys.TryGetValue(macroFile, out var text) && HotkeyBox.Parse(text) is Keys keys && keys != Keys.None ? keys : null;

	internal void SetMacroHotkey(string macroFile, Keys keys) {
		// one macro per hotkey
		foreach (var other in Settings.MacroHotkeys.Where(p => HotkeyBox.Parse(p.Value) == keys).Select(p => p.Key).ToList()) Settings.MacroHotkeys.Remove(other);
		if (keys == Keys.None) Settings.MacroHotkeys.Remove(macroFile);
		else Settings.MacroHotkeys[macroFile] = HotkeyBox.Describe(keys);
		Settings.Save();
	}

	string? MacroForKeys(Keys keys) {
		if (keys == Keys.None || Settings.MacroHotkeys.Count == 0) return null;
		foreach (var (file, text) in Settings.MacroHotkeys)
			if (HotkeyBox.Parse(text) == keys && File.Exists(Path.Combine(MacrosDir, file))) return Path.Combine(MacrosDir, file);
		return null;
	}

	internal string MacroPath(string name) {
		string file = name.EndsWith(".ahk", StringComparison.OrdinalIgnoreCase) ? name : name + ".ahk";
		string path = Path.Combine(MacrosDir, file);
		return File.Exists(path) ? path : throw new FileNotFoundException($"There is no macro \"{name}\" in {MacrosDir}.");
	}

	internal List<string> MacroFiles() =>
		Directory.Exists(MacrosDir) ? Directory.GetFiles(MacrosDir, "*.ahk").Select(Path.GetFileName).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList()! : new List<string>();

	// The shortcut keys of all menu items, to warn about hotkeys that would hide them.
	internal Dictionary<Keys, string> MenuShortcuts() {
		var shortcuts = new Dictionary<Keys, string>();
		void Collect(ToolStripItemCollection items, string path) {
			foreach (var item in items.OfType<ToolStripMenuItem>()) {
				string name = path + item.Text.Replace("&", "");
				if (item.ShortcutKeys != Keys.None) shortcuts.TryAdd(item.ShortcutKeys, name);
				Collect(item.DropDownItems, name + " > ");
			}
		}
		if (MainMenuStrip != null) Collect(MainMenuStrip.Items, "");
		shortcuts.TryAdd(Keys.Tab, "Tab");
		shortcuts.TryAdd(Keys.Shift | Keys.Tab, "Shift+Tab");
		return shortcuts;
	}

	void MacroHotkeysCommand() {
		var files = MacroFiles();
		if (files.Count == 0) throw new InvalidOperationException("There are no macros in " + MacrosDir + " yet.");
		using var dlg = new MacroHotkeysDialog(files, f => HotkeyOf(f) ?? Keys.None, MenuShortcuts());
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Settings.MacroHotkeys.Clear();
		foreach (var (file, keys) in dlg.Hotkeys) if (keys != Keys.None) Settings.MacroHotkeys[file] = HotkeyBox.Describe(keys);
		Settings.Save();
		SetStatus($"{Settings.MacroHotkeys.Count} macro hotkey{(Settings.MacroHotkeys.Count == 1 ? "" : "s")} set. They work while the editor is active.");
	}

	void OpenMacrosFolder() {
		Directory.CreateDirectory(MacrosDir);
		System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", "\"" + MacrosDir + "\""));
	}

	// ---------------------------------------------------------------- Automation entry point

	string DispatchFromAutomation(string[] request) {
		if (request.Length > 1 && request[0] == "post") {
			postedCalls.Enqueue(request[1..]);
			if (postedCalls.Count == 1) BeginInvoke(RunPostedCalls);
			return "0";
		}
		RunPostedCalls();    // calls run in the order they were made, posted or not
		Restart(outlineTimer);
		string reply = Automation.Dispatch(api, Editor, request);
		if (request.Length > 1) automationLabel.Text = $"Automation: {request[1]}";
		UpdateUi();
		UpdateTitle();
		return reply;
	}

	void RunPostedCalls() {
		if (postedCalls.Count == 0) return;
		while (postedCalls.TryDequeue(out var call)) {
			string reply = Automation.Dispatch(api, Editor, new[] { "call" }.Concat(call).ToArray());
			if (reply.StartsWith('1')) {
				api.LastAsyncError = $"{call[0]}: {reply[1..]}";
				SetStatus("⚠ " + api.LastAsyncError);
				Events.Raise(EditorEvent.AsyncError);
			}
			automationLabel.Text = $"Automation: {call[0]} (posted)";
		}
		UpdateUi();
		UpdateTitle();
	}

	// ---------------------------------------------------------------- UI state

	void UpdateTitle() {
		string title = $"{DocumentName}{(Editor.Modified ? " *" : "")} - {AppName}";
		if (Text != title) Text = title;
		if (Editor.Modified != lastModified) {
			lastModified = Editor.Modified;
			Events.Raise(EditorEvent.ModifiedChanged, lastModified ? 1 : 0);
		}
	}

	void UpdateWordCount() => wordsLabel.Text = $"{api.WordCount()} words";

	void UpdateUi() {
		if (!Editor.IsHandleCreated || updatingUi) return;
		updatingUi = true;
		try {
			var range = api.SelectionRange();
			var font = range.Font;
			toggles["bold"].Checked = font.Bold == tomConstants.tomTrue;
			toggles["italic"].Checked = font.Italic == tomConstants.tomTrue;
			toggles["underline"].Checked = font.Underline > 0;
			toggles["strike"].Checked = font.StrikeThrough == tomConstants.tomTrue;
			toggles["super"].Checked = font.Superscript == tomConstants.tomTrue;
			toggles["sub"].Checked = font.Subscript == tomConstants.tomTrue;
			string name = font.Name ?? "";
			if (!fontCombo.Focused) fontCombo.Text = name;
			if (!sizeCombo.Focused) {
				// the entry shown is the one chosen (none for a mixed selection): otherwise picking the entry chosen
				// before would be no change, and do nothing
				string shown = font.Size > 0 ? font.Size.ToString("0.#", CultureInfo.CurrentCulture) : "";
				sizeCombo.SelectedIndex = sizeCombo.Items.IndexOf(shown);
				sizeCombo.Text = shown;
			}
			var para = range.Para;
			foreach (var (alignment, button) in alignButtons) button.Checked = para.Alignment == (int)alignment;
			int listType = para.ListType & 0xFFFF;
			bulletButton.Checked = listType == (int)ListStyle.Bullet;
			numberButton.Checked = listType != tomConstants.tomListNone && listType != (int)ListStyle.Bullet;
			blockCombo.Enabled = Editor.EnableHtmlFontSizing;
			sizeCombo.ToolTipText = Editor.EnableFontSize ? "Font size" : "Font size (needs free-form mode: View > Free-form fonts)";
			blockCombo.SelectedIndex = Editor.EnableHtmlFontSizing && api.GetBlockStyle() is BlockStyle style ? (int)style : -1;

			int line = Editor.GetLineFromCharIndex(Editor.SelectionStart);
			int column = Editor.SelectionStart - Editor.GetFirstCharIndexFromLine(line);
			positionLabel.Text = $"Ln {line + 1}, Col {column + 1}";
			modeLabel.Text = Editor.EnableHtmlFontSizing ? "Structured" : "Free-form";
		} catch (COMException) {
			// The document can be briefly unavailable while it is being replaced.
		} finally {
			updatingUi = false;
		}
	}

	// ---------------------------------------------------------------- Small dialogs

	void ShowText(string title, string text) {
		using var dlg = new Form { Text = title, StartPosition = FormStartPosition.CenterParent, ShowInTaskbar = false, MinimizeBox = false,
			Size = new Size(LogicalToDeviceUnits(760), LogicalToDeviceUnits(560)) };
		var box = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false,
			Font = new Font("Consolas", 10f), Text = text.Replace("\r\n", "\n").Replace("\n", "\r\n"), BackColor = SystemColors.Window };
		var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(4) };
		var close = new Button { Text = "Close", DialogResult = DialogResult.Cancel, AutoSize = true };
		var copy = new Button { Text = "Copy", AutoSize = true };
		copy.Click += (s, e) => { if (text != "") Clipboard.SetText(text); };
		buttons.Controls.AddRange(new Control[] { close, copy });
		dlg.Controls.Add(box);
		dlg.Controls.Add(buttons);
		dlg.CancelButton = close;
		box.Select(0, 0);
		dlg.ShowDialog(this);
	}

	string[]? Prompt(string title, params (string Label, string Value)[] fields) {
		using var dlg = new Form {
			Text = title, FormBorderStyle = FormBorderStyle.FixedDialog, MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false,
			StartPosition = FormStartPosition.CenterParent, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8)
		};
		var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
		var boxes = new List<TextBox>();
		foreach (var (label, value) in fields) {
			table.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left });
			var box = new TextBox { Text = value, Width = LogicalToDeviceUnits(380) };
			table.Controls.Add(box);
			boxes.Add(box);
		}
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
		buttons.Controls.AddRange(new Control[] { cancel, ok });
		table.Controls.Add(buttons);
		table.SetColumnSpan(buttons, 2);
		dlg.Controls.Add(table);
		dlg.AcceptButton = ok;
		dlg.CancelButton = cancel;
		return dlg.ShowDialog(this) == DialogResult.OK ? boxes.Select(b => b.Text).ToArray() : null;
	}
}

// The automation surface. Every public member can be called from AHK by name (Rte.<Name> in RteClient.ahk), along with
// the RichEditBox members listed in Automation.EditorMembers. A member here wins over a RichEditBox member of the same name.
// These methods never show dialogs, so a calling script is never left waiting for the user.
public class EditorApi
{
	readonly WriterForm form;
	RichEditBox ed => form.Editor;

	internal EditorApi(WriterForm form) {
		this.form = form;
		form.Editor.TextChanged += (s, e) => altCache = null;
	}

	public string Version => $"{WriterForm.AppName} 0.11, RichEditorNET " +
		(typeof(RichEditBox).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0]
			?? typeof(RichEditBox).Assembly.GetName().Version?.ToString());

	// ---- Documents

	/// Starts a new, empty document (unsaved changes are discarded).
	public void New() {
		ed.Html = "";
		form.Bands = form.DefaultBands();
		ed.Modified = false;
		FileConverter = "";
		form.SetFileName(null);
		form.Events.Raise(EditorEvent.Opened);
	}

	/// Opens an .html/.htm, .md/.markdown, .rtf, .docx/.doc, .odt or .txt file (unsaved changes are discarded).
	/// Word and OpenDocument files are converted by Word, openxml, LibreOffice or pandoc (see Converters and
	/// LastConverter); via names the converter(s) to use, e.g. "openxml" (Save then uses them too).
	public void Open(string path, string via = "") {
		string full = Path.GetFullPath(path);
		string extension = Path.GetExtension(full).ToLowerInvariant();
		LastConverter = "";
		// the header and footer: the document's own; documents that can have them but don't have none, others get the default
		PageBands? bands = null;
		bool canHaveBands = extension is ".rtf" or ".docx" or ".doc" or ".odt";
		switch (extension) {
			case ".md": case ".markdown": {
				string markdown = File.ReadAllText(full);
				bands = HeaderFooter.FromMarkdown(markdown);
				MarkdownConverter.Load(ed, markdown, Path.GetDirectoryName(full));
				break;
			}
			case ".txt": ed.Text = File.ReadAllText(full).Replace("\r\n", "\n"); break;
			case ".rtf": {
				string rtf = File.ReadAllText(full, Encoding.Latin1);
				bands = HeaderFooter.FromRtf(rtf);
				Rtf = rtf;
				LegalLists.AdoptTypedNumbers(ed);
				break;
			}
			case ".docx": case ".doc": case ".odt": {
				var (kind, content, converter) = OfficeConverters.Import(full, via);
				bands = extension == ".docx" ? HeaderFooter.FromDocx(full) : extension == ".odt" ? HeaderFooter.FromOdt(full) : kind == "rtf" ? HeaderFooter.FromRtf(content) : null;
				if (kind == "rtf") Rtf = content; else SetHtml(content);
				LegalLists.ConvertTypedLists(ed, 0, ed.TextDocument.Range2(0, 0).StoryLength, requireTab: true,
					converter == OfficeConverters.OpenXml || kind == "html" ? OfficeConverters.TypedIndents : null);    // numbers a converter left as text
				LastConverter = converter;
				break;
			}
			default: {
				string html = File.ReadAllText(full);
				bands = HeaderFooter.FromHtml(html);
				SetHtml(html);
				break;
			}
		}
		form.Bands = bands ?? (canHaveBands ? new PageBands() : form.DefaultBands());
		PageBreaks.Normalize(ed);    // (form feeds in plain text, for one)
		MarksFollowText();
		if (LegalLists.AnyLegal(ed.TextDocument)) LegalLists.Renumber(ed);    // (legal items' indents fitted to their numbers' size)
		ed.ClearUndo();
		ed.Select(0, 0);
		ed.Modified = false;
		FileConverter = extension is ".docx" or ".doc" or ".odt" ? via : "";
		form.SetFileName(full);
		form.AddRecent(full);
		form.Events.Raise(EditorEvent.Opened);
	}

	/// Saves to the current file and returns its path.
	public string Save() =>
		SaveAs(form.FileName ?? throw new InvalidOperationException("The document has not been saved yet; use SaveAs(path)."), FileConverter);

	/// The converter(s) chosen for the current Word or OpenDocument file when it was opened or saved ("" = automatic);
	/// Save uses them.
	public string FileConverter { get; private set; } = "";

	/// Saves as HTML, Markdown, RTF, Word, OpenDocument or text, chosen by the extension (.html, .md, .rtf, .docx, .odt,
	/// .txt), and returns the full path. .docx and .odt are converted (see Converters and LastConverter); via names the
	/// converter(s) to use, e.g. "openxml".
	public string SaveAs(string path, string via = "") {
		string full = Path.GetFullPath(path);
		string extension = Path.GetExtension(full).ToLowerInvariant();
		LastConverter = "";
		if (extension is ".docx" or ".odt") Export(full, via);
		// (ordinary lists with hanging indents that fit their numbers: RichEdit's own RTF indents them so that Word and
		// LibreOffice put the numbers too far in)
		else if (extension == ".rtf")    // (and pictures as PNG, and the header and footer)
			File.WriteAllText(full, WithBands(OfficeConverters.PngPictures(OfficeConverters.WithAltTexts(OfficeConverters.ListIndentsRtf(RawRtf), AltTexts())), full), Encoding.Latin1);
		else {
			string content = extension switch {
				".md" or ".markdown" => MarkdownConverter.ToMarkdown(ed, ImageFiles(full)) + (OwnBands ? HeaderFooter.MarkdownLine(form.Bands) + "\n" : ""),
				".txt" => ed.Text.Replace("\n", "\r\n"),
				_ => HtmlFile(Path.GetFileNameWithoutExtension(full))
			};
			File.WriteAllText(full, content, new UTF8Encoding(false));
		}
		ed.Modified = false;
		FileConverter = extension is ".docx" or ".odt" ? via : "";
		form.SetFileName(full);
		form.AddRecent(full);
		form.Events.Raise(EditorEvent.Saved);
		return full;
	}

	/// The full path of the current file ("" for a new document).
	public string FileName => form.FileName ?? "";

	/// The document as HTML, with list numbering that RichEditorNET's HTML leaves out (start numbers, restarts,
	/// punctuation such as "a)"; the punctuation needs the <style> element at the start).
	public string GetHtml() {
		string body = HtmlBody(out string css);
		return HtmlLists.StyleElement(css) + body;
	}

	/// Replaces the document with HTML (a whole page or a part of one), including the list numbering and tables GetHtml writes.
	public void SetHtml(string html) {
		RtfTables.LoadHtml(ed, BodyOf(PageBreaks.MarkHtml(html)), HtmlLists.Load, HtmlLists.Load, TableWidth());
		if (PageBreaks.FromMarkers(ed) + PageBreaks.Normalize(ed) > 0) ed.ClearUndo();
	}

	// The library's HTML, with page breaks, real tables (it writes their cells as paragraphs) and list numbering.
	string HtmlBody(out string css) {
		string body = HtmlLists.Export(ed, RtfTables.ExportHtml(ed, PageBreaks.ExportHtml(ed, ed.Html)), out css);
		css = ParagraphCss + css;
		return body;
	}

	// Paragraphs as the editor shows them: browsers put a gap between paragraphs (and headings), and give an empty
	// paragraph no height, where the editor has neither gaps nor vanishing blank lines.
	const string ParagraphCss = "p, h1, h2, h3, h4, h5, h6 { margin: 0; }\r\np:empty::before { content: \"\\00a0\"; }\r\n";

	/// The document as HTML (get), or replaced by HTML (set); see GetHtml and SetHtml.
	public string Html {
		get => GetHtml();
		set => SetHtml(value);
	}

	string HtmlFile(string title) {
		string body = HtmlBody(out string css);
		return HtmlDocument(body, title, css, OwnBands ? HeaderFooter.HtmlMeta(form.Bands) : "");
	}

	// Whether the document's header and footer differ from the default (HTML and Markdown keep them only then: a
	// document that has the default gets whatever the default is when it is opened).
	bool OwnBands => !form.Bands.SameAs(form.DefaultBands());

	// RTF with the document's header, footer and page layout, for path.
	string WithBands(string rtf, string path) =>
		HeaderFooter.AddToRtf(rtf, form.Bands, form.PageLayout(), Path.GetFileName(path), path, ed.Font.SizeInPoints);

	// ---- Rich Text Format, Word and OpenDocument

	// RichEdit's own RTF (without the pictures' alt texts, for changes that go through a RichEdit).
	string RawRtf => ((RichTextBox)ed).Rtf ?? "";

	/// The document as RTF (get), or replaced by RTF (set). RTF keeps everything the editor has.
	public string Rtf {
		get => OfficeConverters.WithAltTexts(RawRtf, AltTexts());
		set {
			value = OfficeConverters.SimplePictures(value);    // (Word's picture wrappers, which RichEdit can drop)
			value = PageBreaks.FromPageBreakBefore(value);       // (LibreOffice's and Word's "page break before" paragraphs)
			value = HeaderFooter.StripFromRtf(value);             // (the header and footer are the document's Bands)
			((RichTextBox)ed).Rtf = OfficeConverters.WithoutAltTexts(value);    // (RichEdit would put their accents into the text)
			PageBreaks.Normalize(ed);    // (Word's \page can be in the middle of a paragraph)
			RestorePictures(value);       // (with the alt texts, from the RTF as it was)
			LegalLists.LevelsFromNumbers(ed.TextDocument);    // (RichEdit's RTF loses the levels of legal items)
			ed.ClearUndo();
			ed.Select(0, 0);
		}
	}

	// Each embedded object's alt text (an editor picture's; "" for others), in document order. RichEditorNET gives the
	// selected picture's, so each is selected in turn (and the selection put back).
	List<string> AltTexts() {
		var doc = ed.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var alts = new List<string>();
		if (story.IndexOf('\uFFFC') < 0) return alts;
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		bool startActive = (selection.Flags & 1) != 0;    // tomSelStartActive: the caret at the selection's start
		doc.Freeze();
		try {
			for (int i = story.IndexOf('\uFFFC'); i >= 0; i = story.IndexOf('\uFFFC', i + 1)) {
				string alt = "";
				try { selection.SetRange(i, i + 1); alt = ed.GetImageAltText() ?? ""; }
				catch (Exception ex) when (ex is COMException or InvalidOperationException or NullReferenceException) { }
				alts.Add(alt.Trim());
			}
		} finally {
			selection.SetRange(start, end);
			if (startActive && end > start) selection.Flags |= 1;
			doc.Unfreeze();
		}
		return alts;
	}

	// Pictures that come from RTF are plain RichEdit pictures, which RichEditorNET's HTML (and so Markdown and the
	// built-in .docx converter) leaves out: each becomes an editor picture again, made from the image in the RTF, at
	// its size and with its link.
	void RestorePictures(string rtf) {
		var pictures = OfficeConverters.Pictures(rtf);
		if (pictures.Count == 0) return;
		var doc = ed.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var positions = story.Select((c, i) => (c, i)).Where(x => x.c == '\uFFFC').Select(x => x.i).ToList();
		if (positions.Count != pictures.Count) return;    // other embedded objects as well: they can't be told apart
		int oldWidth = ed.DefaultImageWidth, oldHeight = ed.DefaultImageHeight;
		try {
			for (int k = positions.Count - 1; k >= 0; k--) {
				if (pictures[k] is not { } picture) continue;
				int at = positions[k];
				string url = Regex.Match(doc.Range2(at, at + 1).URL ?? "", @"""?([^""]+)""?").Groups[1].Value;
				int before = doc.Range2(0, 0).StoryLength;
				var selection = doc.Selection2;
				selection.SetRange(at, at + 1);
				selection.Text = "";
				// a picture that was a link's only text takes the link's hidden address with it: back to where the link began
				at -= Math.Max(0, before - doc.Range2(0, 0).StoryLength - 1);
				selection.SetRange(at, at);
				ed.DefaultImageWidth = picture.Width;
				ed.DefaultImageHeight = picture.Height;
				using var stream = new MemoryStream(picture.Image);
				ed.InsertImage(stream, picture.Alt);
				if (url != "") doc.Range2(at, at + 1).URL = "\"" + url.Replace("\"", "%22") + "\"";
			}
		} catch (Exception ex) {
			form.SetStatus("Some pictures could not be made editor pictures again: " + ex.Message);
		} finally {
			ed.DefaultImageWidth = oldWidth;
			ed.DefaultImageHeight = oldHeight;
		}
	}

	/// Writes a copy of the document as Word (.docx); the current file name does not change. via names the converters
	/// to try, e.g. "openxml" or "word,pandoc" (default: Word if it isn't running, openxml, LibreOffice, pandoc).
	/// Returns the full path.
	public string ExportDocx(string path, string via = "") => Export(Path.ChangeExtension(Path.GetFullPath(path), ".docx"), via);

	/// Writes a copy of the document as OpenDocument text (.odt), like ExportDocx (default order: LibreOffice, Word if
	/// it isn't running, pandoc).
	public string ExportOdt(string path, string via = "") => Export(Path.ChangeExtension(Path.GetFullPath(path), ".odt"), via);

	string Export(string full, string? via) {
		OfficeConverters.DefaultFont = (ed.Font.Name, ed.Font.SizeInPoints);
		var alts = AltTexts();
		LastConverter = OfficeConverters.Export(full, RawRtf, () => HtmlFile(Path.GetFileNameWithoutExtension(full)), via,
			rtf => WithBands(OfficeConverters.WithAltTexts(rtf, alts), full));    // (alt texts last: RichEdit, which some changes go through, mangles them)
		// the header and footer: parts of their own if the converter had no RTF to make them from, and the exact copy
		try {
			if (full.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) HeaderFooter.AddToDocx(full, form.Bands, form.PageLayout(), Path.GetFileName(full), full, parts: true);
			else HeaderFooter.AddToOdt(full, form.Bands, Path.GetFileName(full), full);
			OfficeConverters.AltTextsToFile(full, AltTexts());    // (pictures' alt texts that the converter left out)
		} catch (Exception ex) {
			form.SetStatus("⚠ The header and footer could not be added to " + Path.GetFileName(full) + ": " + ex.Message);
		}
		return full;
	}

	/// The converter used by the last Open, Save, SaveAs or Export of a Word or OpenDocument file ("word",
	/// "libreoffice", "openxml" or "pandoc"; "" if none was needed).
	public string LastConverter { get; private set; } = "";

	/// Which converters for Word and OpenDocument files are available.
	public string Converters() => OfficeConverters.Describe();

	// ---- Tables

	/// Inserts a table at the caret (headerRow = 1 makes the first row bold) and puts the caret in its first cell.
	public void InsertTable(int rows = 3, int columns = 3, bool headerRow = true) {
		if (rows < 1 || columns < 1 || rows > 500 || columns > 63) throw new ArgumentException("A table has 1 to 500 rows and 1 to 63 columns.");
		RtfTables.Insert(ed, rows, columns, headerRow, TableWidth());
	}

	/// Whether the caret is in a table.
	public bool InTable() => RtfTables.At(ed, ed.SelectionStart) != null;

	/// Inserts an empty row below (below = 1) or above (0) the caret's row.
	public void InsertRow(bool below = true) { RtfTables.InsertRow(ed, below); Redraw(); }

	/// Inserts an empty column right (right = 1) or left (0) of the caret's column; the table keeps its width.
	public void InsertColumn(bool right = true) { RtfTables.InsertColumn(ed, right); Redraw(); }

	/// Deletes the caret's row (and the table, if it was the only row).
	public void DeleteRow() { RtfTables.DeleteRow(ed); Redraw(); }

	/// Deletes the caret's column (and the table, if it was the only column).
	public void DeleteColumn() { RtfTables.DeleteColumn(ed); Redraw(); }

	/// Deletes the table the caret is in.
	public void DeleteTable() { RtfTables.DeleteTable(ed); Redraw(); }

	/// The number of table rows the selection covers completely (0 if it is not whole rows of one table).
	public int SelectedRows() => RtfTables.SelectedRows(ed).Rows.Count;

	/// Deletes the table rows the selection covers completely (the whole table if that is all of them), as the Delete
	/// key does; returns how many were deleted.
	public int DeleteSelectedRows() { int count = RtfTables.DeleteSelectedRows(ed); Redraw(); return count; }

	// RichEdit doesn't repaint where a removed table row was (its borders stayed on screen): the whole editor is redrawn.
	void Redraw() => ed.Invalidate();

	// ---- Page breaks

	// ---- Pictures' alt text (iteration 15)

	List<string>? altCache;    // the pictures' alt texts in document order, from RichEditorNET's HTML; null: to be read again

	/// The alt text of the picture at a position ("" if it has none).
	public string PictureAltText(int position) {
		var doc = ed.TextDocument;
		if (doc.Range2(position, position + 1).Text != "\uFFFC") throw new ArgumentException("There is no picture at " + position + ".");
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		int index = story[..position].Count(c => c == '\uFFFC');
		altCache ??= Regex.Matches(ed.Html, "<img [^>]*>", RegexOptions.IgnoreCase)
			.Select(m => WebUtility.HtmlDecode(Regex.Match(m.Value, "\\balt=\"([^\"]*)\"").Groups[1].Value)).ToList();
		int pictures = story.Count(c => c == '\uFFFC');
		return altCache.Count == pictures && index < altCache.Count ? altCache[index] : "";
	}

	/// Gives the picture at a position a new alt text ("" for none). RichEditorNET can't change a picture's alt text, so
	/// the picture is replaced by the same picture (image, size and link) with the new one, in one undo step.
	public void SetPictureAltText(int position, string alt) {
		var doc = ed.TextDocument;
		if (doc.Range2(position, position + 1).Text != "\uFFFC") throw new ArgumentException("There is no picture at " + position + ".");
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		selection.SetRange(position, position + 1);
		var pictures = OfficeConverters.Pictures(((RichTextBox)ed).SelectedRtf ?? "");
		if (pictures.Count != 1 || pictures[0] is not { } picture) {
			selection.SetRange(start, end);
			throw new InvalidOperationException("This picture can't be read, so its alt text can't be changed.");
		}
		doc.BeginEditCollection();
		try {
			string url = Regex.Match(doc.Range2(position, position + 1).URL ?? "", @"""?([^""]+)""?").Groups[1].Value;
			int before = doc.Range2(0, 0).StoryLength;
			selection.Text = "";
			// a picture that was a link's only text takes the link's hidden address with it: back to where the link began
			int at = position - Math.Max(0, before - doc.Range2(0, 0).StoryLength - 1);
			selection.SetRange(at, at);
			int oldWidth = ed.DefaultImageWidth, oldHeight = ed.DefaultImageHeight;
			try {
				ed.DefaultImageWidth = picture.Width;
				ed.DefaultImageHeight = picture.Height;
				using var stream = new MemoryStream(picture.Image);
				ed.InsertImage(stream, alt);
			} finally {
				ed.DefaultImageWidth = oldWidth;
				ed.DefaultImageHeight = oldHeight;
			}
			if (url != "") doc.Range2(at, at + 1).URL = "\"" + url.Replace("\"", "%22") + "\"";
		} finally {
			doc.EndEditCollection();
		}
		altCache = null;
		selection.SetRange(start, end);
	}

	/// Inserts a page break at the caret (replacing the selection): the text after it starts a new page when printed,
	/// and in Word and LibreOffice. A numbered list it splits goes on with its numbers after it. Ctrl+Enter.
	public void InsertPageBreak() {
		var doc = ed.TextDocument;
		string hint = "";
		doc.BeginEditCollection();
		try {
			PageBreaks.Insert(ed);
			// the list items after the break go on from those before it (legal numbers always do)
			var para = ParaAt(doc.Selection2.Start);
			int type = para.ListType & 0xFFFF;
			if (type != tomConstants.tomListNone && type != tomConstants.tomListBullet && !LegalLists.IsLegal(para.ListType)) {
				if (para.ListLevelIndex == 0) {
					bool continued = true;
					try { ContinueNumbering(); } catch (InvalidOperationException) { continued = false; }
					if (continued) hint = "Page break inserted. The numbering after it goes on from the list before it, but stays as it is: if you add or remove items "
						+ "before the break, put the caret in the first item after it and use Paragraph > Continue numbering.";
				} else {
					hint = "Page break inserted. The sub-list after it is numbered from the start again (a break before or after the list keeps its numbering).";
				}
			}
			if (LegalLists.AnyLegal(doc)) LegalLists.Renumber(ed);
		} finally {
			doc.EndEditCollection();
		}
		Redraw();
		if (hint != "") form.SetStatus(hint);
	}

	/// The number of page breaks in the document.
	public int PageBreakCount() => PageBreaks.Positions(ed.TextDocument).Count;

	/// The caret's table position as "row,column" (from 1), or "" outside tables.
	public string TableCell() => RtfTables.At(ed, ed.SelectionStart) is { } at ? $"{at.Row + 1},{at.Cell + 1}" : "";

	/// The number of tables in the document.
	public int TableCount => RtfTables.Find(ed.TextDocument).Count;

	// Moves the caret to the start of the next (delta 1) or previous (-1) cell, as Tab and Shift+Tab do.
	internal void NextCell(int delta) {
		if (RtfTables.At(ed, ed.SelectionStart) is not { } at) return;
		var cells = at.Table.Rows.SelectMany(r => r.Cells).ToList();
		int index = at.Table.Rows.Take(at.Row).Sum(r => r.Cells.Count) + at.Cell + delta;
		if (index < 0 || index >= cells.Count) return;
		ed.Select(cells[index].Start, cells[index].End - cells[index].Start);
	}

	// After opening: paragraph and cell marks take the size of the text before them, as in Word. HTML has no paragraph
	// marks (they came back at the default size, and list numbers, drawn in the mark's format, with them), and
	// LibreOffice's RTF gives table marks its own default (12 pt, making rows taller). Table row markers take the size of
	// their row's first and last cell.
	void MarksFollowText() {
		var doc = ed.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		static bool Mark(char c) => c is '\r' or '\u0007' or '\uFFF9' or '\uFFFB' or '\v';
		float SizeAt(int i) => i >= 0 && i < story.Length ? doc.Range2(i, i + 1).Font.Size : -1;
		void SetSize(int i, float size) {
			if (size <= 0) return;
			var r = doc.Range2(i, i + 1);
			if (Math.Abs(r.Font.Size - size) > 0.01) r.Font.Size = size;
		}
		// paragraph and cell marks after text
		for (int i = 1; i < story.Length; i++)
			if (story[i] is '\r' or '\u0007' && !Mark(story[i - 1]) && story[i - 1] != '\uFFFC') SetSize(i, SizeAt(i - 1));
		// table rows: U+FFF9 and its paragraph mark take the first cell's text size, U+FFFB and its mark the last cell's mark
		for (int i = 0; i < story.Length; i++) {
			if (story[i] == '\uFFF9' && i + 2 < story.Length) { float size = SizeAt(i + 2); SetSize(i, size); if (story[i + 1] == '\r') SetSize(i + 1, size); }
			else if (story[i] == '\uFFFB' && i > 0) { float size = SizeAt(i - 1); SetSize(i, size); if (i + 1 < story.Length && story[i + 1] == '\r') SetSize(i + 1, size); }
		}
	}

	// ---- Font sizes (RichEditorNET's set the size of the WinForms selection, which is short in tables, and do sums
	// with the "undefined" size of a selection with several sizes: Decrease made such text 1 point)

	/// Sets the font size of the selection (free-form mode). Where the selection covers a paragraph's whole text, its
	/// paragraph mark gets the size too, so that its list number (and a legal number) follows, as in Word.
	public void SetFontSize(float size) {
		if (size <= 0 || size > 1638) throw new ArgumentOutOfRangeException(nameof(size), "Font sizes are from 1 to 1638 points.");
		ChangeFontSize(_ => size);
	}

	/// Makes each part of the selection points larger (keeping their differences).
	public void IncreaseFontSize(float points = 1) => ChangeFontSize(size => size + points);

	/// Makes each part of the selection points smaller (keeping their differences, down to 1 point).
	public void DecreaseFontSize(float points = 1) => ChangeFontSize(size => size - points);

	void ChangeFontSize(Func<float, float> change) {
		if (!ed.EnableFontSize) throw new InvalidOperationException("Font sizes need free-form mode (View > Free-form fonts). Structured documents use block styles.");
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		float New(float size) => Math.Clamp(MathF.Round(change(size > 0 ? size : ed.Font.SizeInPoints) * 2, MidpointRounding.AwayFromZero) / 2, 1, 1638);
		if (start == end) {    // the caret: the size of what is typed next
			selection.Font.Size = New(selection.Font.Size);
			return;
		}
		var ranges = new List<(int Start, int End)> { (start, end) };
		foreach (var p in LegalLists.Paragraphs(doc)) {
			int textStart = p.Start + p.PrefixLength, textEnd = p.Text.Length > 0 ? p.Start + p.Text.Length : p.Start;
			if (textEnd <= textStart || start > textStart || end < textEnd || p.End <= start || p.Start >= end) continue;
			if (p.PrefixLength > 0) ranges.Add((p.Start, textStart));    // a legal number
			if (p.End > textEnd) ranges.Add((textEnd, p.End));           // the paragraph mark (list number, cell mark)
		}
		int storyEnd = doc.Range2(0, 0).StoryLength;
		if (end >= storyEnd - 1) ranges.Add((storyEnd - 1, storyEnd));    // Select all: the document's last mark too (what is typed at the end)
		doc.BeginEditCollection();
		try {
			foreach (var (from, to) in ranges) {
				// run by run, so that each keeps its own size when sizes are changed by points
				int at = from;
				while (at < to) {
					var run = doc.Range2(at, at);
					run.MoveEnd(tomConstants.tomCharFormat, 1);
					int runEnd = Math.Min(Math.Max(run.End, at + 1), to);
					var part = doc.Range2(at, runEnd);
					part.Font.Size = New(part.Font.Size);
					at = runEnd;
				}
			}
			if (LegalLists.AnyLegal(doc)) LegalLists.Renumber(ed);    // legal numbers' indents fit their new size
		} finally {
			doc.EndEditCollection();
			selection.SetRange(start, end);
		}
	}

	// ---- Iteration 10: legal numbering

	/// Makes the selected paragraphs a legal numbered list (1., 1.1, 1.1.1 by list level), or, if the caret is in one,
	/// ends it for the selected paragraphs. Paragraphs that start with a typed number such as "2.1 " take its level.
	public void LegalNumbering() {
		var range = SelectionRange();
		if (InLegalList()) LegalLists.Remove(ed, range.Start, range.End);
		else LegalLists.Make(ed, range.Start, range.End);
	}

	/// Paragraphs that begin with typed list numbers or bullets ("1. ", "a) ", "(iv) ", "• ", or with a tab after them),
	/// two or more in a row, become real list items: numbering style, punctuation and start from the numbers, levels
	/// from the indents; compound numbers ("1.2") become legal items. Works on the selected paragraphs, or the whole
	/// document when nothing is selected. Returns how many paragraphs changed.
	public int ConvertTypedNumbers() {
		var range = SelectionRange();
		bool all = range.Start == range.End;
		return LegalLists.ConvertTypedLists(ed, all ? 0 : range.Start, all ? ed.TextDocument.Range2(0, 0).StoryLength : range.End, requireTab: false);
	}

	/// True if the caret is in a legal numbered item.
	public bool InLegalList() => LegalLists.IsLegal(ed.TextDocument.Range2(ed.SelectionStart, ed.SelectionStart).Para.ListType);

	/// True if the selection includes a legal numbered item.
	public bool TouchesLegalList() {
		var range = SelectionRange();
		return LegalLists.Touched(ed.TextDocument, range.Start, range.End).Any(p => p.Legal);
	}

	/// Puts every legal number right (the editor does this itself; for scripts that change the text directly).
	/// Returns 1 if anything changed.
	public bool Renumber() {
		bool changed = false;
		LegalLists.Grouped(ed, () => changed = LegalLists.Renumber(ed));
		return changed;
	}

	/// How legal numbers look: "Top" (1. 1.1 1.1.1), "All" (1. 1.1. 1.1.1.) or "None" (1 1.1 1.1.1). Setting it
	/// renumbers the document's legal lists and keeps it in the settings.
	public string LegalFormat {
		get => LegalLists.Format.ToString();
		set {
			if (!Enum.TryParse<LegalLists.Periods>(value, true, out var format)) throw new ArgumentException("LegalFormat is Top, All or None.");
			LegalLists.Format = format;
			form.Settings.LegalFormat = format.ToString();
			form.Settings.Save();
			if (LegalLists.AnyLegal(ed.TextDocument)) LegalLists.Grouped(ed, () => { });
		}
	}

	/// Makes paragraphs that start with typed compound numbers ("1.2<tab>Text") a legal list; returns how many changed.
	/// Opening a document does this by itself.
	public int AdoptTypedNumbers() => LegalLists.AdoptTypedNumbers(ed);

	/// Enter in a legal list: a new item, renumbered; in an empty item, the list ends there. Returns 0 if the caret is
	/// not in a legal item.
	public bool LegalEnter() {
		var selection = ed.TextDocument.Selection2;
		var p = ParagraphAt(selection.Start);
		if (p == null || !p.Legal) return false;
		if (selection.Start == selection.End && p.Text.Length == p.PrefixLength) {
			LegalLists.Remove(ed, p.Start, p.Start);
			return true;
		}
		LegalLists.Grouped(ed, () => selection.TypeText("\r"));
		return true;
	}

	/// Backspace in a legal list: at the start of an item's text, its number goes (it becomes a normal paragraph); with a
	/// selection, the selection goes and the list is renumbered. Returns 0 when RichEdit's own Backspace is right.
	public bool LegalBackspace() {
		var selection = ed.TextDocument.Selection2;
		if (selection.Start != selection.End) return DeleteLegalSelection();
		var p = ParagraphAt(selection.Start);
		if (p == null || !p.Legal || selection.Start > p.Start + p.PrefixLength) return false;
		LegalLists.Remove(ed, p.Start, p.Start);
		return true;
	}

	/// Delete in a legal list: at the end of an item, the next paragraph joins it (without its number); with a selection,
	/// the selection goes and the list is renumbered. Returns 0 when RichEdit's own Delete is right.
	public bool LegalDelete() {
		var selection = ed.TextDocument.Selection2;
		if (selection.Start != selection.End) return DeleteLegalSelection();
		var p = ParagraphAt(selection.Start);
		if (p == null || selection.Start != p.End - 1 || p.End >= ed.TextDocument.Range2(0, 0).StoryLength) return false;
		var next = ParagraphAt(p.End);
		if (next == null || (!p.Legal && !next.Legal)) return false;
		LegalLists.Grouped(ed, () => {
			var doc = ed.TextDocument;
			doc.Range2(p.End - 1, next.Start + next.PrefixLength).Text = "";
			KeepFormat(p);
			doc.Selection2.SetRange(p.End - 1, p.End - 1);
		});
		return true;
	}

	// Deletes a selection that includes legal items; the paragraph it ends in keeps the first paragraph's list format
	// (RichEdit would give it the last one's).
	bool DeleteLegalSelection() {
		var selection = ed.TextDocument.Selection2;
		var first = ParagraphAt(selection.Start);
		if (first == null) return false;
		LegalLists.Grouped(ed, () => {
			selection.Text = "";
			KeepFormat(first);
		});
		return true;
	}

	void KeepFormat(LegalLists.Paragraph first) {
		var para = ed.TextDocument.Range2(first.Start, first.Start).Para;
		if (first.Legal) {
			if (!LegalLists.IsLegal(para.ListType) || para.ListLevelIndex != first.Level) LegalLists.SetLegal(para, first.Level, first.ListStart);
		} else if (LegalLists.IsLegal(para.ListType)) {
			para.ListType = first.ListType;
			para.SetIndents(0, 0, para.RightIndent);
			var now = ParagraphAt(first.Start);
			if (now != null && LegalLists.Prefix.Match(now.Text) is { Success: true } m && !first.Text.StartsWith(m.Value))
				ed.TextDocument.Range2(now.Start, now.Start + m.Length).Text = "";
		}
	}

	LegalLists.Paragraph? ParagraphAt(int position) =>
		LegalLists.Paragraphs(ed.TextDocument).FirstOrDefault(p => position >= p.Start && position < Math.Max(p.End, p.Start + 1));

	// New tables are as wide as the editor shows text (at most 6.5 inches, the width of a Letter page's text).
	int TableWidth() => ed.ClientSize.Width > 200 ? Math.Min(9360, (ed.ClientSize.Width - 40) * 1440 / Math.Max(96, ed.DeviceDpi)) : 9360;
	/// The document as Markdown. With embedImages = 1 (the default) images are data: URIs, so the text is complete on
	/// its own; with 0 they are written as image1.png, image2.png ... (as when viewing the source).
	public string GetMarkdown(bool embedImages = true) =>
		MarkdownConverter.ToMarkdown(ed, embedImages ? null : (number, data, extension) => $"image{number}.{extension}");

	/// Replaces the document with Markdown. Relative image paths are resolved from the current file's folder.
	public void SetMarkdown(string markdown) =>
		MarkdownConverter.Load(ed, markdown, form.FileName != null ? Path.GetDirectoryName(form.FileName) : Environment.CurrentDirectory);

	/// The document as Markdown (get), or replaced by Markdown (set); see GetMarkdown and SetMarkdown.
	public string Markdown {
		get => GetMarkdown();
		set => SetMarkdown(value);
	}

	// Saving document.md writes its images to document_images\image1.png ... and refers to them from the Markdown.
	// Images left from an earlier save are removed first.
	static Func<int, byte[], string, string> ImageFiles(string markdownPath) {
		string name = Path.GetFileNameWithoutExtension(markdownPath) + "_images";
		string folder = Path.Combine(Path.GetDirectoryName(markdownPath)!, name);
		if (Directory.Exists(folder))
			foreach (var old in Directory.GetFiles(folder).Where(f => Regex.IsMatch(Path.GetFileName(f), @"^image\d+\.(png|jpg|gif|bmp|tiff)$", RegexOptions.IgnoreCase)))
				File.Delete(old);
		return (number, data, extension) => {
			Directory.CreateDirectory(folder);
			File.WriteAllBytes(Path.Combine(folder, $"image{number}.{extension}"), data);
			return Uri.EscapeDataString(name) + $"/image{number}.{extension}";
		};
	}

	// ---- Text and selection

	/// Inserts text at the caret, replacing the selection; the caret ends up after the new text.
	public void InsertText(string text) {
		ed.SelectedText = text;
		ed.ScrollToCaret();
	}

	/// Selects length characters from start (0-based).
	public void Select(int start, int length = 0) {
		ed.Select(Math.Clamp(start, 0, ed.TextLength), Math.Max(0, length));
		ed.ScrollToCaret();
	}

	/// Moves the caret; a negative position goes to the end of the document.
	public void GoTo(int position) => Select(position < 0 ? ed.TextLength : position);

	/// Finds the next occurrence after the selection (wrapping round), selects it and returns its position, or -1.
	public int Find(string text, bool matchCase = false, bool wholeWord = false) {
		if (string.IsNullOrEmpty(text)) throw new ArgumentException("There is nothing to find.");
		var options = (matchCase ? RichTextBoxFinds.MatchCase : 0) | (wholeWord ? RichTextBoxFinds.WholeWord : 0);
		int from = ed.SelectionStart + ed.SelectionLength;
		int index = from < ed.TextLength ? ed.Find(text, from, options) : -1;
		if (index < 0) index = ed.Find(text, 0, options);
		if (index >= 0) ed.ScrollToCaret();
		return index;
	}

	/// Replaces every occurrence and returns the number of replacements.
	public int Replace(string text, string replacement, bool matchCase = false, bool wholeWord = false) {
		if (string.IsNullOrEmpty(text)) throw new ArgumentException("There is nothing to find.");
		var options = (matchCase ? RichTextBoxFinds.MatchCase : 0) | (wholeWord ? RichTextBoxFinds.WholeWord : 0);
		int count = 0, position = 0;
		while (position < ed.TextLength) {
			int index = ed.Find(text, position, options);
			if (index < 0) break;
			ed.SelectedText = replacement;
			count++;
			position = index + replacement.Length;
		}
		return count;
	}

	public int WordCount() => Regex.Matches(ed.Text, @"[\p{L}\p{N}][\p{L}\p{N}'’\-]*").Count;
	public int CharacterCount() => ed.TextLength;

	/// Changes the case of the selection, keeping its formatting: "upper", "lower", "title", "sentence" or "toggle".
	public void ChangeCase(string mode) {
		int type = mode.ToLowerInvariant() switch {
			"upper" => tomConstants.tomUpperCase,
			"lower" => tomConstants.tomLowerCase,
			"title" => tomConstants.tomTitleCase,
			"sentence" => tomConstants.tomSentenceCase,
			"toggle" => tomConstants.tomToggleCase,
			_ => throw new ArgumentException("Case must be upper, lower, title, sentence or toggle.")
		};
		SelectionRange().ChangeCase(type);
	}

	// ---- Formatting (these wrap RichEditBox methods of the same name)

	/// RichEditorNET 1.0.11: ToggleSuperscript sets Superscript and then clears Subscript, and RichEdit clears both,
	/// so turning it on has no effect. If that happens, the format is set through TOM (fixed upstream in the PR).
	public void ToggleSuperscript() => ToggleScript(true);
	public void ToggleSubscript() => ToggleScript(false);

	void ToggleScript(bool superscript) {
		int IsOn() { var font = SelectionRange().Font; return superscript ? font.Superscript : font.Subscript; }
		bool wasOn = IsOn() == tomConstants.tomTrue;
		if (superscript) ed.ToggleSuperscript(); else ed.ToggleSubscript();
		if (wasOn || IsOn() == tomConstants.tomTrue) return;
		var range = SelectionRange().Font;
		if (superscript) range.Superscript = tomConstants.tomTrue; else range.Subscript = tomConstants.tomTrue;
	}

	/// Toggles a list; numbered lists start at 1 (RichEditorNET 1.0.11 starts new decimal lists at 0).
	public void ToggleList(ListStyle style) {
		if (InLegalList()) {    // a legal list becomes a list of this style (its numbers go)
			LegalLists.Remove(ed, SelectionRange().Start, SelectionRange().End);
		}
		ed.ToggleList(style);
		var para = SelectionRange().Para;
		if (style != ListStyle.Bullet && (para.ListType & 0xFFFF) == (int)style && para.ListStart == 0) para.ListStart = 1;
	}

	/// Sets the block style of the selected paragraphs (needs the structured mode, EnableHtmlFontSizing = 1).
	public void SetBlockStyle(BlockStyle style) {
		if (!ed.EnableHtmlFontSizing) throw new InvalidOperationException("Block styles need a structured document (View > Structured document, or EnableHtmlFontSizing = 1).");
		ed.SetBlockStyle(style);
		// RichEditorNET formats only the paragraph's text, so on an empty paragraph the style would be lost: apply it to
		// the paragraph mark and the caret too, so that what is typed next gets it.
		var para = ed.TextDocument.Range2(ed.SelectionStart, ed.SelectionStart);
		para.Expand(tomConstants.tomParagraph);
		if (ed.SelectionLength == 0 && (para.Text ?? "").Trim('\r', '\n') == "") {
			ApplyBlockFont(ed.TextDocument.Range2(para.Start, para.End).Font, style);
			ApplyBlockFont(ed.TextDocument.Selection2.Font, style);
		}
	}

	// The library's block style fonts (internal), or the same values if they cannot be reached.
	static readonly Type? blockStyleHelper = typeof(RichEditBox).Assembly.GetType("EllipticBit.RichEditorNET.Formatting.BlockStyleHelper");
	static T BlockConstant<T>(string name, T fallback) =>
		blockStyleHelper?.GetField(name, BindingFlags.NonPublic | BindingFlags.Static)?.GetValue(null) is T value ? value : fallback;

	void ApplyBlockFont(ITextFont font, BlockStyle style) {
		switch (style) {
			case BlockStyle.Preformatted:
				font.Bold = tomConstants.tomFalse;
				font.Size = BlockConstant("PreformattedPointSize", 10f);
				font.Name = BlockConstant("PreformattedFontName", "Courier New");
				break;
			case BlockStyle.Paragraph:
				font.Bold = tomConstants.tomFalse;
				font.Size = BlockConstant("ParagraphPointSize", 12f);
				font.Name = ed.Font.Name;
				break;
			default:
				font.Bold = tomConstants.tomTrue;
				font.Size = BlockConstant("HeadingPointSizes", new[] { 0f, 24f, 18f, 14f, 12f, 10f, 8f })[(int)style];
				font.Name = ed.Font.Name;
				break;
		}
	}

	public void ClearFontColor() => SelectionRange().Font.ForeColor = tomConstants.tomAutoColor;
	public void ClearHighlight() => SelectionRange().Font.BackColor = tomConstants.tomAutoColor;

	/// True if the caret is in a list item.
	public bool InList() => (SelectionRange().Para.ListType & 0xFFFF) != tomConstants.tomListNone;

	/// Nests (positive) or un-nests (negative) the selected list items by that many levels, like Tab / Shift+Tab.
	/// With automatic level styles, numbered items take the style of their new level: 1. -> a. -> i. -> 1. ...
	public void Nest(int levels = 1) {
		if (TouchesLegalList()) {    // legal items: their new levels' indents and numbers, in one undo step
			var range = SelectionRange();
			int start = range.Start, end = range.End;
			LegalLists.Grouped(ed, () => {
				for (int i = 0; i < Math.Abs(levels); i++) NestOnce(Math.Sign(levels));
				LegalLists.Reindent(ed.TextDocument, start, Math.Max(start, end));
			});
			return;
		}
		for (int i = 0; i < Math.Abs(levels); i++) NestOnce(Math.Sign(levels));
	}


	void NestOnce(int delta) {
		int selStart = ed.SelectionStart, selLength = ed.SelectionLength, selEnd = selStart + selLength;
		var starts = new List<int>();
		var r = ed.TextDocument.Range2(selStart, selStart);
		r.StartOf(tomConstants.tomParagraph, 0);
		do starts.Add(r.Start);
		while (selLength > 0 && r.Move(tomConstants.tomParagraph, 1) != 0 && r.Start < selEnd && r.Start < ed.TextLength);
		try {
			foreach (int start in starts) {
				ed.Select(start, 0);
				var para = SelectionRange().Para;
				if ((para.ListType & 0xFFFF) == tomConstants.tomListNone) continue;
				int before = para.ListLevelIndex;
				if (delta > 0) ed.IncreaseListLevel(); else ed.DecreaseListLevel();
				para = SelectionRange().Para;
				int level = para.ListLevelIndex;
				// numbered items take the numbering of their new level from Paragraph > List levels
				if (level != before && form.AutoLevelStyles && (para.ListType & 0xFFFF) != tomConstants.tomListBullet) ApplyLevel(para, level);
			}
		} finally {
			ed.Select(selStart, selLength);
		}
	}

	// ---- Images

	/// Inserts a PNG, JPEG or GIF image at the caret. Without a width and height it keeps the image's own size,
	/// scaled down to at most 600 pixels wide; with just one of them, the other follows the image's proportions.
	/// (RichEditorNET shows inserted images at DefaultImageWidth x DefaultImageHeight, so those are set for the insert.)
	public void InsertImageFile(string path, string altText = "", int width = 0, int height = 0) {
		byte[] data = File.ReadAllBytes(path);
		using (var buffer = new MemoryStream(data))
		using (var image = Image.FromStream(buffer)) {
			double aspect = (double)image.Height / image.Width;
			if (width <= 0 && height <= 0) {
				width = Math.Min(image.Width, 600);
				height = (int)Math.Round(width * aspect);
			} else if (height <= 0) {
				height = (int)Math.Round(width * aspect);
			} else if (width <= 0) {
				width = (int)Math.Round(height / aspect);
			}
		}
		int oldWidth = ed.DefaultImageWidth, oldHeight = ed.DefaultImageHeight;
		ed.DefaultImageWidth = Math.Max(1, width);
		ed.DefaultImageHeight = Math.Max(1, height);
		try {
			using var stream = new MemoryStream(data);
			ed.InsertImage(stream, altText);
		} finally {
			ed.DefaultImageWidth = oldWidth;
			ed.DefaultImageHeight = oldHeight;
		}
	}

	public void InsertLinkedThumbnailFile(string imageUrl, string path, string altText = "") {
		using var stream = File.OpenRead(path);
		ed.InsertLinkedThumbnail(imageUrl, stream, altText);
	}

	// ---- Events (RteClient2.ahk: Rte.On / Rte.Off subscribe A_ScriptHwnd with these)

	/// Sends an event to an AHK window (A_ScriptHwnd) from now on.
	public void Subscribe(string eventName, long window) => form.Events.Subscribe(ParseEvent(eventName), (IntPtr)window);

	/// Stops sending an event ("*" for all events) to an AHK window.
	public void Unsubscribe(string eventName, long window) =>
		form.Events.Unsubscribe(eventName == "*" ? null : ParseEvent(eventName), (IntPtr)window);

	/// The event names, one per line.
	public string EventNames => string.Join("\n", Enum.GetNames<EditorEvent>());

	/// The error of the last call made with Rte.Post that failed.
	public string LastAsyncError { get; internal set; } = "";

	static EditorEvent ParseEvent(string name) =>
		Enum.TryParse<EditorEvent>(name, true, out var e) && Enum.IsDefined(e) ? e
			: throw new ArgumentException($"Unknown event \"{name}\". Events: {string.Join(", ", Enum.GetNames<EditorEvent>())}.");

	// ---- Iteration 3: recent files, settings, printing

	/// The recently opened or saved files, newest first, one per line.
	public string RecentFiles => string.Join("\n", form.Settings.RecentFiles);

	/// The settings file (window, view settings, recent files, page setup).
	public string SettingsFile => Settings.FilePath;

	/// Exports the document as a PDF file through the "Microsoft Print to PDF" printer, and returns its full path.
	/// The page setup (File > Page Setup) applies. This can take a few seconds.
	public string ExportPdf(string path) {
		const string pdfPrinter = "Microsoft Print to PDF";
		if (!PrinterSettings.InstalledPrinters.Cast<string>().Contains(pdfPrinter))
			throw new InvalidOperationException($"The \"{pdfPrinter}\" printer is not installed (Windows Features: Microsoft Print to PDF).");
		string full = Path.GetFullPath(path);
		if (File.Exists(full)) File.Delete(full);
		using (var printer = form.CreatePrinter(pdfPrinter)) {
			var document = printer.Document;
			document.PrinterSettings.PrintToFile = true;
			document.PrinterSettings.PrintFileName = full;
			document.PrintController = new StandardPrintController();    // no "Printing..." window
			document.Print();
		}
		// The printer writes the file in the background; wait until it is complete.
		var clock = System.Diagnostics.Stopwatch.StartNew();
		while (clock.ElapsedMilliseconds < 20000) {
			try {
				if (File.Exists(full)) {
					using var file = File.Open(full, FileMode.Open, FileAccess.Read, FileShare.None);
					if (file.Length > 0) return full;
				}
			} catch (IOException) { }
			Thread.Sleep(100);
		}
		throw new TimeoutException("The PDF printer did not finish the file within 20 s.");
	}

	/// Prints without a dialog, to the default printer or the printer named. The page setup applies.
	public void Print(string printerName = "") {
		using var printer = form.CreatePrinter(printerName == "" ? null : printerName);
		printer.Document.PrintController = new StandardPrintController();
		printer.Document.Print();
	}

	// ---- Iteration 8: list levels (Paragraph > List levels)

	/// Makes the selected paragraphs a numbered list (or ends it), numbered as List levels sets for their levels.
	public void NumberedList() {
		if (InLegalList()) {    // a legal list becomes an ordinary numbered list, keeping its levels
			var range = SelectionRange();
			LegalLists.Grouped(ed, () => {
				var doc = ed.TextDocument;
				foreach (var p in Enumerable.Reverse(LegalLists.Touched(doc, range.Start, range.End)).Where(p => p.Legal)) {
					if (p.PrefixLength > 0) doc.Range2(p.Start, p.Start + p.PrefixLength).Text = "";
					var para = doc.Range2(p.Start, p.Start).Para;
					para.ListType = tomConstants.tomListNumberAsArabic | tomConstants.tomListPeriod;
					ApplyLevel(para, p.Level);
				}
			});
			return;
		}
		int type = SelectionRange().Para.ListType & 0xFFFF;
		if (type != tomConstants.tomListNone && type != tomConstants.tomListBullet) {
			ed.ToggleList((ListStyle)type);    // the same style again ends the list
			return;
		}
		ed.ToggleList(ListStyle.Decimal);
		ForEachSelectedParagraph(para => ApplyLevel(para, para.ListLevelIndex));
	}

	/// Applies List levels to the whole list the caret is in: each item gets the numbering of its level.
	public void ApplyListLevels() {
		if (!InList()) throw new InvalidOperationException("Put the caret in a list first.");
		var doc = ed.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		// from the caret's paragraph back and forward over the neighbouring list paragraphs
		var first = doc.Range2(ed.SelectionStart, ed.SelectionStart);
		first.Expand(tomConstants.tomParagraph);
		while (first.Start > 0) {
			var before = doc.Range2(first.Start - 1, first.Start - 1);
			before.Expand(tomConstants.tomParagraph);
			if ((before.Para.ListType & 0xFFFF) == tomConstants.tomListNone) break;
			first = before;
		}
		var range = doc.Range2(first.Start, first.Start);
		range.Expand(tomConstants.tomParagraph);
		while ((range.Para.ListType & 0xFFFF) != tomConstants.tomListNone) {
			var para = range.Para;
			ApplyLevel(para, para.ListLevelIndex);
			if (range.End >= storyEnd || range.End <= range.Start) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
	}

	/// Starts the numbering again at start, from the list item with the caret (that item and the rest of its level).
	/// RichEdit numbers on through items that share a start value, so two lists that start at the same number are kept
	/// apart by an empty paragraph (at the top level; a nested list cannot be split that way).
	public void RestartNumbering(int start = 1) {
		if (InLegalList()) { LegalLists.Restart(ed, ed.SelectionStart, start); return; }
		var items = ListParagraphs(out int index);
		var para = ParaAt(items[index]);
		int type = para.ListType & 0xFFFF;
		if (type == tomConstants.tomListNone || type == tomConstants.tomListBullet) throw new InvalidOperationException("Put the caret in a numbered list item first.");
		int level = para.ListLevelIndex;
		var before = PreviousAtLevel(items, index, level);
		if (before is int previous && ParaAt(previous).ListStart == start && ParaAt(previous).ListType == para.ListType) {
			if (level > 0) throw new InvalidOperationException($"The items above also start at {start}, and RichEdit numbers on through items with the same start; a nested list cannot be split. Choose another start number.");
			// split the list with an empty, normal paragraph
			int at = items[index];
			ed.TextDocument.Range2(at, at).Text = "\r";
			var separator = ParaAt(at);
			separator.ListType = tomConstants.tomListNone;
			separator.SetIndents(0, 0, 0);
			ed.Select(at + 1, 0);    // back in the item (now after the separator)
			items = ListParagraphs(out index);
		}
		foreach (int item in ItemsOfRun(items, index, level)) ParaAt(item).ListStart = start;
	}

	/// Continues the numbering of the previous list of the same kind, for the list item with the caret and the rest of
	/// its list. Empty paragraphs between the two lists are removed, so that they become one list.
	public void ContinueNumbering() {
		if (InLegalList()) { LegalLists.Continue(ed, ed.SelectionStart); return; }
		var items = ListParagraphs(out int index);
		var para = ParaAt(items[index]);
		int type = para.ListType & 0xFFFF;
		if (type == tomConstants.tomListNone || type == tomConstants.tomListBullet) throw new InvalidOperationException("Put the caret in a numbered list item first.");
		int level = para.ListLevelIndex;
		// the first item of the caret's run
		int first = index;
		for (int k = index - 1; k >= 0; k--) {
			var other = ParaAt(items[k]);
			if (other.ListLevelIndex < level) break;
			if (other.ListLevelIndex > level) continue;
			if (other.ListStart != para.ListStart || other.ListType != para.ListType) break;
			first = k;
		}
		var doc = ed.TextDocument;
		if (PreviousAtLevel(items, first, level) is int inSameList) {
			// the item just before belongs to the same block of list paragraphs: take its start value, and the two runs join
			var previous = ParaAt(inSameList);
			if (previous.ListType != para.ListType) throw new InvalidOperationException("The list before has another numbering style; Paragraph > List levels can make them the same.");
			foreach (int item in ItemsOfRun(items, first, level)) ParaAt(item).ListStart = previous.ListStart;
			return;
		}
		if (level > 0) throw new InvalidOperationException("There is no earlier item at this level in this list to continue from.");
		// an earlier top-level list, before non-list paragraphs
		int position = items[0];
		var range = doc.Range2(position, position);
		range.Expand(tomConstants.tomParagraph);
		bool onlyEmpty = true;
		int? found = null;
		while (range.Start > 0) {
			var back = doc.Range2(range.Start - 1, range.Start - 1);
			back.Expand(tomConstants.tomParagraph);
			range = back;
			var p = range.Para;
			if ((p.ListType & 0xFFFF) == tomConstants.tomListNone) {
				if ((range.Text ?? "").Trim('\r', '\n', ' ', '\t') != "") onlyEmpty = false;
				continue;
			}
			if (p.ListLevelIndex == 0) { found = range.Start; break; }
		}
		if (found is not int previousItem || (ParaAt(previousItem).ListType & 0xFFFF) != type)
			throw new InvalidOperationException("There is no earlier numbered list of the same kind to continue.");
		var last = ParaAt(previousItem);
		if (last.ListType != para.ListType) throw new InvalidOperationException("The list before has another numbering style; Paragraph > List levels can make them the same.");
		if (onlyEmpty) {
			// only empty paragraphs between: remove them, so that the two become one list
			var lastRange = doc.Range2(previousItem, previousItem);
			lastRange.Expand(tomConstants.tomParagraph);
			doc.Range2(lastRange.End, items[0]).Text = "";
			items = ListParagraphs(out index);
			// the runs join when they share the start value
			int startValue = last.ListStart;
			int firstOfRun = items.FindIndex(i => i >= lastRange.End);
			foreach (int item in ItemsOfRun(items, Math.Max(0, firstOfRun), 0)) ParaAt(item).ListStart = startValue;
			return;
		}
		int next = DisplayedNumber(previousItem) + 1;
		foreach (int item in ItemsOfRun(items, first, level)) ParaAt(item).ListStart = next;
	}

	ITextPara ParaAt(int position) => ed.TextDocument.Range2(position, position).Para;

	// The starts of the list paragraphs around the caret (a list ends at a paragraph that is not a list item), and the
	// index of the caret's paragraph among them.
	List<int> ListParagraphs(out int index) {
		var doc = ed.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var range = doc.Range2(ed.SelectionStart, ed.SelectionStart);
		range.Expand(tomConstants.tomParagraph);
		if ((range.Para.ListType & 0xFFFF) == tomConstants.tomListNone) throw new InvalidOperationException("Put the caret in a numbered list item first.");
		while (range.Start > 0) {
			var back = doc.Range2(range.Start - 1, range.Start - 1);
			back.Expand(tomConstants.tomParagraph);
			if ((back.Para.ListType & 0xFFFF) == tomConstants.tomListNone) break;
			range = back;
		}
		var starts = new List<int>();
		while ((range.Para.ListType & 0xFFFF) != tomConstants.tomListNone) {
			starts.Add(range.Start);
			if (range.End >= storyEnd || range.End <= range.Start) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
		index = Math.Max(0, starts.FindLastIndex(start => start <= ed.SelectionStart));
		return starts;
	}

	// The item at the same level before items[index] in the same sublist (deeper items are skipped), or null.
	int? PreviousAtLevel(List<int> items, int index, int level) {
		for (int k = index - 1; k >= 0; k--) {
			int otherLevel = ParaAt(items[k]).ListLevelIndex;
			if (otherLevel < level) return null;
			if (otherLevel == level) return items[k];
		}
		return null;
	}

	// items[index] and the items after it at the same level, up to the end of their sublist.
	IEnumerable<int> ItemsOfRun(List<int> items, int index, int level) {
		for (int k = index; k < items.Count; k++) {
			int otherLevel = ParaAt(items[k]).ListLevelIndex;
			if (otherLevel < level) yield break;
			if (otherLevel == level) yield return items[k];
		}
	}

	// The number RichEdit shows for a list item: its start value plus the items before it with the same start value.
	int DisplayedNumber(int position) {
		var doc = ed.TextDocument;
		var para = ParaAt(position);
		int level = para.ListLevelIndex, startValue = para.ListStart, count = 0;
		var range = doc.Range2(position, position);
		range.Expand(tomConstants.tomParagraph);
		while (range.Start > 0) {
			var back = doc.Range2(range.Start - 1, range.Start - 1);
			back.Expand(tomConstants.tomParagraph);
			range = back;
			var p = range.Para;
			if ((p.ListType & 0xFFFF) == tomConstants.tomListNone || p.ListLevelIndex < level) break;
			if (p.ListLevelIndex > level) continue;
			if (p.ListStart != startValue || p.ListType != para.ListType) break;
			count++;
		}
		return startValue + count;
	}

	/// The list levels, one per line: level (1-9), style, punctuation and start number, separated by tabs.
	public string ListLevels => string.Join("\n", form.Settings.ListLevels.Select((l, i) => $"{i + 1}\t{l.Style}\t{l.Punctuation}\t{l.Start}"));

	/// Sets a list level (1-9): style Decimal, LowerAlpha, UpperAlpha, LowerRoman, UpperRoman or Bullet; punctuation
	/// ".", ")", "()" or ""; and the start number.
	public void SetListLevel(int level, string style, string punctuation = ".", int start = 1) {
		if (level < 1 || level > 9) throw new ArgumentOutOfRangeException(nameof(level), "Levels are 1 to 9.");
		var parsed = Enum.Parse<ListStyle>(style, ignoreCase: true);
		if (!ListLevel.Punctuations.Contains(punctuation)) throw new ArgumentException("Punctuation must be \".\", \")\", \"()\" or \"\".");
		var levels = form.Settings.ListLevels.ToList();
		while (levels.Count < 9) levels.Add(ListLevel.Defaults()[levels.Count]);
		levels[level - 1] = new ListLevel { Style = parsed.ToString(), Punctuation = punctuation, Start = Math.Max(0, start) };
		form.Settings.ListLevels = levels;
		form.Settings.Save();
	}

	// One list paragraph gets the numbering of a level from the scheme.
	void ApplyLevel(ITextPara para, int level) {
		if (LegalLists.IsLegal(para.ListType)) return;    // legal items keep their numbering
		var levels = form.Settings.ListLevels;
		var scheme = level < levels.Count ? levels[level] : ListLevel.Defaults()[Math.Min(level, 8)];
		para.ListType = scheme.ListType;
		para.ListStart = scheme.Start;
	}

	void ForEachSelectedParagraph(Action<ITextPara> action) {
		int selStart = ed.SelectionStart, selEnd = selStart + ed.SelectionLength;
		var range = ed.TextDocument.Range2(selStart, selStart);
		range.Expand(tomConstants.tomParagraph);
		while (true) {
			action(range.Para);
			if (range.End >= selEnd || range.End <= range.Start || range.End >= ed.TextLength) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
	}

	// ---- Iteration 6: keys and macros

	/// Presses a key, as recorded macros do: Left, Right, Up, Down, Home, End, Backspace, Delete, Enter or Tab, optionally
	/// with Ctrl+ and/or Shift+ ("Ctrl+Shift+Left"). Shift extends the selection; Ctrl moves by words, paragraphs or the
	/// whole document (Home/End). count repeats the key.
	public void Key(string key, int count = 1) {
		var parts = key.Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
		bool ctrl = parts.Any(p => p.Equals("Ctrl", StringComparison.OrdinalIgnoreCase) || p.Equals("Control", StringComparison.OrdinalIgnoreCase));
		bool shift = parts.Any(p => p.Equals("Shift", StringComparison.OrdinalIgnoreCase));
		string name = parts.Length > 0 ? parts[^1].ToLowerInvariant() : "";
		var selection = ed.TextDocument.Selection2;
		int extend = shift ? tomConstants.tomExtend : tomConstants.tomMove;
		for (int i = 0; i < Math.Max(1, count); i++) {
			switch (name) {
				case "left": selection.MoveLeft(ctrl ? tomConstants.tomWord : tomConstants.tomCharacter, 1, extend); break;
				case "right": selection.MoveRight(ctrl ? tomConstants.tomWord : tomConstants.tomCharacter, 1, extend); break;
				case "up": selection.MoveUp(ctrl ? tomConstants.tomParagraph : tomConstants.tomLine, 1, extend); break;
				case "down": selection.MoveDown(ctrl ? tomConstants.tomParagraph : tomConstants.tomLine, 1, extend); break;
				case "home": selection.HomeKey(ctrl ? tomConstants.tomStory : tomConstants.tomLine, extend); break;
				case "end": selection.EndKey(ctrl ? tomConstants.tomStory : tomConstants.tomLine, extend); break;
				case "backspace": if (ctrl || !TouchesLegalList() || !LegalBackspace()) selection.Delete(ctrl ? tomConstants.tomWord : tomConstants.tomCharacter, selection.End > selection.Start ? 1 : -1); break;
				case "delete": if (ctrl || !TouchesLegalList() || !LegalDelete()) selection.Delete(ctrl ? tomConstants.tomWord : tomConstants.tomCharacter, 1); break;
				case "enter": if (!TouchesLegalList() || !LegalEnter()) selection.TypeText("\r"); break;
				case "tab": if (InTable()) NextCell(shift ? -1 : 1); else if (InList()) Nest(shift ? -1 : 1); else selection.TypeText("\t"); break;
				default: throw new ArgumentException($"Unknown key \"{key}\". Keys: Left, Right, Up, Down, Home, End, Backspace, Delete, Enter, Tab, with Ctrl+ and/or Shift+.");
			}
		}
		ed.ScrollToCaret();
	}

	/// Runs a macro from the rte_macros folder by name (without .ahk); it runs as a separate script.
	public void RunMacro(string name) => form.RunScript(form.MacroPath(name));

	/// The macros, one per line: file name and hotkey (if any), separated by a tab.
	public string Macros => string.Join("\n", form.MacroFiles().Select(f => f + "\t" + (form.HotkeyOf(f) is Keys k ? HotkeyBox.Describe(k) : "")));

	/// Sets the hotkey of a macro ("Ctrl+Alt+1", "F9", ...), or removes it with "". Works while the editor is active.
	public void SetMacroHotkey(string name, string keys) {
		string file = Path.GetFileName(form.MacroPath(name));
		var parsed = keys.Trim() == "" ? Keys.None : HotkeyBox.Parse(keys) ?? throw new ArgumentException($"\"{keys}\" is not a key combination (examples: Ctrl+Alt+1, F9, Ctrl+Shift+M).");
		form.SetMacroHotkey(file, parsed);
	}

	/// Starts recording a macro, as Tools > Record macro does.
	public void StartRecording() => form.StartRecording();

	/// Stops recording and returns the recorded AHK script ("" if nothing was recorded), without saving it.
	public string StopRecording() => form.StopRecording();

	/// Whether a macro is being recorded.
	public bool IsRecording => form.IsRecording;

	// ---- Iteration 5: spelling (the Windows spell checker, which also draws the red underlines)

	/// The spell checker's language, e.g. en-US ("" if the Windows spell checker is not available).
	public string SpellingLanguage => form.Spelling?.Language ?? "";

	/// Selects the next spelling error after the caret (wrapping round) and returns the word, or "" if there is none.
	public string NextSpellingError() {
		var spelling = RequireSpelling();
		int caret = ed.SelectionStart + ed.SelectionLength;
		var doc = ed.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		// From the caret's paragraph to the end, then from the start up to the caret.
		foreach (var (from, to) in new[] { (caret, storyEnd), (0, caret) }) {
			var para = doc.Range2(from, from);
			para.Expand(tomConstants.tomParagraph);
			while (para.Start < to) {
				int start = para.Start;
				foreach (var error in spelling.Check(para.Text ?? "", start))
					if (error.Start >= from && error.Start < to) {
						ed.Select(error.Start, error.Length);
						ed.ScrollToCaret();
						return error.Word;
					}
				if (para.End >= storyEnd || para.End <= start) break;
				para.SetRange(para.End, para.End);
				para.Expand(tomConstants.tomParagraph);
			}
		}
		return "";
	}

	/// All spelling errors, one per line: start, length and word, separated by tabs.
	public string SpellingErrors() {
		var spelling = RequireSpelling();
		var lines = new List<string>();
		ForEachParagraph((start, text) => {
			foreach (var error in spelling.Check(text, start)) lines.Add($"{error.Start}\t{error.Length}\t{error.Word}");
		});
		return string.Join("\n", lines);
	}

	/// Spelling suggestions for a word, one per line.
	public string Suggestions(string word) => string.Join("\n", RequireSpelling().Suggest(word));

	/// Adds a word to the Windows dictionary (shared with other programs), and refreshes the red underlines.
	public void AddToDictionary(string word) {
		RequireSpelling().Add(word.Trim());
		if (ed.EnableSpellCheck) {    // re-check, so the underline goes
			ed.EnableSpellCheck = false;
			ed.EnableSpellCheck = true;
		}
	}

	/// Replaces length characters at start with text, keeping the formatting of the first replaced character.
	public void ReplaceRange(int start, int length, string text) {
		var range = ed.TextDocument.Range2(start, start + length);
		range.Text = text;
		ed.Select(start + text.Length, 0);
	}

	internal Spelling.Error? SpellingErrorAt(int position) {
		if (form.Spelling == null) return null;
		var para = ed.TextDocument.Range2(position, position);
		para.Expand(tomConstants.tomParagraph);
		foreach (var error in form.Spelling.Check(para.Text ?? "", para.Start))
			if (position >= error.Start && position <= error.Start + error.Length) return error;
		return null;
	}

	internal void DeleteRepeatedWord(Spelling.Error error) {
		// Delete the word together with the space before it.
		int start = error.Start, length = error.Length;
		if (start > 0 && ed.TextDocument.Range2(start - 1, start).Text == " ") { start--; length++; }
		ReplaceRange(start, length, "");
	}

	Spelling RequireSpelling() => form.Spelling ?? throw new InvalidOperationException("The Windows spell checker is not available.");

	void ForEachParagraph(Action<int, string> action) {
		var doc = ed.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var para = doc.Range2(0, 0);
		para.Expand(tomConstants.tomParagraph);
		while (true) {
			int start = para.Start;
			action(start, para.Text ?? "");
			if (para.End >= storyEnd || para.End <= start) break;
			para.SetRange(para.End, para.End);
			para.Expand(tomConstants.tomParagraph);
		}
	}

	// ---- Iteration 5: outline

	/// The headings, one per line: level (1-6), start position and text, separated by tabs.
	public string Outline => string.Join("\n", form.CurrentOutline().Select(h => $"{h.Level}\t{h.Start}\t{h.Text}"));

	/// Goes to the n-th heading (1 = the first) and scrolls it to the top.
	public void GoToHeading(int number) {
		var headings = form.CurrentOutline();
		if (number < 1 || number > headings.Count) throw new ArgumentOutOfRangeException(nameof(number), $"There are {headings.Count} headings.");
		form.GoToPosition(headings[number - 1].Start);
	}

	/// Whether the outline panel is shown.
	public bool OutlineVisible {
		get => form.OutlineVisible;
		set => form.SetOutlineVisible(value);
	}

	/// Changes the heading level (1-6, or 0 for normal text) of the paragraph at a position.
	public void SetHeadingLevel(int position, int level) {
		int selStart = ed.SelectionStart, selLength = ed.SelectionLength;
		ed.Select(position, 0);
		try {
			SetBlockStyle(level <= 0 ? BlockStyle.Paragraph : (BlockStyle)Math.Min(6, level));
		} finally {
			ed.Select(selStart, selLength);
		}
	}

	internal List<(int Level, int Start, string Text)> Headings() {
		var headings = new List<(int, int, string)>();
		var doc = ed.TextDocument;
		ForEachParagraph((start, text) => {
			text = text.Trim();
			if (text == "") return;
			var font = doc.Range2(start, start + 1).Font;
			if (BlockStyleOf(font.Size, font.Bold == tomConstants.tomTrue, font.Name) is BlockStyle style
				&& style >= BlockStyle.Heading1 && style <= BlockStyle.Heading6)
				headings.Add(((int)style, start, text));
		});
		return headings;
	}

	// The library's own rule for recognising block styles (internal), or the same sizes if it cannot be reached.
	static readonly MethodInfo? blockStyleFromFont = typeof(RichEditBox).Assembly
		.GetType("EllipticBit.RichEditorNET.Formatting.BlockStyleHelper")?.GetMethod("GetBlockStyleFromFont", BindingFlags.NonPublic | BindingFlags.Static);

	static BlockStyle? BlockStyleOf(float size, bool bold, string? name) {
		if (blockStyleFromFont != null) return (BlockStyle)blockStyleFromFont.Invoke(null, new object?[] { size, bold, name })!;
		int level = Array.IndexOf(new[] { 0f, 24f, 18f, 14f, 12f, 10f, 8f }, size);
		return bold && level > 0 ? (BlockStyle)level : null;
	}

	// ---- The document's header and footer (fields: {title} {file} {page} {pages} {date} {time}); printed, and saved
	// with the document (iteration 13)

	void BandsChanged() => ed.Modified = true;

	/// Sets the document's header (left, center and right parts) and shows it.
	public void SetHeader(string left = "", string center = "", string right = "") {
		form.Bands.Header = new[] { left, center, right };
		form.Bands.ShowHeader = true;
		BandsChanged();
	}

	/// Sets the document's footer (left, center and right parts) and shows it.
	public void SetFooter(string left = "", string center = "", string right = "") {
		form.Bands.Footer = new[] { left, center, right };
		form.Bands.ShowFooter = true;
		BandsChanged();
	}

	/// Whether the document has a header (printed, and in its files).
	public bool ShowHeader {
		get => form.Bands.ShowHeader;
		set { form.Bands.ShowHeader = value; BandsChanged(); }
	}

	/// Whether the document has a footer.
	public bool ShowFooter {
		get => form.Bands.ShowFooter;
		set { form.Bands.ShowFooter = value; BandsChanged(); }
	}

	/// The header's left, center and right parts, one per line.
	public string HeaderParts => string.Join("\n", PageBands.Parts(form.Bands.Header));

	/// The footer's left, center and right parts, one per line.
	public string FooterParts => string.Join("\n", PageBands.Parts(form.Bands.Footer));

	/// Whether the header and footer are shown above and below the text (View > Header and footer).
	public bool HeaderFooterVisible {
		get => form.Settings.ShowBands;
		set { form.Settings.ShowBands = value; form.Settings.Save(); }
	}

	/// Makes the document's header and footer the default for new documents (and documents without their own).
	public void MakeHeaderFooterDefault() => form.MakeDefault(form.Bands);

	// ---- Window

	/// Shows a message in the status bar.
	public void SetStatus(string text) => form.SetStatus(text);

	/// Brings the editor window to the front.
	public void Activate() {
		if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
		form.Activate();
		ed.Focus();
	}

	/// Closes the editor after replying. With discardChanges = 0, the user is asked about unsaved changes.
	public void Close(bool discardChanges = false) => form.CloseLater(discardChanges);

	/// Lists everything AHK scripts can call.
	public string Help() => Automation.Help();

	// ---- Helpers

	// The selection as TOM has it (WinForms' SelectionLength is short across table rows).
	internal ITextRange2 SelectionRange() {
		var selection = ed.TextDocument.Selection2;
		return ed.TextDocument.Range2(selection.Start, selection.End);
	}

	static readonly MethodInfo? getBlockStyle = typeof(RichEditBox).GetMethod("GetBlockStyle", BindingFlags.NonPublic | BindingFlags.Instance);
	internal BlockStyle? GetBlockStyle() => getBlockStyle?.Invoke(ed, null) as BlockStyle?;

	T WithGithubMarkdown<T>(Func<T> action) {
		bool common = ed.EnableCommonMarkdown, github = ed.EnableGithubMarkdown;
		ed.EnableGithubMarkdown = true;
		try {
			return action();
		} finally {
			ed.EnableCommonMarkdown = common;
			ed.EnableGithubMarkdown = github;
		}
	}

	static string HtmlDocument(string body, string title, string css = "", string meta = "") =>
		"<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<meta charset=\"utf-8\">\r\n" + meta + "<title>" + WebUtility.HtmlEncode(title) + "</title>\r\n"
		+ HtmlLists.StyleElement(css) + "</head>\r\n<body>\r\n"
		+ body + "\r\n</body>\r\n</html>\r\n";

	static string BodyOf(string html) {
		var match = Regex.Match(html, @"<body[^>]*>(.*)</body>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
		return (match.Success ? match.Groups[1].Value : html).Trim();
	}
}

// Calls EditorApi / RichEditBox members by name with text arguments, and formats the results as text.
static class Automation
{
	// Inherited RichTextBox / TextBoxBase members that are useful to scripts (members declared on RichEditBox are always available).
	static readonly HashSet<string> EditorMembers = new(StringComparer.OrdinalIgnoreCase) {
		"Text", "TextLength", "SelectedText", "SelectionStart", "SelectionLength", "SelectAll", "Undo", "Redo", "CanUndo", "CanRedo",
		"Cut", "Copy", "Paste", "Clear", "AppendText", "ScrollToCaret", "Modified", "ReadOnly", "WordWrap", "ZoomFactor",
		"GetLineFromCharIndex", "GetFirstCharIndexFromLine", "Lines"
	};

	const BindingFlags Flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase;

	public static string Dispatch(EditorApi api, RichEditBox editor, string[] request) {
		try {
			if (request.Length < 2) throw new ArgumentException("A request needs a verb (call, get or set) and a member name.");
			string verb = request[0], name = request[1];
			string[] args = request[2..];
			var members = Members(api, editor, name).ToList();
			if (members.Count == 0) throw new MissingMemberException($"Unknown member \"{name}\". Rte.Help() lists the available members.");
			object? result = verb switch {
				"call" => Call(members, name, args),
				"get" => Get(members, name, args),
				"set" => Set(members, name, args),
				_ => throw new ArgumentException($"Unknown verb \"{verb}\".")
			};
			return "0" + Format(result);
		} catch (Exception ex) {
			while (ex is TargetInvocationException { InnerException: { } inner }) ex = inner;
			return "1" + ex.Message;
		}
	}

	static IEnumerable<(object Target, MemberInfo Member)> Members(EditorApi api, RichEditBox editor, string name) {
		foreach (var member in typeof(EditorApi).GetMember(name, Flags))
			yield return (api, member);
		foreach (var member in typeof(RichEditBox).GetMember(name, Flags))
			if (member.DeclaringType == typeof(RichEditBox) || EditorMembers.Contains(member.Name))
				yield return (editor, member);
	}

	static object? Call(List<(object Target, MemberInfo Member)> members, string name, string[] args) {
		Exception? conversionError = null;
		foreach (var (target, member) in members) {
			if (member is MethodInfo method && !method.IsSpecialName) {
				var parameters = method.GetParameters();
				int required = parameters.Count(p => !p.HasDefaultValue);
				if (args.Length < required || args.Length > parameters.Length) continue;
				object?[] values;
				try {
					values = parameters.Select((p, i) => i < args.Length ? Convert(args[i], p.ParameterType) : p.DefaultValue).ToArray();
				} catch (Exception ex) when (ex is FormatException or ArgumentException or NotSupportedException or OverflowException) {
					conversionError = ex;
					continue;
				}
				return method.Invoke(target, values);
			}
			if (member is PropertyInfo property)    // Rte.Text() / Rte.Text("value") also work
				return args.Length == 0 ? property.GetValue(target) : SetProperty(target, property, args[0]);
		}
		if (conversionError != null) throw conversionError;
		var counts = members.Select(m => m.Member).OfType<MethodInfo>().Select(m => m.GetParameters().Length).Distinct();
		throw new ArgumentException($"{name} does not take {args.Length} argument{(args.Length == 1 ? "" : "s")} (it takes {string.Join(" or ", counts)}).");
	}

	static object? Get(List<(object Target, MemberInfo Member)> members, string name, string[] args) {
		foreach (var (target, member) in members) {
			if (member is PropertyInfo { CanRead: true } property && property.GetIndexParameters().Length == 0) return property.GetValue(target);
			if (member is MethodInfo method && !method.IsSpecialName && method.GetParameters().All(p => p.HasDefaultValue))
				return Call(members, name, args);    // Rte.WordCount without parentheses
		}
		throw new MissingMemberException($"{name} is not a readable property.");
	}

	static object? Set(List<(object Target, MemberInfo Member)> members, string name, string[] args) {
		if (args.Length != 1) throw new ArgumentException("A property assignment needs exactly one value.");
		foreach (var (target, member) in members)
			if (member is PropertyInfo { CanWrite: true } property) return SetProperty(target, property, args[0]);
		throw new MissingMemberException($"{name} is not a writable property.");
	}

	static object? SetProperty(object target, PropertyInfo property, string value) {
		property.SetValue(target, Convert(value, property.PropertyType));
		return null;
	}

	static object? Convert(string text, Type type) {
		if (Nullable.GetUnderlyingType(type) is { } underlying) return text == "" ? null : Convert(text, underlying);
		if (type == typeof(string) || type == typeof(object)) return text;
		if (type == typeof(bool)) return text.Trim().ToLowerInvariant() switch {
			"1" or "true" or "yes" or "on" => true,
			"0" or "false" or "no" or "off" or "" => false,
			_ => throw new FormatException($"\"{text}\" is not a boolean (use 1 or 0).")
		};
		if (type.IsEnum) return Enum.Parse(type, text.Trim(), ignoreCase: true);
		if (type == typeof(Color)) {
			string t = text.Trim();
			if (t.StartsWith("0x", StringComparison.OrdinalIgnoreCase)) return Color.FromArgb(255, Color.FromArgb(int.Parse(t[2..], NumberStyles.HexNumber)));
			return ColorTranslator.FromHtml(t);
		}
		if (typeof(IConvertible).IsAssignableFrom(type)) return System.Convert.ChangeType(text.Trim(), type, CultureInfo.InvariantCulture);
		throw new NotSupportedException($"A {type.Name} cannot be passed from AHK.");
	}

	static bool CanConvert(Type type) {
		type = Nullable.GetUnderlyingType(type) ?? type;
		return type == typeof(string) || type == typeof(object) || type.IsEnum || type == typeof(Color) || typeof(IConvertible).IsAssignableFrom(type);
	}

	static string Format(object? value) => value switch {
		null => "",
		string s => s,
		bool b => b ? "1" : "0",
		Color c => $"#{c.R:X2}{c.G:X2}{c.B:X2}",
		Enum e => e.ToString(),
		IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
		IEnumerable items => string.Join("\n", items.Cast<object?>().Select(Format)),
		_ => value.ToString() ?? ""
	};

	public static string Help() {
		var lines = new List<string> {
			"Call any of these from AHK after #Include \"RteClient2.ahk\":",
			"  Rte.Name(args...)  calls a method      Rte.Name  reads a property      Rte.Name := value  sets one",
			"Arguments: text, numbers, 1/0 for booleans, enum names (Heading1, LowerAlpha, Wavy...), colors as #RRGGBB or names.",
			"RteClient2.ahk also has:",
			"  Rte.Post(Name, args...)  queues a call and returns at once (errors: Rte.LastAsyncError / the AsyncError event)",
			"  Rte.On(Event, Callback) / Rte.Off(Event [, Callback])  calls Callback(e) when an event happens",
			"Events and their e properties (all have e.Name):",
			"  TextChanged (e.TextLength)   SelectionChanged (e.Start, e.Length)   ModifiedChanged (e.Modified)",
			"  Saved   Opened (also New)   Closed   AsyncError",
			""
		};
		void Section(string title, Type type, Func<MemberInfo, bool> include) {
			lines.Add(title);
			var members = type.GetMembers(BindingFlags.Public | BindingFlags.Instance)
				.Where(m => m is MethodInfo { IsSpecialName: false } || m is PropertyInfo)
				.Where(m => m.DeclaringType != typeof(object) && include(m))
				.Where(m => m is not MethodInfo method || method.GetParameters().All(p => CanConvert(p.ParameterType)))    // e.g. not InsertImage(Stream)
				.GroupBy(m => Describe(m)).Select(g => g.Key).OrderBy(s => s, StringComparer.OrdinalIgnoreCase);
			lines.AddRange(members.Select(m => "  " + m));
			lines.Add("");
		}
		Section("RTE Writer (EditorApi):", typeof(EditorApi), m => m.DeclaringType == typeof(EditorApi));
		var api = typeof(EditorApi).GetMembers(BindingFlags.Public | BindingFlags.Instance).Select(m => m.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
		Section("Editor (RichEditBox):", typeof(RichEditBox), m => (m.DeclaringType == typeof(RichEditBox) || EditorMembers.Contains(m.Name)) && !api.Contains(m.Name)
			&& !(m is PropertyInfo p && p.PropertyType.Namespace == typeof(ITextDocument2).Namespace));
		return string.Join("\r\n", lines);
	}

	static string Describe(MemberInfo member) => member switch {
		MethodInfo m => $"{m.Name}({string.Join(", ", m.GetParameters().Select(p => $"{TypeName(p.ParameterType)} {p.Name}{(p.HasDefaultValue ? " = " + Format(p.DefaultValue) : "")}"))})"
			+ (m.ReturnType == typeof(void) ? "" : " -> " + TypeName(m.ReturnType)),
		PropertyInfo p => $"{p.Name} : {TypeName(p.PropertyType)}{(p.CanWrite && p.SetMethod!.IsPublic ? "" : " (read-only)")}",
		_ => member.Name
	};

	static string TypeName(Type type) => Nullable.GetUnderlyingType(type) is { } u ? TypeName(u) + "?" : type switch {
		_ when type == typeof(string) => "text",
		_ when type == typeof(bool) => "1/0",
		_ when type == typeof(int) || type == typeof(float) || type == typeof(double) => "number",
		_ when type.IsEnum => $"{type.Name} ({string.Join("|", Enum.GetNames(type))})",
		_ when type == typeof(Color) => "color",
		_ => type.Name
	};
}

// Editor events for AHK scripts. Subscribers are windows (an AHK script's A_ScriptHwnd). Each event is posted to them as
// the registered message "RTE_WRITER_EVENT" (wParam = event number, lParam = event data), so the editor never waits for
// a script; windows that no longer exist are dropped. The numbers are part of the protocol (RteClient2.ahk).
enum EditorEvent { TextChanged = 1, SelectionChanged, ModifiedChanged, Saved, Opened, Closed, AsyncError }

sealed class EventHub
{
	public static readonly int Message = RegisterWindowMessage("RTE_WRITER_EVENT");
	readonly Dictionary<EditorEvent, HashSet<IntPtr>> subscribers = new();

	public void Subscribe(EditorEvent e, IntPtr window) {
		if (!IsWindow(window)) throw new ArgumentException("That window does not exist.");
		if (!subscribers.TryGetValue(e, out var windows)) subscribers[e] = windows = new HashSet<IntPtr>();
		windows.Add(window);
	}

	public void Unsubscribe(EditorEvent? e, IntPtr window) {
		foreach (var (key, windows) in subscribers)
			if (e == null || key == e) windows.Remove(window);
	}

	public void Raise(EditorEvent e, long data = 0) {
		if (!subscribers.TryGetValue(e, out var windows)) return;
		foreach (var window in windows.ToArray())
			if (!IsWindow(window) || !PostMessage(window, Message, (IntPtr)(int)e, (IntPtr)data)) windows.Remove(window);
	}

	[DllImport("user32", CharSet = CharSet.Unicode)] static extern int RegisterWindowMessage(string name);
	[DllImport("user32")] static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	[DllImport("user32")] static extern bool IsWindow(IntPtr hWnd);
}

// Settings remembered between sessions, stored as JSON in %APPDATA%\RTE Writer\settings.json, or in the file named by
// the RTE_WRITER_SETTINGS environment variable. A missing or unreadable file gives the defaults.
public sealed class Settings
{
	public const int MaxRecentFiles = 10;
	public int[]? Window { get; set; }                  // normal (not maximized) bounds: x, y, width, height
	public bool Maximized { get; set; }
	public List<string> RecentFiles { get; set; } = new();
	public string? LastFolder { get; set; }
	public bool Structured { get; set; } = true;
	public bool WordWrap { get; set; } = true;
	public bool SpellCheck { get; set; } = true;
	public bool AutoLevelStyles { get; set; } = true;
	public float Zoom { get; set; } = 1f;
	public int[] Margins { get; set; } = { 100, 100, 100, 100 };    // left, right, top, bottom, in 1/100 inch
	public bool Landscape { get; set; }
	// Printed header and footer: left, center and right parts, with the fields {title} {file} {page} {pages} {date} {time}
	public bool ShowHeader { get; set; }
	public string[] Header { get; set; } = DefaultHeader;
	public bool ShowFooter { get; set; } = true;
	public bool ShowOutline { get; set; } = true;
	public bool ShowBands { get; set; } = true;    // View > Header and footer: the strips above and below the text
	public List<ListLevel> ListLevels { get; set; } = ListLevel.Defaults();    // Paragraph > List levels
	public int TableRows { get; set; } = 3;    // Table > Insert table
	public int TableColumns { get; set; } = 3;
	public bool TableHeader { get; set; } = true;
	public string LegalFormat { get; set; } = "Top";    // legal numbering: Top (1. 1.1), All (1. 1.1.) or None (1 1.1)
	public Dictionary<string, string> MacroHotkeys { get; set; } = new(StringComparer.OrdinalIgnoreCase);    // macro file name -> keys, e.g. "Ctrl+Alt+1"
	public int OutlineWidth { get; set; } = 220;    // at 96 dpi
	public string[] Footer { get; set; } = DefaultFooter;
	public static string[] DefaultHeader => new[] { "{title}", "", "{date}" };
	public static string[] DefaultFooter => new[] { "", "Page {page} of {pages}", "" };

	// Always three parts, whatever the file contained.
	public static string[] Parts(string[]? parts) => Enumerable.Range(0, 3).Select(i => parts != null && i < parts.Length ? parts[i] ?? "" : "").ToArray();

	public static string FilePath => Environment.GetEnvironmentVariable("RTE_WRITER_SETTINGS") is { Length: > 0 } path ? path
		: Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RTE Writer", "settings.json");

	public static Settings Load() {
		try {
			if (File.Exists(FilePath)) return JsonSerializer.Deserialize<Settings>(File.ReadAllText(FilePath)) ?? new Settings();
		} catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { }
		return new Settings();
	}

	public void Save() {
		try {
			Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
			File.WriteAllText(FilePath, JsonSerializer.Serialize(this, new JsonSerializerOptions {
				WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping    // "Ctrl+Alt+1", not "Ctrl+Alt+1"
			}));
		} catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
	}
}

// Prints the editor's content: RichEdit lays out each page itself (EM_FORMATRANGE), and a footer shows the page number.
sealed class DocumentPrinter : IDisposable
{
	public readonly PrintDocument Document = new();
	readonly RichEditBox editor;
	readonly string title, path;
	readonly Settings settings;
	readonly PageBands bands;
	int nextChar, page, pageCount;

	const int EM_FORMATRANGE = 0x0400 + 57;
	const float BandHeight = 25, BandGap = 20;    // header / footer line height and distance from the text, 1/100 inch

	[StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }
	[StructLayout(LayoutKind.Sequential)] struct CHARRANGE { public int cpMin, cpMax; }
	[StructLayout(LayoutKind.Sequential)] struct FORMATRANGE { public IntPtr hdc, hdcTarget; public RECT rc, rcPage; public CHARRANGE chrg; }
	[DllImport("user32")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref FORMATRANGE lParam);
	[DllImport("user32")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

	public DocumentPrinter(RichEditBox editor, string title, string path, Settings settings, PageBands bands, string? printerName) {
		this.editor = editor;
		this.bands = bands;
		this.title = title;
		this.path = path;
		this.settings = settings;
		Document.DocumentName = title;
		if (!string.IsNullOrEmpty(printerName)) Document.PrinterSettings.PrinterName = printerName;    // before the page settings
		Document.DefaultPageSettings.Landscape = settings.Landscape;
		if (settings.Margins is { Length: 4 } m) Document.DefaultPageSettings.Margins = new Margins(m[0], m[1], m[2], m[3]);
		Document.BeginPrint += (s, e) => { nextChar = 0; page = 0; pageCount = 0; };
		Document.PrintPage += PrintPage;
		Document.EndPrint += (s, e) => SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero);    // frees RichEdit's cache
	}

	void PrintPage(object? sender, PrintPageEventArgs e) {
		page++;
		var g = e.Graphics!;
		var margins = e.MarginBounds;
		float hardX = e.PageSettings.HardMarginX, hardY = e.PageSettings.HardMarginY;
		int start = nextChar;
		int end = editor.TextDocument.Range2(0, 0).StoryLength - 1;    // without the final paragraph mark
		IntPtr hdc = g.GetHdc();
		try {
			var range = new FORMATRANGE {
				hdc = hdc, hdcTarget = hdc,
				rc = Twips(margins, hardX, hardY), rcPage = Twips(e.PageBounds, hardX, hardY),
				chrg = new CHARRANGE { cpMin = nextChar, cpMax = -1 }
			};
			if (pageCount == 0) pageCount = CountPages(range, end);    // for {pages}
			range.chrg.cpMin = nextChar;
			if (Document.PrintController?.IsPreview == true) nextChar = PreviewPage(range, e.PageBounds, out preview);
			else nextChar = (int)SendMessage(editor.Handle, EM_FORMATRANGE, (IntPtr)1, ref range);
			nextChar = PageBreaks.AfterBreak(editor.TextDocument, nextChar);    // (a page break's paragraph mark stays with it)
		} finally {
			g.ReleaseHdc(hdc);
		}
		if (preview != null) {
			using (preview) {
				g.InterpolationMode = InterpolationMode.HighQualityBicubic;
				g.DrawImage(preview, new RectangleF(-hardX, -hardY, e.PageBounds.Width, e.PageBounds.Height));
			}
			preview = null;
		}
		// The header sits just above the text area and the footer just below it, in the margins.
		if (bands.ShowHeader)
			DrawBand(g, bands.Header, new RectangleF(margins.Left - hardX, Math.Max(5, margins.Top - BandGap - BandHeight) - hardY, margins.Width, BandHeight));
		if (bands.ShowFooter)
			DrawBand(g, bands.Footer, new RectangleF(margins.Left - hardX, margins.Bottom + BandGap - hardY, margins.Width, BandHeight));
		e.HasMorePages = nextChar > start && nextChar < end;
	}

	Bitmap? preview;
	const int PreviewScale = 2;    // the preview's picture of a page: twice the screen's resolution

	// Print Preview records each page as a metafile at the printer's resolution, and shrinking that to the screen loses
	// RichEdit's thin lines (table borders, 1/2 point): there the text is drawn into a picture of the page instead, laid
	// out for the printer as before (so the pages break in the same places), at twice the screen's resolution.
	int PreviewPage(FORMATRANGE range, Rectangle page, out Bitmap? picture) {
		int width = (int)Math.Ceiling(page.Width * 0.96 * PreviewScale), height = (int)Math.Ceiling(page.Height * 0.96 * PreviewScale);
		picture = new Bitmap(width, height);
		using var g = Graphics.FromImage(picture);
		g.Clear(Color.White);
		IntPtr hdc = g.GetHdc();
		try {
			SetGraphicsMode(hdc, 2);    // GM_ADVANCED, for the scale
			var scale = new XFORM { eM11 = PreviewScale, eM22 = PreviewScale };
			SetWorldTransform(hdc, ref scale);
			// the same areas, relative to the page's corner (the picture's) rather than the printable area's
			int dx = (int)(range.rcPage.Left), dy = (int)(range.rcPage.Top);
			range.hdc = hdc;
			range.rc = new RECT { Left = range.rc.Left - dx, Top = range.rc.Top - dy, Right = range.rc.Right - dx, Bottom = range.rc.Bottom - dy };
			range.rcPage = new RECT { Left = 0, Top = 0, Right = range.rcPage.Right - dx, Bottom = range.rcPage.Bottom - dy };
			return (int)SendMessage(editor.Handle, EM_FORMATRANGE, (IntPtr)1, ref range);
		} finally {
			g.ReleaseHdc(hdc);
		}
	}

	[StructLayout(LayoutKind.Sequential)] struct XFORM { public float eM11, eM12, eM21, eM22, eDx, eDy; }
	[DllImport("gdi32")] static extern int SetGraphicsMode(IntPtr hdc, int mode);
	[DllImport("gdi32")] static extern bool SetWorldTransform(IntPtr hdc, ref XFORM transform);

	/// The number of pages the document prints on now (laid out for the printer, without printing).
	public int PageCount() {
		var settings = Document.DefaultPageSettings;
		using var g = Document.PrinterSettings.CreateMeasurementGraphics(settings);
		float hardX = settings.HardMarginX, hardY = settings.HardMarginY;
		var pageBounds = settings.Bounds;
		var margins = new Rectangle(settings.Margins.Left, settings.Margins.Top,
			pageBounds.Width - settings.Margins.Left - settings.Margins.Right, pageBounds.Height - settings.Margins.Top - settings.Margins.Bottom);
		IntPtr hdc = g.GetHdc();
		try {
			var range = new FORMATRANGE { hdc = hdc, hdcTarget = hdc, rc = Twips(margins, hardX, hardY), rcPage = Twips(pageBounds, hardX, hardY), chrg = new CHARRANGE { cpMax = -1 } };
			return CountPages(range, editor.TextDocument.Range2(0, 0).StoryLength - 1);
		} finally {
			g.ReleaseHdc(hdc);
		}
	}

	// Lays out the whole document without drawing (EM_FORMATRANGE with wParam 0) to count the pages.
	int CountPages(FORMATRANGE range, int end) {
		int count = 0, position = 0;
		while (true) {
			count++;
			range.chrg.cpMin = position;
			int next = PageBreaks.AfterBreak(editor.TextDocument, (int)SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range));
			if (next <= position || next >= end) break;
			position = next;
		}
		SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero);
		return count;
	}

	// Each part gets its own column: the center part the width it needs (at most half the line), the left and right
	// parts what is left on either side. Text that does not fit is shortened in the middle ("C:\Users\...\Report.html").
	void DrawBand(Graphics g, string[] parts, RectangleF band) {
		string[] texts = Settings.Parts(parts).Select(Fill).ToArray();
		using var font = new Font("Segoe UI", 8f);
		const float gap = 12;    // 1/100 inch between columns
		float centerWidth = texts[1] == "" ? 0 : Math.Min(g.MeasureString(texts[1], font).Width + 4, band.Width / 2);
		float sideWidth = Math.Max(0, (band.Width - centerWidth) / 2 - (centerWidth > 0 ? gap : gap / 2));
		var columns = new[] {
			new RectangleF(band.Left, band.Top, sideWidth, band.Height),
			new RectangleF(band.Left + (band.Width - centerWidth) / 2, band.Top, centerWidth, band.Height),
			new RectangleF(band.Right - sideWidth, band.Top, sideWidth, band.Height)
		};
		var alignments = new[] { StringAlignment.Near, StringAlignment.Center, StringAlignment.Far };
		for (int i = 0; i < 3; i++) {
			if (texts[i] == "" || columns[i].Width <= 0) continue;
			using var format = new StringFormat { Alignment = alignments[i], LineAlignment = StringAlignment.Center,
				Trimming = StringTrimming.EllipsisPath, FormatFlags = StringFormatFlags.NoWrap };
			g.DrawString(texts[i], font, Brushes.DimGray, columns[i], format);
		}
	}

	string Fill(string text) {
		if (string.IsNullOrEmpty(text)) return "";
		var now = DateTime.Now;
		return Regex.Replace(text, @"\{(title|file|page|pages|date|time)\}", m => m.Groups[1].Value.ToLowerInvariant() switch {
			"title" => title,
			"file" => path != "" ? path : title,
			"page" => page.ToString(CultureInfo.CurrentCulture),
			"pages" => pageCount.ToString(CultureInfo.CurrentCulture),
			"date" => now.ToShortDateString(),
			_ => now.ToShortTimeString()
		}, RegexOptions.IgnoreCase);
	}

	// 1/100 inch (relative to the page) to twips (relative to the printable area, where the printer's HDC starts)
	static RECT Twips(Rectangle r, float hardX, float hardY) => new() {
		Left = (int)((r.Left - hardX) * 14.4f), Top = (int)((r.Top - hardY) * 14.4f),
		Right = (int)((r.Right - hardX) * 14.4f), Bottom = (int)((r.Bottom - hardY) * 14.4f)
	};

	public void Dispose() => Document.Dispose();
}

// Records editor actions as RteClient2 calls (one per line). Typed text is collected into InsertText calls, and
// repeated keys into Key(key, count).
sealed class MacroRecorder
{
	readonly List<string> lines = new();
	readonly StringBuilder typed = new();
	string? lastKey;
	int lastKeyCount;

	public bool Recording { get; private set; }
	public int Steps => lines.Count + (typed.Length > 0 ? 1 : 0);

	public void Start() {
		lines.Clear();
		typed.Clear();
		lastKey = null;
		Recording = true;
	}

	public void Type(char c) {
		if (!Recording) return;
		typed.Append(c);
		lastKey = null;
	}

	// Takes back the last typed character; false if there is none.
	public bool Backspace() {
		if (typed.Length == 0) return false;
		typed.Length--;
		return true;
	}

	// A line of AHK code that is not an Rte call (e.g. running another macro).
	public void AddLine(string line) {
		if (!Recording) return;
		Flush();
		lastKey = null;
		lines.Add(line);
	}

	public void Add(string call) {
		if (!Recording) return;
		Flush();
		lastKey = null;
		lines.Add("Rte." + call);
	}

	public void AddKey(string key) {
		if (!Recording) return;
		Flush();
		if (lastKey == key && lines.Count > 0) {
			lines[^1] = "Rte." + Call("Key", key, ++lastKeyCount);
		} else {
			lastKey = key;
			lastKeyCount = 1;
			lines.Add("Rte." + Call("Key", key));
		}
	}

	void Flush() {
		if (typed.Length == 0) return;
		lines.Add("Rte." + Call("InsertText", typed.ToString()));
		typed.Clear();
		lastKey = null;
	}

	public List<string> Stop() {
		Flush();
		Recording = false;
		return new List<string>(lines);
	}

	public static string Call(string member, params object?[] args) => $"{member}({string.Join(", ", args.Select(Literal))})";

	// An AHK v2 expression for a value.
	public static string Literal(object? value) => value switch {
		null => "\"\"",
		string s => "\"" + s.Replace("`", "``").Replace("\"", "`\"").Replace("\r\n", "`n").Replace("\n", "`n").Replace("\r", "`n").Replace("\t", "`t") + "\"",
		bool b => b ? "1" : "0",
		Enum e => Literal(e.ToString()),
		Color c => Literal($"#{c.R:X2}{c.G:X2}{c.B:X2}"),
		IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
		_ => Literal(value.ToString())
	};
}

// A text box that shows the key combination pressed in it (with Ctrl or Alt, or a function key).
// Backspace or Delete clears it.
sealed class HotkeyBox : TextBox
{
	public Keys Hotkey { get; private set; }
	public event EventHandler? HotkeyChanged;

	public HotkeyBox() {
		ReadOnly = true;
		BackColor = SystemColors.Window;
		ShortcutsEnabled = false;
		Text = "(none)";
	}

	public void SetHotkey(Keys keys) {
		Hotkey = keys;
		Text = keys == Keys.None ? "(none)" : Describe(keys);
		HotkeyChanged?.Invoke(this, EventArgs.Empty);
	}

	public static string Describe(Keys keys) => new KeysConverter().ConvertToInvariantString(keys) ?? "";

	public static Keys? Parse(string text) {
		try {
			return new KeysConverter().ConvertFromInvariantString(text.Trim()) is Keys keys ? keys : null;
		} catch (Exception ex) when (ex is ArgumentException or FormatException or NotSupportedException) {
			return null;
		}
	}

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
		var code = keyData & Keys.KeyCode;
		var modifiers = keyData & Keys.Modifiers;
		if (code is Keys.ShiftKey or Keys.ControlKey or Keys.Menu or Keys.LWin or Keys.RWin) return true;
		if (modifiers == Keys.None && code is Keys.Back or Keys.Delete) { SetHotkey(Keys.None); return true; }
		if (modifiers == Keys.None && code is Keys.Tab or Keys.Enter or Keys.Escape) return base.ProcessCmdKey(ref msg, keyData);
		if ((modifiers & (Keys.Control | Keys.Alt)) != 0 || (code >= Keys.F1 && code <= Keys.F24)) SetHotkey(keyData);
		return true;
	}
}

// Shown when a recording stops: name, optional hotkey, and the script, which can be edited before it is saved.
sealed class SaveMacroDialog : Form
{
	readonly TextBox nameBox = new(), scriptBox = new();
	readonly HotkeyBox hotkeyBox = new();
	readonly Label warning = new() { AutoSize = true, ForeColor = Color.Firebrick };

	public string MacroName => string.Concat(nameBox.Text.Trim().Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
	public string Script => scriptBox.Text;
	public Keys Hotkey => hotkeyBox.Hotkey;

	public SaveMacroDialog(string name, string script, Dictionary<Keys, string> menuShortcuts, WriterForm owner) {
		Text = "Save recorded macro";
		StartPosition = FormStartPosition.CenterParent;
		ShowInTaskbar = false;
		MinimizeBox = false;
		Size = new Size(LogicalToDeviceUnits(720), LogicalToDeviceUnits(520));
		MinimumSize = new Size(LogicalToDeviceUnits(480), LogicalToDeviceUnits(360));
		Padding = new Padding(8);

		nameBox.Text = name;
		nameBox.Width = hotkeyBox.Width = LogicalToDeviceUnits(260);
		scriptBox.Multiline = true;
		scriptBox.ScrollBars = ScrollBars.Both;
		scriptBox.WordWrap = false;
		scriptBox.Font = new Font("Consolas", 10f);
		scriptBox.Dock = DockStyle.Fill;
		scriptBox.Text = script;
		hotkeyBox.HotkeyChanged += (s, e) => {
			warning.Text = menuShortcuts.TryGetValue(hotkeyBox.Hotkey, out var item) ? $"This replaces the menu shortcut of {item} while the editor is active." : "";
		};

		var fields = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Top };
		fields.Controls.Add(new Label { Text = "Name:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
		fields.Controls.Add(nameBox, 1, 0);
		fields.Controls.Add(new Label { Text = "Hotkey (optional):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
		fields.Controls.Add(hotkeyBox, 1, 1);
		fields.Controls.Add(warning, 1, 2);
		fields.Controls.Add(new Label { Text = "Script (you can edit it):", AutoSize = true, Margin = new Padding(3, 8, 3, 3) }, 0, 3);
		fields.SetColumnSpan(fields.GetControlFromPosition(0, 3)!, 2);

		var save = new Button { Text = "Save", AutoSize = true };
		var discard = new Button { Text = "Discard", DialogResult = DialogResult.Cancel, AutoSize = true };
		save.Click += (s, e) => {
			if (MacroName == "") { MessageBox.Show(this, "Give the macro a name.", Text); return; }
			DialogResult = DialogResult.OK;
		};
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom };
		buttons.Controls.AddRange(new Control[] { discard, save });
		Controls.Add(scriptBox);
		Controls.Add(fields);
		Controls.Add(buttons);
		CancelButton = discard;
	}
}

// Tools > Macro hotkeys: one hotkey box per macro.
sealed class MacroHotkeysDialog : Form
{
	readonly Dictionary<string, HotkeyBox> boxes = new();
	readonly Label warning = new() { AutoSize = true, ForeColor = Color.Firebrick, MaximumSize = new Size(600, 0) };
	readonly Dictionary<Keys, string> menuShortcuts;

	public IEnumerable<(string File, Keys Keys)> Hotkeys => boxes.Select(p => (p.Key, p.Value.Hotkey));

	public MacroHotkeysDialog(List<string> files, Func<string, Keys> current, Dictionary<Keys, string> menuShortcuts) {
		this.menuShortcuts = menuShortcuts;
		Text = "Macro hotkeys";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = MaximizeBox = ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Padding = new Padding(10);

		var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
		layout.Controls.Add(new Label { AutoSize = true, Text = "Click a box and press a key combination with Ctrl or Alt, or a function key. Backspace clears it.\nThe hotkeys work while the editor window is active." });
		var scroller = new Panel { AutoScroll = true, Width = LogicalToDeviceUnits(560), Height = LogicalToDeviceUnits(Math.Min(360, 34 * files.Count + 10)) };
		var table = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, Dock = DockStyle.Top };
		int row = 0;
		foreach (var file in files) {
			var box = new HotkeyBox { Width = LogicalToDeviceUnits(170) };
			box.SetHotkey(current(file));
			box.HotkeyChanged += (s, e) => Check();
			var clear = new Button { Text = "Clear", AutoSize = true };
			clear.Click += (s, e) => box.SetHotkey(Keys.None);
			table.Controls.Add(new Label { Text = Path.GetFileNameWithoutExtension(file), AutoSize = true, Anchor = AnchorStyles.Left, MaximumSize = new Size(LogicalToDeviceUnits(280), 0) }, 0, row);
			table.Controls.Add(box, 1, row);
			table.Controls.Add(clear, 2, row);
			boxes[file] = box;
			row++;
		}
		scroller.Controls.Add(table);
		layout.Controls.Add(scroller);
		layout.Controls.Add(warning);
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
		buttons.Controls.AddRange(new Control[] { cancel, ok });
		layout.Controls.Add(buttons);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		Check();
	}

	void Check() {
		var messages = new List<string>();
		foreach (var group in boxes.Where(p => p.Value.Hotkey != Keys.None).GroupBy(p => p.Value.Hotkey)) {
			string keys = HotkeyBox.Describe(group.Key);
			if (group.Count() > 1) messages.Add($"{keys} is used by more than one macro; only one of them will run.");
			if (menuShortcuts.TryGetValue(group.Key, out var item)) messages.Add($"{keys} replaces the menu shortcut of {item}.");
		}
		warning.Text = string.Join("\n", messages);
	}
}

// The editor control: RichEditBox with a hook for the right-click menu, so the spelling menu can come before the
// library's formatting toolbar. The hook gets the screen point and the character position, and returns true if it
// showed a menu.
sealed class WriterEditBox : RichEditBox
{
	public Func<Point, int, bool>? ContextMenuHook;

	/// Ctrl+click on a link (text or picture): the link's address, to be opened.
	public Action<string>? OpenLink;

	protected override void OnMouseUp(MouseEventArgs e) {
		base.OnMouseUp(e);
		if (e.Button != MouseButtons.Left || ModifierKeys != Keys.Control || OpenLink == null) return;
		int position = PositionAt(PointToScreen(e.Location));
		// the character clicked can be the one after the point (a click on its right half): both are looked at
		foreach (int at in new[] { position, position - 1 }) {
			if (at < 0) continue;
			string url = LinkAddress(TextDocument.Range2(at, at + 1).URL);
			if (url != "") { OpenLink(url); return; }
		}
	}

	// The address in a link's field instruction: "https://..." (TOM gives it in quotes, Word's with switches after it).
	public static string LinkAddress(string? url) {
		if (string.IsNullOrWhiteSpace(url)) return "";
		var quoted = System.Text.RegularExpressions.Regex.Match(url, "\"([^\"]+)\"");
		return (quoted.Success ? quoted.Groups[1].Value : url).Trim();
	}

	/// The alt text of the picture at a position, for the tooltip (set by the form).
	public Func<int, string>? AltTextOf;

	readonly ToolTip pictureTip = new() { InitialDelay = 400, ReshowDelay = 100 };
	int tipPicture = -1;

	/// The position of the picture under a screen point, or -1.
	public int PictureAt(Point screen) {
		int position = PositionAt(screen);
		foreach (int at in new[] { position, position - 1 }) {
			if (at < 0) continue;
			try {
				var range = TextDocument.Range2(at, at + 1);
				if (range.Text != "\uFFFC") continue;
				range.GetPoint(0x20, out int left, out int top);       // tomStart, top left (screen)
				range.GetPoint(0x0A, out int right, out int bottom);   // tomEnd, bottom right
				if (screen.X >= left && screen.X <= right && screen.Y >= top && screen.Y <= bottom) return at;
			} catch (Exception ex) when (ex is COMException or ArgumentException) { }
		}
		return -1;
	}

	protected override void OnMouseMove(MouseEventArgs e) {
		base.OnMouseMove(e);
		int picture = e.Button == MouseButtons.None && AltTextOf != null ? PictureAt(PointToScreen(e.Location)) : -1;
		if (picture == tipPicture) return;
		tipPicture = picture;
		if (picture < 0) { pictureTip.Hide(this); return; }
		string alt = "";
		try { alt = AltTextOf!(picture); } catch (Exception ex) when (ex is ArgumentException or COMException) { }
		pictureTip.Show((alt != "" ? "Alt text: " + alt : "No alt text") + "\n(right-click > Alt text... to change it)", this, e.X + 12, e.Y + 20, 8000);
	}

	protected override void OnMouseLeave(EventArgs e) {
		base.OnMouseLeave(e);
		tipPicture = -1;
		pictureTip.Hide(this);
	}

	int PositionAt(Point screen) {
		try { return TextDocument.RangeFromPoint2(screen.X, screen.Y, 0).Start; }
		catch (COMException) { return GetCharIndexFromPosition(PointToClient(screen)); }
	}

	// WinForms hands RichEdit its Font in pixels, so 11 pt became 11.25 pt (and 11.5 in RTF, which keeps half points):
	// new text gets the Font's own size, in points.
	protected override void OnHandleCreated(EventArgs e) {
		base.OnHandleCreated(e);
		SetDefaultSize();
	}

	protected override void OnFontChanged(EventArgs e) {
		base.OnFontChanged(e);
		if (IsHandleCreated) SetDefaultSize();
	}

	public void SetDefaultSize() {
		var format = new CharFormat2 { cbSize = Marshal.SizeOf<CharFormat2>(), dwMask = 0x80000000u, yHeight = (int)Math.Round(Font.SizeInPoints * 20) };    // CFM_SIZE, in twips
		SendCharFormat(Handle, 0x0444, IntPtr.Zero, ref format);    // EM_SETCHARFORMAT, SCF_DEFAULT
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	struct CharFormat2
	{
		public int cbSize;
		public uint dwMask, dwEffects;
		public int yHeight, yOffset, crTextColor;
		public byte bCharSet, bPitchAndFamily;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string szFaceName;
		public short wWeight, sSpacing;
		public int crBackColor, lcid, dwReserved;
		public short sStyle, wKerning;
		public byte bUnderlineType, bAnimation, bRevAuthor, bUnderlineColor;
	}

	[DllImport("user32.dll", EntryPoint = "SendMessageW")]
	static extern IntPtr SendCharFormat(IntPtr hWnd, int msg, IntPtr wParam, ref CharFormat2 lParam);

	// ---- page breaks: RichEdit shows them as empty lines, so a dashed line with "Page break" is drawn over them

	List<int>? pageBreaks;    // their positions; null: to be found again

	protected override void OnTextChanged(EventArgs e) {
		base.OnTextChanged(e);
		pageBreaks = null;
	}

	/// Draws the page breaks in view (after RichEdit has painted).
	public void DrawPageBreaks(Graphics g) {
		pageBreaks ??= PageBreaks.Positions(TextDocument);
		if (pageBreaks.Count == 0) return;
		var client = ClientRectangle;
		int storyEnd = TextDocument.Range2(0, 0).StoryLength;
		int lineHeight = Math.Max(8, Font.Height);
		using var pen = new Pen(Color.FromArgb(150, 150, 150)) { DashStyle = DashStyle.Dash };
		using var label = new Font("Segoe UI", 8f);
		using var brush = new SolidBrush(Color.FromArgb(120, 120, 120));
		using var back = new SolidBrush(BackColor);
		foreach (int at in pageBreaks) {
			int top = GetPositionFromCharIndex(at).Y;
			if (top > client.Bottom) break;
			// the break's line and its paragraph mark's (an empty line) make one band, with the dashed line in its middle
			int after = Math.Min(storyEnd - 1, TextDocument.Range2(at + 1, at + 2).Text == "\r" ? at + 2 : at + 1);
			int bottom = after > at ? GetPositionFromCharIndex(after).Y : top + lineHeight;
			if (after >= storyEnd - 1 || bottom <= top) bottom = top + 2 * lineHeight;
			if (bottom < client.Top) continue;
			int y = (top + bottom) / 2;
			g.DrawLine(pen, client.Left + 4, y, client.Right - 4, y);
			const string text = "Page break";
			var size = g.MeasureString(text, label);
			float x = client.Left + (client.Width - size.Width) / 2;
			g.FillRectangle(back, x - 4, y - size.Height / 2, size.Width + 8, size.Height);
			g.DrawString(text, label, brush, x, y - size.Height / 2);
		}
	}

	protected override void WndProc(ref Message m) {
		const int WM_CONTEXTMENU = 0x007B, WM_PAINT = 0x000F;
		if (m.Msg == WM_PAINT) {
			base.WndProc(ref m);
			if ((pageBreaks == null || pageBreaks.Count > 0) && IsHandleCreated) {
				using var g = CreateGraphics();
				DrawPageBreaks(g);
			}
			return;
		}
		if (m.Msg == WM_CONTEXTMENU && ContextMenuHook != null) {
			long lp = m.LParam.ToInt64();
			int x = (short)(lp & 0xFFFF), y = (short)((lp >> 16) & 0xFFFF);
			bool keyboard = x == -1 && y == -1;    // Shift+F10 or the menu key
			var point = keyboard ? PointToScreen(GetPositionFromCharIndex(SelectionStart)) : new Point(x, y);
			// the character under the mouse as RichEdit's own hit-testing has it (TOM, screen coordinates)
			int position = keyboard ? TextDocument.Selection2.Start : PositionAt(point);
			if (ContextMenuHook(point, position)) return;
		}
		base.WndProc(ref m);
	}
}

// The Windows spell checker (Windows 8 and later): the same one RichEdit uses for its red underlines.
sealed class Spelling
{
	public const int SuggestAction = 1, ReplaceAction = 2, DeleteAction = 3;    // CORRECTIVE_ACTION
	public record Error(int Start, int Length, string Word, int Action, string Replacement);

	readonly ISpellChecker checker;
	public readonly string Language;

	Spelling(ISpellChecker checker, string language) {
		this.checker = checker;
		Language = language;
	}

	// The checker for the current language (or English); null if Windows has no spell checker.
	public static Spelling? Create() {
		try {
			var type = Type.GetTypeFromCLSID(new Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC"));
			if (type == null || Activator.CreateInstance(type) is not ISpellCheckerFactory factory) return null;
			foreach (var language in new[] { CultureInfo.CurrentCulture.Name, CultureInfo.CurrentUICulture.Name, "en-US" })
				if (language != "" && factory.IsSupported(language) != 0) return new Spelling(factory.CreateSpellChecker(language), language);
		} catch (Exception ex) when (ex is COMException or InvalidCastException) { }
		return null;
	}

	// The errors in a paragraph's text; offset is the paragraph's start in the document.
	public List<Error> Check(string text, int offset) {
		var errors = new List<Error>();
		if (string.IsNullOrWhiteSpace(text)) return errors;
		var found = checker.Check(text);
		while (found.Next(out var error) == 0) {
			int start = (int)error.StartIndex, length = (int)error.Length;
			if (start + length > text.Length) continue;
			errors.Add(new Error(offset + start, length, text.Substring(start, length), error.CorrectiveAction, error.Replacement ?? ""));
		}
		return errors;
	}

	public List<string> Suggest(string word, int max = 8) {
		var suggestions = new List<string>();
		var found = checker.Suggest(word);
		var one = new string[1];
		while (suggestions.Count < max && found.Next(1, one, IntPtr.Zero) == 0) suggestions.Add(one[0]);
		return suggestions;
	}

	public void Add(string word) => checker.Add(word);
}

[ComImport, Guid("8E018A9D-2415-4677-BF08-794EA61F94BB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISpellCheckerFactory
{
	System.Runtime.InteropServices.ComTypes.IEnumString SupportedLanguages { get; }
	int IsSupported([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
	ISpellChecker CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
}

[ComImport, Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISpellChecker
{
	string LanguageTag { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
	IEnumSpellingError Check([MarshalAs(UnmanagedType.LPWStr)] string text);
	System.Runtime.InteropServices.ComTypes.IEnumString Suggest([MarshalAs(UnmanagedType.LPWStr)] string word);
	void Add([MarshalAs(UnmanagedType.LPWStr)] string word);
}

[ComImport, Guid("803E3BD4-2828-4410-8290-418D1D73C762"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IEnumSpellingError
{
	[PreserveSig] int Next(out ISpellingError error);
}

[ComImport, Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISpellingError
{
	uint StartIndex { get; }
	uint Length { get; }
	int CorrectiveAction { get; }
	string Replacement { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
}

// The COM IDropTarget interface (OLE drag and drop).
[ComImport, Guid("00000122-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleDropTarget
{
	[PreserveSig] int OleDragEnter([MarshalAs(UnmanagedType.Interface)] object dataObject, int keyState, POINTL point, ref int effect);
	[PreserveSig] int OleDragOver(int keyState, POINTL point, ref int effect);
	[PreserveSig] int OleDragLeave();
	[PreserveSig] int OleDrop([MarshalAs(UnmanagedType.Interface)] object dataObject, int keyState, POINTL point, ref int effect);
}

[StructLayout(LayoutKind.Sequential)] public struct POINTL { public int X, Y; }

// One level of the list scheme (Paragraph > List levels): numbering style, punctuation and start number.
public sealed class ListLevel
{
	public string Style { get; set; } = "Decimal";      // a ListStyle name: Decimal, LowerAlpha, UpperAlpha, LowerRoman, UpperRoman, Bullet
	public string Punctuation { get; set; } = ".";      // ".", ")", "()" or ""
	public int Start { get; set; } = 1;

	public static readonly string[] Punctuations = { ".", ")", "()", "" };

	// 1. a. i. 1. a. i. ...
	public static List<ListLevel> Defaults() => Enumerable.Range(0, 9)
		.Select(i => new ListLevel { Style = (i % 3) switch { 0 => "Decimal", 1 => "LowerAlpha", _ => "LowerRoman" } }).ToList();

	// The TOM ListType: the numbering (low word) and the punctuation (high word: 0 is "1)").
	[System.Text.Json.Serialization.JsonIgnore]
	public int ListType {
		get {
			var style = Enum.TryParse<ListStyle>(Style, true, out var s) ? s : ListStyle.Decimal;
			if (style == ListStyle.Bullet) return tomConstants.tomListBullet;
			int punctuation = Punctuation switch { ")" => 0, "()" => tomConstants.tomListParentheses, "" => tomConstants.tomListPlain, _ => tomConstants.tomListPeriod };
			return (int)style | punctuation;
		}
	}

	// How the first number looks, e.g. "a)" or "(i)".
	public string Sample(int number) {
		var style = Enum.TryParse<ListStyle>(Style, true, out var s) ? s : ListStyle.Decimal;
		if (style == ListStyle.Bullet) return "•";
		string n = style switch {
			ListStyle.LowerAlpha => Letters(number).ToLowerInvariant(),
			ListStyle.UpperAlpha => Letters(number),
			ListStyle.LowerRoman => Roman(number).ToLowerInvariant(),
			ListStyle.UpperRoman => Roman(number),
			_ => number.ToString(CultureInfo.CurrentCulture)
		};
		return Punctuation switch { ")" => n + ")", "()" => "(" + n + ")", "" => n, _ => n + "." };
	}

	static string Letters(int n) {
		string result = "";
		for (n = Math.Max(1, n); n > 0; n = (n - 1) / 26) result = (char)('A' + (n - 1) % 26) + result;
		return result;
	}

	static string Roman(int n) {
		if (n <= 0) return "0";
		var sb = new StringBuilder();
		foreach (var (value, digits) in new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") })
			while (n >= value) { sb.Append(digits); n -= value; }
		return sb.ToString();
	}
}

// Paragraph > List levels: for each level 1-9 the numbering style, punctuation and start number, like a word
// processor's multilevel list; with a preview.
sealed class ListLevelsDialog : Form
{
	readonly ComboBox[] styles = new ComboBox[9], punctuations = new ComboBox[9];
	readonly NumericUpDown[] starts = new NumericUpDown[9];
	readonly Label preview = new() { AutoSize = true, Font = new Font("Segoe UI", 9.5f), Padding = new Padding(8, 4, 8, 4), BackColor = SystemColors.Window };
	readonly CheckBox apply = new() { Text = "Apply to the current list", AutoSize = true };

	static readonly (string Name, string Label)[] Styles = {
		("Decimal", "1, 2, 3"), ("LowerAlpha", "a, b, c"), ("UpperAlpha", "A, B, C"), ("LowerRoman", "i, ii, iii"), ("UpperRoman", "I, II, III"), ("Bullet", "• bullet")
	};
	static readonly string[] PunctuationLabels = { "1.", "1)", "(1)", "1" };

	public List<ListLevel> Levels => Enumerable.Range(0, 9).Select(i => new ListLevel {
		Style = Styles[styles[i].SelectedIndex].Name, Punctuation = ListLevel.Punctuations[punctuations[i].SelectedIndex], Start = (int)starts[i].Value
	}).ToList();
	public bool ApplyToCurrentList => apply.Checked;

	readonly ComboBox legalFormat = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	static readonly (LegalLists.Periods Format, string Label)[] LegalFormats = {
		(LegalLists.Periods.Top, "1.   1.1   1.1.1"), (LegalLists.Periods.All, "1.   1.1.   1.1.1."), (LegalLists.Periods.None, "1   1.1   1.1.1") };
	public LegalLists.Periods LegalFormat => LegalFormats[Math.Max(0, legalFormat.SelectedIndex)].Format;

	public ListLevelsDialog(List<ListLevel> levels, bool inList, LegalLists.Periods legal = LegalLists.Periods.Top) {
		Text = "List levels";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = MaximizeBox = ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Padding = new Padding(10);

		var table = new TableLayoutPanel { ColumnCount = 4, AutoSize = true };
		foreach (var heading in new[] { "Level", "Numbering", "Punctuation", "Start at" })
			table.Controls.Add(new Label { Text = heading, AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 3, 3, 6) });
		for (int i = 0; i < 9; i++) {
			var level = i < levels.Count ? levels[i] : ListLevel.Defaults()[i];
			styles[i] = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = LogicalToDeviceUnits(110) };
			styles[i].Items.AddRange(Styles.Select(st => (object)st.Label).ToArray());
			styles[i].SelectedIndex = Math.Max(0, Array.FindIndex(Styles, st => st.Name.Equals(level.Style, StringComparison.OrdinalIgnoreCase)));
			punctuations[i] = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = LogicalToDeviceUnits(70) };
			punctuations[i].Items.AddRange(PunctuationLabels);
			punctuations[i].SelectedIndex = Math.Max(0, Array.IndexOf(ListLevel.Punctuations, level.Punctuation));
			starts[i] = new NumericUpDown { Minimum = 0, Maximum = 999, Value = Math.Clamp(level.Start, 0, 999), Width = LogicalToDeviceUnits(60) };
			int row = i;
			styles[i].SelectedIndexChanged += (s, e) => { punctuations[row].Enabled = starts[row].Enabled = styles[row].SelectedIndex != Styles.Length - 1; UpdatePreview(); };
			punctuations[i].SelectedIndexChanged += (s, e) => UpdatePreview();
			starts[i].ValueChanged += (s, e) => UpdatePreview();
			table.Controls.Add(new Label { Text = (i + 1).ToString(), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 6, 3, 3) });
			table.Controls.Add(styles[i]);
			table.Controls.Add(punctuations[i]);
			table.Controls.Add(starts[i]);
			punctuations[i].Enabled = starts[i].Enabled = styles[i].SelectedIndex != Styles.Length - 1;
		}

		var defaults = new Button { Text = "Defaults", AutoSize = true };
		defaults.Click += (s, e) => {
			var d = ListLevel.Defaults();
			for (int i = 0; i < 9; i++) {
				styles[i].SelectedIndex = Array.FindIndex(Styles, st => st.Name == d[i].Style);
				punctuations[i].SelectedIndex = 0;
				starts[i].Value = 1;
			}
		};
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
		buttons.Controls.AddRange(new Control[] { cancel, ok, defaults });
		apply.Checked = apply.Enabled = inList;
		if (!inList) apply.Text += " (the caret is not in a list)";

		var right = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Margin = new Padding(16, 0, 0, 0) };
		right.Controls.Add(new Label { Text = "Preview", AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 3, 3, 6) });
		right.Controls.Add(preview);
		var top = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
		top.Controls.Add(table);
		top.Controls.Add(right);
		var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
		layout.Controls.Add(top);
		layout.Controls.Add(new Label { AutoSize = true, Margin = new Padding(3, 8, 3, 3), MaximumSize = new Size(LogicalToDeviceUnits(520), 0),
			Text = "New numbered lists use level 1; Tab and Shift+Tab give items the numbering of their new level (with Paragraph > Automatic numbering style per level)." });
		var legalRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 6, 0, 0) };
		legalFormat.Items.AddRange(LegalFormats.Select(f => (object)f.Label).ToArray());
		legalFormat.SelectedIndex = Math.Max(0, Array.FindIndex(LegalFormats, f => f.Format == legal));
		legalFormat.Width = LogicalToDeviceUnits(150);
		legalRow.Controls.Add(new Label { Text = "Legal numbering (Paragraph > Legal numbering):", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 3, 3) });
		legalRow.Controls.Add(legalFormat);
		layout.Controls.Add(legalRow);
		layout.Controls.Add(apply);
		layout.Controls.Add(buttons);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		UpdatePreview();
	}

	void UpdatePreview() {
		var levels = Levels;
		var sb = new StringBuilder();
		for (int i = 0; i < 4; i++)
			for (int n = 0; n < (i < 3 ? 2 : 1); n++)
				sb.Append(new string(' ', i * 6)).Append(levels[i].Sample(levels[i].Start + n)).Append("  Level ").Append(i + 1).Append('\n');
		preview.Text = sb.ToString().TrimEnd();
	}
}

// Table > Insert table: rows, columns and a bold header row.
sealed class InsertTableDialog : Form
{
	readonly NumericUpDown rows, columns;
	readonly CheckBox header = new() { Text = "Header row (bold)", AutoSize = true, Margin = new Padding(3, 8, 3, 3) };

	public int Rows => (int)rows.Value;
	public int Columns => (int)columns.Value;
	public bool Header => header.Checked;

	public InsertTableDialog(int rowCount, int columnCount, bool headerRow) {
		Text = "Insert table";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = MaximizeBox = ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Padding = new Padding(10);
		rows = new NumericUpDown { Minimum = 1, Maximum = 500, Value = Math.Clamp(rowCount, 1, 500), Width = LogicalToDeviceUnits(70) };
		columns = new NumericUpDown { Minimum = 1, Maximum = 63, Value = Math.Clamp(columnCount, 1, 63), Width = LogicalToDeviceUnits(70) };
		header.Checked = headerRow;
		var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true };
		table.Controls.Add(new Label { Text = "&Rows:", AutoSize = true, Anchor = AnchorStyles.Left });
		table.Controls.Add(rows);
		table.Controls.Add(new Label { Text = "&Columns:", AutoSize = true, Anchor = AnchorStyles.Left });
		table.Controls.Add(columns);
		table.Controls.Add(header);
		table.SetColumnSpan(header, 2);
		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
		buttons.Controls.AddRange(new Control[] { cancel, ok });
		table.Controls.Add(buttons);
		table.SetColumnSpan(buttons, 2);
		Controls.Add(table);
		AcceptButton = ok;
		CancelButton = cancel;
	}
}

// File > Header and Footer: show / hide each, and edit its left, center and right parts.
// The header or footer shown above or below the text: its left, center and right parts in grey, with a dotted line
// between it and the text. Clicking it opens File > Header and Footer.
sealed class BandStrip : Control
{
	readonly bool header;
	string[] parts = { "", "", "" };
	readonly ToolTip tip = new();

	public BandStrip(bool header) {
		this.header = header;
		Dock = header ? DockStyle.Top : DockStyle.Bottom;
		Cursor = Cursors.Hand;
		Visible = false;
		SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
		BackColor = SystemColors.Window;
		Height = 10 + Font.Height;
		tip.SetToolTip(this, (header ? "Header" : "Footer") + " (as on the first page). Click to change it. It is printed on every page, and saved in Word and OpenDocument files.");
	}

	protected override void OnFontChanged(EventArgs e) { base.OnFontChanged(e); Height = LogicalToDeviceUnits(10) + Font.Height; }
	protected override void OnCreateControl() { base.OnCreateControl(); Height = LogicalToDeviceUnits(10) + Font.Height; }

	/// Whether it is shown (Visible is also false while the window is not).
	public bool Shown { get; private set; }

	public void Display(bool show, string[] parts) {
		this.parts = parts;
		Shown = Visible = show;
		Invalidate();
	}

	protected override void OnPaint(PaintEventArgs e) {
		var g = e.Graphics;
		var area = ClientRectangle;
		int gap = LogicalToDeviceUnits(4);
		using var pen = new Pen(Color.FromArgb(190, 190, 190)) { DashStyle = DashStyle.Dot };
		int lineY = header ? area.Bottom - 1 : area.Top;
		g.DrawLine(pen, area.Left, lineY, area.Right - LogicalToDeviceUnits(16), lineY);
		var text = new Rectangle(area.Left, header ? area.Top : area.Top + gap, area.Width - LogicalToDeviceUnits(16), area.Height - gap);
		var color = SystemColors.GrayText;
		const TextFormatFlags flags = TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter;
		int third = text.Width / 3;
		TextRenderer.DrawText(g, parts[0], Font, new Rectangle(text.Left, text.Top, third, text.Height), color, flags | TextFormatFlags.Left);
		TextRenderer.DrawText(g, parts[1], Font, new Rectangle(text.Left + third, text.Top, text.Width - 2 * third, text.Height), color, flags | TextFormatFlags.HorizontalCenter);
		TextRenderer.DrawText(g, parts[2], Font, new Rectangle(text.Right - third, text.Top, third, text.Height), color, flags | TextFormatFlags.Right);
	}

	protected override void Dispose(bool disposing) {
		if (disposing) tip.Dispose();
		base.Dispose(disposing);
	}
}

sealed class HeaderFooterDialog : Form
{
	readonly CheckBox showHeader = new() { Text = "Show a header", AutoSize = true };
	readonly CheckBox showFooter = new() { Text = "Show a footer", AutoSize = true };
	readonly CheckBox makeDefault = new() { Text = "Use for new documents too", AutoSize = true, Margin = new Padding(3, 6, 3, 0) };
	readonly TextBox[] header = { new(), new(), new() }, footer = { new(), new(), new() };

	/// Whether the header and footer become the default for new documents.
	public bool MakeDefault => makeDefault.Checked;

	public HeaderFooterDialog(PageBands bands, PageBands defaultBands) {
		defaults_ = defaultBands;
		Text = "Header and Footer";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = MaximizeBox = ShowInTaskbar = false;
		StartPosition = FormStartPosition.CenterParent;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Padding = new Padding(10);

		var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
		layout.Controls.Add(Band("Header (above the text)", showHeader, header));
		layout.Controls.Add(Band("Footer (below the text)", showFooter, footer));
		layout.Controls.Add(new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(560), 0), Margin = new Padding(3, 6, 3, 6),
			Text = "Fields: {title} document name, {file} full path, {page} page number, {pages} number of pages, {date}, {time}.\n"
				+ "Example: Page {page} of {pages}\n"
				+ "They are printed, and saved with the document: Word and OpenDocument files get them as their own header and footer."
		});
		layout.Controls.Add(makeDefault);

		var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
		var defaults = new Button { Text = "Defaults", AutoSize = true };
		defaults.Click += (s, e) => LoadFrom(defaults_);
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
		buttons.Controls.AddRange(new Control[] { cancel, ok, defaults });
		layout.Controls.Add(buttons);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		LoadFrom(bands);
	}

	PageBands defaults_ = new();

	GroupBox Band(string title, CheckBox show, TextBox[] parts) {
		var box = new GroupBox { Text = title, AutoSize = true, Padding = new Padding(8), Dock = DockStyle.Fill };
		var table = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, Dock = DockStyle.Fill };
		table.Controls.Add(show, 0, 0);
		table.SetColumnSpan(show, 3);
		string[] labels = { "Left", "Center", "Right" };
		for (int i = 0; i < 3; i++) {
			table.Controls.Add(new Label { Text = labels[i], AutoSize = true }, i, 1);
			parts[i].Width = LogicalToDeviceUnits(180);
			table.Controls.Add(parts[i], i, 2);
		}
		show.CheckedChanged += (s, e) => { foreach (var part in parts) part.Enabled = show.Checked; };
		box.Controls.Add(table);
		return box;
	}

	void LoadFrom(PageBands bands) {
		showHeader.Checked = bands.ShowHeader;
		showFooter.Checked = bands.ShowFooter;
		var h = PageBands.Parts(bands.Header);
		var f = PageBands.Parts(bands.Footer);
		for (int i = 0; i < 3; i++) {
			header[i].Text = h[i];
			footer[i].Text = f[i];
			header[i].Enabled = showHeader.Checked;
			footer[i].Enabled = showFooter.Checked;
		}
	}

	public PageBands Result() => new() {
		ShowHeader = showHeader.Checked, Header = header.Select(b => b.Text).ToArray(),
		ShowFooter = showFooter.Checked, Footer = footer.Select(b => b.Text).ToArray()
	};
}

// OLE drop target placed in front of RichEdit's own: dropped files that pass `accepts` go to `drop` (with the screen
// point); everything else (dragged text, other files) goes on to RichEdit's drop target unchanged.
sealed class FileDropTarget : IOleDropTarget
{
	const int DROPEFFECT_NONE = 0, DROPEFFECT_COPY = 1;
	readonly IOleDropTarget? inner;
	readonly Func<string[], bool> accepts;
	readonly Action<string[], Point?> drop;
	string[]? files;    // the files being dragged, while we handle the drag

	FileDropTarget(IOleDropTarget? inner, Func<string[], bool> accepts, Action<string[], Point?> drop) {
		this.inner = inner;
		this.accepts = accepts;
		this.drop = drop;
	}

	[DllImport("ole32")] static extern int OleInitialize(IntPtr reserved);
	[DllImport("ole32")] static extern int RegisterDragDrop(IntPtr hwnd, IOleDropTarget target);
	[DllImport("ole32")] static extern int RevokeDragDrop(IntPtr hwnd);
	[DllImport("user32", CharSet = CharSet.Unicode)] static extern IntPtr GetProp(IntPtr hwnd, string name);

	public static FileDropTarget? Register(IntPtr window, Func<string[], bool> accepts, Action<string[], Point?> drop, out string status) {
		OleInitialize(IntPtr.Zero);
		// OLE keeps the registered drop target of a window in this property.
		IntPtr existing = GetProp(window, "OleDropTargetInterface");
		var inner = existing != IntPtr.Zero ? Marshal.GetObjectForIUnknown(existing) as IOleDropTarget : null;
		RevokeDragDrop(window);
		var target = new FileDropTarget(inner, accepts, drop);
		int hr = RegisterDragDrop(window, target);
		status = $"RegisterDragDrop 0x{hr:X8}, RichEdit's drop target {(inner != null ? "wrapped" : "not found")}";
		return hr == 0 ? target : null;
	}

	public static string[]? GetFiles(object dataObject) {
		try {
			return new DataObject(dataObject).GetData(DataFormats.FileDrop) as string[];
		} catch (Exception ex) when (ex is COMException or ExternalException or InvalidCastException) {
			return null;
		}
	}

	public int OleDragEnter(object dataObject, int keyState, POINTL point, ref int effect) {
		files = GetFiles(dataObject) is { } dragged && accepts(dragged) ? dragged : null;
		if (files != null) {
			effect = DROPEFFECT_COPY;
			return 0;
		}
		if (inner != null) return inner.OleDragEnter(dataObject, keyState, point, ref effect);
		effect = DROPEFFECT_NONE;
		return 0;
	}

	public int OleDragOver(int keyState, POINTL point, ref int effect) {
		if (files != null) {
			effect = DROPEFFECT_COPY;
			return 0;
		}
		if (inner != null) return inner.OleDragOver(keyState, point, ref effect);
		effect = DROPEFFECT_NONE;
		return 0;
	}

	public int OleDragLeave() {
		if (files != null) {
			files = null;
			return 0;
		}
		return inner?.OleDragLeave() ?? 0;
	}

	public int OleDrop(object dataObject, int keyState, POINTL point, ref int effect) {
		if (files != null) {
			var dropped = files;
			files = null;
			effect = DROPEFFECT_COPY;
			drop(dropped, new Point(point.X, point.Y));
			return 0;
		}
		if (inner != null) return inner.OleDrop(dataObject, keyState, point, ref effect);
		effect = DROPEFFECT_NONE;
		return 0;
	}
}

// Hidden top-level window that receives automation requests from other processes as WM_COPYDATA and sends the reply
// back to the sender's window (wParam) as WM_COPYDATA, before returning. Runs on the UI thread.
sealed class AutomationServer : NativeWindow, IDisposable
{
	public const int RequestId = 0x52544531, ReplyId = 0x52544532;    // "RTE1", "RTE2"
	const int WM_COPYDATA = 0x004A;
	readonly Func<string[], string> dispatch;

	[StructLayout(LayoutKind.Sequential)]
	struct COPYDATASTRUCT { public IntPtr dwData; public int cbData; public IntPtr lpData; }

	[DllImport("user32", CharSet = CharSet.Unicode)]
	static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref COPYDATASTRUCT lParam);

	public AutomationServer(Func<string[], string> dispatch) {
		this.dispatch = dispatch;
		CreateHandle(new CreateParams { Caption = "RTE Writer automation server " + Environment.ProcessId });
	}

	protected override void WndProc(ref Message m) {
		if (m.Msg == WM_COPYDATA) {
			var data = Marshal.PtrToStructure<COPYDATASTRUCT>(m.LParam);
			if (data.dwData == (IntPtr)RequestId) {
				string request = Marshal.PtrToStringUni(data.lpData, data.cbData / 2).TrimEnd('\0');
				Reply(m.WParam, dispatch(request.Split('\u001F')));
				m.Result = (IntPtr)1;
				return;
			}
		}
		base.WndProc(ref m);
	}

	void Reply(IntPtr window, string text) {
		IntPtr buffer = Marshal.StringToHGlobalUni(text);
		try {
			var data = new COPYDATASTRUCT { dwData = (IntPtr)ReplyId, cbData = (text.Length + 1) * 2, lpData = buffer };
			SendMessage(window, WM_COPYDATA, Handle, ref data);
		} finally {
			Marshal.FreeHGlobal(buffer);
		}
	}

	public void Dispose() => DestroyHandle();
}

// Modeless Find / Replace window; it is hidden rather than closed so it keeps its settings.
sealed class FindReplaceForm : Form
{
	readonly WriterForm owner;
	readonly TextBox findBox = new(), replaceBox = new();
	readonly CheckBox matchCase = new() { Text = "Match case", AutoSize = true }, wholeWord = new() { Text = "Whole words", AutoSize = true };
	readonly Label replaceLabel = new() { Text = "Replace with:", AutoSize = true, Anchor = AnchorStyles.Left };
	readonly Button replaceButton = new() { Text = "Replace", AutoSize = true }, replaceAllButton = new() { Text = "Replace all", AutoSize = true };

	public string FindText => findBox.Text;
	public string ReplaceText => replaceBox.Text;
	public bool MatchCase => matchCase.Checked;
	public bool WholeWord => wholeWord.Checked;

	public FindReplaceForm(WriterForm owner) {
		this.owner = owner;
		Text = "Find and Replace";
		FormBorderStyle = FormBorderStyle.FixedToolWindow;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.Manual;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Padding = new Padding(8);
		findBox.Width = replaceBox.Width = LogicalToDeviceUnits(260);

		var findButton = new Button { Text = "Find next", AutoSize = true };
		var closeButton = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel };
		findButton.Click += (s, e) => owner.FindNext();
		replaceButton.Click += (s, e) => owner.ReplaceNext();
		replaceAllButton.Click += (s, e) => owner.ReplaceAll();
		closeButton.Click += (s, e) => Hide();

		var table = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, Dock = DockStyle.Fill };
		table.Controls.Add(new Label { Text = "Find what:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
		table.Controls.Add(findBox, 1, 0);
		table.Controls.Add(findButton, 2, 0);
		table.Controls.Add(replaceLabel, 0, 1);
		table.Controls.Add(replaceBox, 1, 1);
		table.Controls.Add(replaceButton, 2, 1);
		var options = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
		options.Controls.AddRange(new Control[] { matchCase, wholeWord });
		table.Controls.Add(options, 1, 2);
		table.Controls.Add(replaceAllButton, 2, 2);
		table.Controls.Add(closeButton, 2, 3);
		Controls.Add(table);
		AcceptButton = findButton;
		CancelButton = closeButton;
		FormClosing += (s, e) => {
			if (e.CloseReason == CloseReason.UserClosing) {
				e.Cancel = true;
				Hide();
			}
		};
	}

	public void ShowFor(bool replace, string? initial) {
		replaceLabel.Visible = replaceBox.Visible = replaceButton.Visible = replaceAllButton.Visible = replace;
		if (initial != null) findBox.Text = initial;
		if (!Visible) {
			var area = owner.RectangleToScreen(owner.ClientRectangle);
			Show(owner);
			Location = new Point(area.Right - Width - LogicalToDeviceUnits(30), area.Top + LogicalToDeviceUnits(70));
		}
		Activate();
		findBox.SelectAll();
		findBox.Focus();
	}
}
