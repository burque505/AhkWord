// Real tables for RichEditBox (used by RTE Writer, rte_writer35.cs, through #load).
// Version 2: table properties. Each row's alignment on the page, height and cell padding, and each cell's width,
// borders (width and color per side), shading and vertical alignment are read from the rows' RTF (RowFormats) and set by
// rewriting the rows (Format); HTML gets them as CSS and reads them back, and Word's cell backgrounds (\clcbpat), which
// RichEdit ignores, are turned into the solid shading it shows (WordShading).
// The built-in .docx converter's tables get their formats in Word's own terms (TablesToDocx, TablesFromDocx).
// RichEdit has tables of its own: in the text, a row is U+FFF9 + paragraph mark, then each cell's text followed by
// U+0007, then U+FFFB + paragraph mark. Tables are created and changed through RTF (\trowd ... \cellx ... \cell \row),
// because RichEditorNET declares TOM's ITextRange2.Row with the wrong type. Rows are changed by rewriting them (a row
// inserted between two rows would end up nested in the next one).
// RichEditorNET's HTML writes each cell as a paragraph between "row" markers and cannot read <table>; ExportHtml and
// LoadHtml do that part.
using System.Drawing;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static class RtfTables
{
	public sealed record Cell(int Start, int End);                        // the cell's text, without its U+0007
	public sealed record Row(int Start, int End, List<Cell> Cells);       // U+FFF9 ... U+FFFB and its paragraph mark
	public sealed record Table(int Start, int End, List<Row> Rows);

	const char RowStart = '\uFFF9', RowEnd = '\uFFFB', CellEnd = '\u0007';
	const string ThinBorders = @"\clbrdrt\brdrw10\brdrs\clbrdrl\brdrw10\brdrs\clbrdrb\brdrw10\brdrs\clbrdrr\brdrw10\brdrs ";

	static string Story(ITextDocument2 doc) => doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";

	/// The document's tables (rows that follow each other make one table).
	public static List<Table> Find(ITextDocument2 doc) {
		string s = Story(doc);
		var tables = new List<Table>();
		var rows = new List<Row>();
		for (int i = 0; i < s.Length; i++) {
			if (s[i] != RowStart) continue;
			int end = s.IndexOf(RowEnd, i);
			if (end < 0) break;
			var cells = new List<Cell>();
			int cell = i + 2;    // after U+FFF9 and its paragraph mark
			for (int j = cell; j < end; j++)
				if (s[j] == CellEnd) { cells.Add(new Cell(cell, j)); cell = j + 1; }
			int rowEnd = Math.Min(s.Length, end + 2);
			if (rows.Count > 0 && rows[^1].End != i) { tables.Add(new Table(rows[0].Start, rows[^1].End, rows)); rows = new List<Row>(); }
			rows.Add(new Row(i, rowEnd, cells));
			i = end;
		}
		if (rows.Count > 0) tables.Add(new Table(rows[0].Start, rows[^1].End, rows));
		return tables;
	}

	/// The table, row and cell at a position, or null outside tables.
	public static (Table Table, int Row, int Cell)? At(RichEditBox editor, int position) {
		foreach (var table in Find(editor.TextDocument)) {
			if (position < table.Start || position >= table.End) continue;
			for (int r = 0; r < table.Rows.Count; r++) {
				var row = table.Rows[r];
				if (position >= row.End) continue;
				int c = row.Cells.FindIndex(cell => position <= cell.End);
				return (table, r, Math.Max(0, c < 0 ? row.Cells.Count - 1 : c));
			}
		}
		return null;
	}

	// ================================================================ creating tables

	/// RTF for a table: cells holds each cell's text (plain), or null for empty cells; widths in twips per column;
	/// formats (if given) each row's format, whose cells' borders, shading and alignment are used (widths too, if set).
	public static string BuildRtf(int rows, IList<int> widths, bool header, Func<int, int, string?>? cells = null, List<RowFormat>? formats = null) {
		var colors = new List<Color?> { null };
		var body = new StringBuilder();
		for (int r = 0; r < rows; r++) {
			var format = formats != null && r < formats.Count ? formats[r] : null;
			body.Append(@"\trowd\trleft0").Append(format != null ? RowWords(format) : @"\trgaph108");
			int x = 0;
			for (int c = 0; c < widths.Count; c++) {
				var cell = format != null && c < format.Cells.Count ? format.Cells[c] : null;
				x += cell != null && cell.Width > 0 && formats!.All(f => f.Cells.Count == widths.Count) ? cell.Width : widths[c];
				body.Append(cell != null ? CellWords(cell, color => ColorIndex(colors, color)) + " " : ThinBorders).Append(@"\cellx").Append(x);
			}
			body.Append(@"\pard\intbl ");
			for (int c = 0; c < widths.Count; c++) {
				string text = Escape(cells?.Invoke(r, c) ?? "");
				// a header cell's mark is bold too, so that text typed into the empty cell is bold
				body.Append(header && r == 0 ? @"\b " + text + @"\cell\b0 " : text + @"\cell ");
			}
			body.Append(@"\row ");
		}
		return @"{\rtf1\ansi " + (colors.Count > 1 ? ColorTableRtf(colors) : "") + body + "}";
	}

	static string Escape(string text) {
		var sb = new StringBuilder();
		foreach (char c in text)
			if (c is '\\' or '{' or '}') sb.Append('\\').Append(c);
			else if (c == '\t') sb.Append(@"\tab ");
			else if (c is '\r' or '\n' or '\v') sb.Append(@"\line ");
			else if (c > 127) sb.Append(@"\u").Append((int)(short)c).Append('?');
			else sb.Append(c);
		return sb.ToString();
	}

	/// Equal column widths that fill widthTwips.
	public static List<int> EqualWidths(int columns, int widthTwips) =>
		Enumerable.Range(0, columns).Select(c => widthTwips * (c + 1) / columns - widthTwips * c / columns).ToList();

	/// Inserts a table at the caret, in a paragraph of its own (the paragraph is split at the caret if needed), and puts
	/// the caret in its first cell. Returns the table's start.
	public static int Insert(RichEditBox editor, int rows, int columns, bool header, int widthTwips, Func<int, int, string?>? cells = null) =>
		InsertRtf(editor, BuildRtf(Math.Max(1, rows), EqualWidths(Math.Max(1, columns), widthTwips), header, cells));

	static int InsertRtf(RichEditBox editor, string rtf) {
		var doc = editor.TextDocument;
		int position = editor.SelectionStart;
		if (At(editor, position) != null) throw new InvalidOperationException("Tables cannot be put inside tables here; put the caret after the table.");
		var para = doc.Range2(position, position);
		para.Expand(tomConstants.tomParagraph);
		if (position > para.Start) {    // split the paragraph: the table goes before the rest of it
			doc.Range2(position, position).Text = "\r";
			position++;
			para = doc.Range2(position, position);
			para.Expand(tomConstants.tomParagraph);
		}
		bool emptyParagraph = (para.Text ?? "").Trim('\r') == "";
		bool lastParagraph = para.End >= doc.Range2(0, 0).StoryLength;
		editor.Select(position, 0);
		((RichTextBox)editor).SelectedRtf = rtf;
		var table = Find(doc).FirstOrDefault(t => t.Start >= position) ?? throw new InvalidOperationException("The table could not be inserted.");
		// an empty paragraph the table was put into is not needed after it (unless it is the document's last one)
		if (emptyParagraph && !lastParagraph) {
			var after = doc.Range2(table.End, table.End);
			after.Expand(tomConstants.tomParagraph);
			if ((after.Text ?? "") == "\r") doc.Range2(after.Start, after.End).Text = "";
		}
		table = Find(doc).First(t => t.Start >= position);
		editor.Select(table.Rows[0].Cells[0].Start, 0);
		return table.Start;
	}

	/// Puts formatted content (an RTF document) into a cell, replacing what is there.
	public static void SetCellRtf(RichEditBox editor, int tableStart, int row, int column, string rtf) {
		var table = Find(editor.TextDocument).First(t => t.Start == tableStart);
		var cell = table.Rows[row].Cells[column];
		editor.Select(cell.Start, cell.End - cell.Start);
		// the document's last paragraph mark would add an empty paragraph to the cell
		((RichTextBox)editor).SelectedRtf = Regex.Replace(rtf, @"\\par\s*\}\s*$", "}");
	}

	// ================================================================ rows and columns

	// A row's RTF, in parts: everything before the cell definitions, the definitions (each ending with \cellxN), the cell
	// contents (each before its \cell), and what follows (\row ...).
	sealed class RowRtf
	{
		public string Head = "", Tail = "";
		public List<string> Definitions = new(), Contents = new();

		public static RowRtf Parse(string rtf) {
			var row = new RowRtf();
			int trowd = rtf.IndexOf(@"\trowd", StringComparison.Ordinal);
			int rowEnd = rtf.LastIndexOf(@"\row", StringComparison.Ordinal);
			var definitions = Regex.Matches(rtf[trowd..rowEnd], @"\\cellx-?\d+").Cast<Match>().ToList();
			int defsStart = trowd + Math.Max(0, Regex.Match(rtf[trowd..rowEnd], @"\\cl[a-z]|\\cellx").Index);
			int defsEnd = trowd + definitions[^1].Index + definitions[^1].Length;
			row.Head = rtf[..defsStart];
			int from = defsStart;
			foreach (var d in definitions) {
				int end = trowd + d.Index + d.Length;
				row.Definitions.Add(rtf[from..end]);
				from = end;
			}
			string contents = rtf[defsEnd..rowEnd];
			if (contents.StartsWith(' ')) contents = contents[1..];    // the space that ends \cellxN
			var cellMarks = Regex.Matches(contents, @"(?<!\\)\\cell(?![a-z])").Cast<Match>().ToList();
			int at = 0;
			foreach (var m in cellMarks) {
				row.Contents.Add(contents[at..m.Index]);
				at = m.Index + m.Length;
				if (at < contents.Length && contents[at] == ' ') at++;    // the space ends the control word; it isn't text
			}
			row.Tail = contents[at..] + rtf[rowEnd..];
			return row;
		}

		public List<int> Widths() {
			var xs = Definitions.Select(d => int.Parse(Regex.Match(d, @"\\cellx(-?\d+)$").Groups[1].Value)).ToList();
			return xs.Select((x, i) => x - (i == 0 ? 0 : xs[i - 1])).ToList();
		}

		public void SetWidths(List<int> widths) {
			int x = 0;
			for (int i = 0; i < Definitions.Count; i++) {
				x += widths[i];
				Definitions[i] = Regex.Replace(Definitions[i], @"\\cellx-?\d+$", @"\cellx" + x);
			}
		}

		// The paragraph properties at the start of the first cell (\pard\intbl ...), which the other cells carry on.
		public string LeadingProperties => Regex.Match(Contents.Count > 0 ? Contents[0] : "", @"^\s*\\pard\\intbl(\\[a-z]+-?\d*\s?)*").Value.Trim();

		public string Build() {
			var sb = new StringBuilder(Head);
			foreach (var d in Definitions) sb.Append(d);
			foreach (var c in Contents) sb.Append(c).Append(@"\cell ");
			return sb.Append(Tail.TrimStart()).ToString();
		}

		// The same row with empty cells.
		public string BuildEmpty() {
			var empty = new RowRtf { Head = Head, Definitions = Definitions, Tail = Tail };
			empty.Contents = Contents.Select((c, i) => i == 0 ? (LeadingProperties == "" ? @"\pard\intbl " : LeadingProperties + " ") : "").ToList();
			return empty.Build();
		}
	}

	static string GetRtf(RichEditBox editor, int start, int end) {
		editor.Select(start, end - start);
		return ((RichTextBox)editor).SelectedRtf ?? "";
	}

	static void SetRtf(RichEditBox editor, int start, int end, string rtf) {
		editor.Select(start, end - start);
		((RichTextBox)editor).SelectedRtf = rtf;
	}

	// The row and cell at the caret, or an error.
	static (Table Table, int Row, int Cell) Current(RichEditBox editor) =>
		At(editor, editor.SelectionStart) ?? throw new InvalidOperationException("Put the caret in a table first.");

	static void CaretTo(RichEditBox editor, int tableStart, int row, int column) {
		var table = Find(editor.TextDocument).FirstOrDefault(t => t.Start == tableStart) ?? Find(editor.TextDocument).FirstOrDefault(t => t.Start >= tableStart);
		if (table == null) return;
		var r = table.Rows[Math.Clamp(row, 0, table.Rows.Count - 1)];
		editor.Select(r.Cells[Math.Clamp(column, 0, r.Cells.Count - 1)].Start, 0);
	}

	/// Inserts an empty row above or below the caret's row.
	public static void InsertRow(RichEditBox editor, bool below) {
		var (table, r, c) = Current(editor);
		var row = table.Rows[r];
		string rtf = GetRtf(editor, row.Start, row.End);
		var parsed = RowRtf.Parse(rtf);
		string empty = parsed.BuildEmpty();
		int trowd = empty.IndexOf(@"\trowd", StringComparison.Ordinal);
		string current = rtf[..(rtf.LastIndexOf(@"\row", StringComparison.Ordinal) + 4)];
		string emptyRow = empty[trowd..(empty.LastIndexOf(@"\row", StringComparison.Ordinal) + 4)];
		string both = below ? current + " " + emptyRow + " }" : rtf[..rtf.IndexOf(@"\trowd", StringComparison.Ordinal)] + emptyRow + " " + current[current.IndexOf(@"\trowd", StringComparison.Ordinal)..] + " }";
		SetRtf(editor, row.Start, row.End, both);
		CaretTo(editor, table.Start, below ? r + 1 : r, c);
	}

	/// Deletes the caret's row (the table, if it is the only row).
	public static void DeleteRow(RichEditBox editor) {
		var (table, r, c) = Current(editor);
		if (table.Rows.Count == 1) { DeleteTable(editor); return; }
		var row = table.Rows[r];
		editor.Select(row.Start, row.End - row.Start);
		editor.SelectedText = "";
		CaretTo(editor, table.Start, Math.Min(r, table.Rows.Count - 2), c);
	}

	/// Inserts an empty column left or right of the caret's column; the table keeps its width.
	public static void InsertColumn(RichEditBox editor, bool right) {
		var (table, r, c) = Current(editor);
		int at = right ? c + 1 : c;
		foreach (var row in Enumerable.Reverse(table.Rows)) {
			var parsed = RowRtf.Parse(GetRtf(editor, row.Start, row.End));
			int index = Math.Min(at, parsed.Definitions.Count);
			var widths = parsed.Widths();
			int total = widths.Sum();
			string lead = parsed.LeadingProperties;
			parsed.Definitions.Insert(index, parsed.Definitions[Math.Max(0, Math.Min(index, parsed.Definitions.Count - 1))]);
			widths.Insert(index, total / Math.Max(1, widths.Count));
			parsed.SetWidths(Scale(widths, total));
			if (index == 0) parsed.Contents.Insert(0, (lead == "" ? @"\pard\intbl" : lead) + " ");
			else parsed.Contents.Insert(index, "");
			SetRtf(editor, row.Start, row.End, parsed.Build());
		}
		CaretTo(editor, table.Start, r, at);
	}

	/// Deletes the caret's column (the table, if it is the only column); the other columns keep their widths.
	public static void DeleteColumn(RichEditBox editor) {
		var (table, r, c) = Current(editor);
		if (table.Rows.All(row => row.Cells.Count <= 1)) { DeleteTable(editor); return; }
		foreach (var row in Enumerable.Reverse(table.Rows)) {
			var parsed = RowRtf.Parse(GetRtf(editor, row.Start, row.End));
			if (c >= parsed.Definitions.Count || parsed.Definitions.Count == 1) continue;
			var widths = parsed.Widths();
			string lead = parsed.LeadingProperties;
			parsed.Definitions.RemoveAt(c);
			widths.RemoveAt(c);
			parsed.SetWidths(widths);
			parsed.Contents.RemoveAt(c);
			if (c == 0 && parsed.Contents.Count > 0 && !parsed.Contents[0].TrimStart().StartsWith(@"\pard")) parsed.Contents[0] = lead + " " + parsed.Contents[0];
			SetRtf(editor, row.Start, row.End, parsed.Build());
		}
		CaretTo(editor, table.Start, r, Math.Max(0, c - 1));
	}

	/// The rows of one table that the selection covers completely, including the last cell's mark (so selecting the text
	/// of a cell is not selecting its row), with their table; none when the selection reaches outside the table.
	public static (Table? Table, List<int> Rows) SelectedRows(RichEditBox editor) {
		var selection = editor.TextDocument.Selection2;    // (WinForms' SelectionLength is short across table rows)
		int start = selection.Start, end = selection.End;
		var table = end > start ? Find(editor.TextDocument).FirstOrDefault(t => start >= t.Start && start < t.End) : null;
		if (table == null || end > table.End) return (null, new List<int>());
		var rows = Enumerable.Range(0, table.Rows.Count)
			.Where(r => start <= table.Rows[r].Cells[0].Start && end > table.Rows[r].Cells[^1].End).ToList();
		return (table, rows);
	}

	/// Deletes the rows the selection covers completely (the table, if that is all of its rows) and returns how many.
	/// RichEdit itself only empties the cells.
	public static int DeleteSelectedRows(RichEditBox editor) {
		var (table, rows) = SelectedRows(editor);
		if (table == null || rows.Count == 0) return 0;
		if (rows.Count == table.Rows.Count) {
			editor.Select(table.Start, table.End - table.Start);
			editor.SelectedText = "";
			return rows.Count;
		}
		foreach (int r in Enumerable.Reverse(rows)) {
			var row = table.Rows[r];
			editor.Select(row.Start, row.End - row.Start);
			editor.SelectedText = "";
		}
		CaretTo(editor, table.Start, Math.Min(rows[0], table.Rows.Count - rows.Count - 1), 0);
		return rows.Count;
	}

	/// Deletes the table the caret is in.
	public static void DeleteTable(RichEditBox editor) {
		var (table, _, _) = Current(editor);
		editor.Select(table.Start, table.End - table.Start);
		editor.SelectedText = "";
	}

	static List<int> Scale(List<int> widths, int total) {
		int sum = Math.Max(1, widths.Sum());
		var scaled = widths.Select(w => Math.Max(200, w * total / sum)).ToList();
		return scaled;
	}

	/// The column widths of each row of the document, in twips, from its RTF (in document order).
	public static List<List<int>> RowWidths(RichEditBox editor) {
		var rows = new List<List<int>>();
		foreach (Match row in Regex.Matches(((RichTextBox)editor).Rtf ?? "", @"\\trowd(.*?)\\row(?![a-z])", RegexOptions.Singleline)) {
			var xs = Regex.Matches(row.Groups[1].Value, @"\\cellx(-?\d+)").Select(m => int.Parse(m.Groups[1].Value)).ToList();
			rows.Add(xs.Select((x, i) => x - (i == 0 ? 0 : xs[i - 1])).ToList());
		}
		return rows;
	}

	// ================================================================ HTML

	/// The library's HTML with its table markers turned into <table> elements. RichEditorNET writes each cell paragraph
	/// as a block between a block holding U+FFF9 and one holding U+FFFB; the document says which cell each belongs to.
	public static string ExportHtml(RichEditBox editor, string html) {
		var tables = Find(editor.TextDocument);
		if (tables.Count == 0 || !html.Contains(RowStart)) return html;
		var formats = RowFormats(((RichTextBox)editor).Rtf ?? "");
		var rows = tables.SelectMany(t => t.Rows.Select((row, i) => (Table: t, Row: row, First: i == 0, Last: i == t.Rows.Count - 1))).ToList();
		if (formats.Count != rows.Count) formats = new List<RowFormat>();
		string story = Story(editor.TextDocument);
		var blocks = Blocks(html);
		var sb = new StringBuilder();
		int b = 0, rowIndex = 0;
		while (b < blocks.Count) {
			string block = blocks[b];
			if (!IsMarker(block, RowStart)) { sb.Append(block); b++; continue; }
			if (rowIndex >= rows.Count) return html;
			var (table, row, first, last) = rows[rowIndex];
			b++;
			var cells = new List<string>();
			foreach (var cell in row.Cells) {
				int paragraphs = story[cell.Start..cell.End].Count(ch => ch == '\r') + 1;
				var content = new StringBuilder();
				while (paragraphs > 0) {
					if (b >= blocks.Count || IsMarker(blocks[b], RowEnd)) return html;    // the HTML does not match the document
					content.Append(blocks[b]);
					paragraphs -= Regex.Matches(blocks[b], @"<li\b", RegexOptions.IgnoreCase).Count is int items and > 0 ? items : 1;
					b++;
				}
				cells.Add(Unwrap(content.ToString()));
			}
			if (b >= blocks.Count || !IsMarker(blocks[b], RowEnd)) return html;
			b++;
			var format = rowIndex < formats.Count ? formats[rowIndex] : null;
			if (first) sb.Append($"<table style=\"{(format != null ? TableCss(format) : "border-collapse: collapse")}\">");
			sb.Append(format != null && RowCss(format) is { Length: > 0 } rowCss ? $"<tr style=\"{rowCss}\">" : "<tr>");
			for (int c = 0; c < cells.Count; c++) {
				string css = format != null && c < format.Cells.Count ? CellCss(format, format.Cells[c]) : "border: 1px solid #000; padding: 2px 6px";
				sb.Append($"<td style=\"{css}\">").Append(cells[c]).Append("</td>");
			}
			sb.Append("</tr>");
			if (last) sb.Append("</table>");
			rowIndex++;
		}
		return sb.ToString();
	}

	static bool IsMarker(string block, char marker) => Regex.Replace(block, "<[^>]+>", "").Trim() == marker.ToString();

	// A cell that is one plain paragraph gets just its content.
	static string Unwrap(string content) {
		var m = Regex.Match(content, @"^<p>(.*)</p>$", RegexOptions.Singleline);
		return m.Success && !m.Groups[1].Value.Contains("<p", StringComparison.OrdinalIgnoreCase) ? m.Groups[1].Value : content;
	}

	// The top-level elements of an HTML fragment (text between them is kept with the next one).
	static List<string> Blocks(string html) {
		var blocks = new List<string>();
		int depth = 0, start = 0;
		foreach (Match tag in Regex.Matches(html, @"<(/?)([a-zA-Z][a-zA-Z0-9]*)[^>]*?(/?)>")) {
			string name = tag.Groups[2].Value.ToLowerInvariant();
			if (name is "br" or "img" or "hr" or "meta" or "input" || tag.Groups[3].Value == "/") {
				if (depth == 0) { blocks.Add(html[start..(tag.Index + tag.Length)]); start = tag.Index + tag.Length; }
				continue;
			}
			if (tag.Groups[1].Value == "/") {
				depth--;
				if (depth == 0) { blocks.Add(html[start..(tag.Index + tag.Length)]); start = tag.Index + tag.Length; }
			} else {
				depth++;
			}
		}
		if (start < html.Length) blocks.Add(html[start..]);
		return blocks;
	}

	/// Loads HTML with load (which loads HTML without tables), then puts real tables where its <table> elements were.
	/// Cells with formatting are filled through a second editor (loadCell loads HTML into it).
	public static void LoadHtml(RichEditBox editor, string html, Action<RichEditBox, string> load, Action<RichEditBox, string> loadCell, int widthTwips) {
		var tables = new List<string>();
		string replaced = ReplaceTables(html, tables);
		load(editor, replaced);
		if (tables.Count == 0) return;
		var doc = editor.TextDocument;
		for (int t = tables.Count - 1; t >= 0; t--) {
			string marker = Marker(t);
			var range = doc.Range2(0, 0);
			if (range.FindText(marker, tomConstants.tomForward, 0) == 0) continue;
			var rows = ParseTable(tables[t], out bool header, out var widths, out var formats);
			if (rows.Count == 0) continue;
			int columns = rows.Max(r => r.Count);
			if (widths == null || widths.Count != columns) widths = EqualWidths(columns, widthTwips);
			foreach (var f in formats) while (f.Cells.Count < columns) f.Cells.Add(new CellFormat { Borders = Enumerable.Repeat(new Border(DefaultBorder, null), 4).ToArray() });
			editor.Select(range.Start, range.End - range.Start);
			editor.SelectedText = "";
			string? Plain(int r, int c) => c < rows[r].Count && IsPlain(rows[r][c].Html) ? WebUtility.HtmlDecode(rows[r][c].Html).Trim() : null;
			int tableStart = InsertRtf(editor, BuildRtf(rows.Count, widths, header, Plain, formats));
			FillCells(editor, tableStart, rows.Select(r => r.Select(c => IsPlain(c.Html) ? null : c.Html).ToList()).ToList(), header, loadCell);
		}
		editor.ClearUndo();
		editor.Select(0, 0);
	}

	/// Fills the cells that have content (null: leave as it is) by loading each into a second editor and copying its RTF.
	public static void FillCells(RichEditBox editor, int tableStart, List<List<string?>> cells, bool header, Action<RichEditBox, string> loadCell) {
		if (cells.All(row => row.All(c => c == null))) return;
		using var scratch = new RichEditBox { Font = editor.Font, EnableHtmlFontSizing = editor.EnableHtmlFontSizing };
		_ = scratch.Handle;
		for (int r = 0; r < cells.Count; r++)
			for (int c = 0; c < cells[r].Count; c++) {
				if (cells[r][c] is not string content) continue;
				loadCell(scratch, content);
				if (header && r == 0) scratch.TextDocument.Range2(0, scratch.TextLength).Font.Bold = tomConstants.tomTrue;
				SetCellRtf(editor, tableStart, r, c, ((RichTextBox)scratch).Rtf ?? "");
			}
	}

	static bool IsPlain(string html) => !html.Contains('<') && !html.Contains('*') && !html.Contains('`') && !html.Contains('[') && !html.Contains('_') && !html.Contains('~');

	static string Marker(int n) => "\uE010" + n + "\uE011";

	static string ReplaceTables(string html, List<string> tables) {
		var sb = new StringBuilder();
		int position = 0;
		foreach (var (start, end) in Elements(html, "table")) {
			sb.Append(html, position, start - position);
			sb.Append("<p>").Append(Marker(tables.Count)).Append("</p>");
			tables.Add(html[start..end]);
			position = end;
		}
		sb.Append(html, position, html.Length - position);
		return sb.ToString();
	}

	// The outermost elements with a tag name: (start, end) of each.
	static IEnumerable<(int Start, int End)> Elements(string html, string name) {
		int depth = 0, start = 0;
		foreach (Match tag in Regex.Matches(html, @"<(/?)" + name + @"\b[^>]*>", RegexOptions.IgnoreCase)) {
			if (tag.Groups[1].Value == "") { if (depth++ == 0) start = tag.Index; }
			else if (depth > 0 && --depth == 0) yield return (start, tag.Index + tag.Length);
		}
	}

	sealed record HtmlCell(string Html);

	static List<List<HtmlCell>> ParseTable(string table, out bool header, out List<int>? widths, out List<RowFormat> formats) {
		var rows = new List<List<HtmlCell>>();
		header = false;
		widths = null;
		formats = new List<RowFormat>();
		// the table's own border: border="0" or a style without one leaves its cells without lines; otherwise thin lines
		string tableAttributes = Regex.Match(table, @"^<table\b([^>]*)>", RegexOptions.IgnoreCase).Groups[1].Value;
		var tableStyle = Style(tableAttributes);
		Border tableBorder = Attribute(tableAttributes, "border") is "0" || (tableStyle.TryGetValue("border", out var tb) && CssBorder(tb) is { Width: 0 })
			? new Border(0, null) : new Border(DefaultBorder, null);
		foreach (var (start, end) in Elements(table, "tr")) {
			string tr = table[start..end];
			string rowAttributes = Regex.Match(tr, @"^<tr\b([^>]*)>", RegexOptions.IgnoreCase).Groups[1].Value;
			var cells = new List<HtmlCell>();
			var cellFormats = new List<CellFormat>();
			var rowWidths = new List<int>();
			int? padding = null;
			foreach (Match cell in Regex.Matches(tr, @"<(td|th)\b([^>]*)>(.*?)</\1>", RegexOptions.IgnoreCase | RegexOptions.Singleline)) {
				cellFormats.Add(CssCell(cell.Groups[2].Value, tableBorder));
				if (padding == null && Style(cell.Groups[2].Value).TryGetValue("padding", out var pad)) {
					var parts = pad.Split(' ', StringSplitOptions.RemoveEmptyEntries);
					padding = CssLength(parts.Length > 1 ? parts[1] : parts[0]);
				}
				if (rows.Count == 0 && cell.Groups[1].Value.Equals("th", StringComparison.OrdinalIgnoreCase)) header = true;
				cells.Add(new HtmlCell(cell.Groups[3].Value.Trim()));
				var w = Regex.Match(cell.Groups[2].Value, @"width\s*[:=]\s*""?([\d.]+)\s*(pt|px)?", RegexOptions.IgnoreCase);
				if (w.Success && float.TryParse(w.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float n))
					rowWidths.Add((int)(w.Groups[2].Value.Equals("pt", StringComparison.OrdinalIgnoreCase) ? n * 20 : n * 15));
			}
			if (rows.Count == 0 && rowWidths.Count == cells.Count && cells.Count > 0) widths = rowWidths;
			if (rows.Count == 0 && Regex.IsMatch(table[..start], @"<thead\b", RegexOptions.IgnoreCase)) header = true;
			if (cells.Count > 0) {
				rows.Add(cells);
				var format = CssRow(tableAttributes, rowAttributes, cellFormats);
				if (padding is int p) format.Padding = p;
				formats.Add(format);
			}
		}
		return rows;
	}

	// ================================================================ table properties (version 2)

	public enum Side { Top, Left, Bottom, Right }

	/// One side of a cell: its line width in twips (0: no border) and color (null: automatic, black).
	public readonly record struct Border(int Width, Color? Color);

	/// A cell's width (twips), borders (by Side), shading (null: none) and vertical alignment (0 top, 1 center, 2 bottom).
	public sealed class CellFormat
	{
		public int Width;
		public Border[] Borders = new Border[4];
		public Color? Shading;
		public int VerticalAlign;
	}

	/// A row's alignment on the page (0 left, 1 center, 2 right), height (twips: 0 automatic, > 0 at least, < 0 exactly),
	/// cell padding (twips, left and right of each cell's text) and cells.
	public sealed class RowFormat
	{
		public int Alignment, Height, Padding = 108;
		public List<CellFormat> Cells = new();
	}

	public const int DefaultBorder = 10;    // twips: a half-point line, as new tables have

	static readonly Regex ControlWord = new(@"\\([a-z]+)(-?\d+)? ?", RegexOptions.Compiled);

	/// The colors of an RTF document's color table (index 0, the automatic color, is null).
	public static List<Color?> ColorTable(string rtf) {
		var colors = new List<Color?>();
		var table = Regex.Match(rtf, @"\{\\colortbl\s*([^}]*)\}");
		if (!table.Success) return new List<Color?> { null };
		foreach (string entry in table.Groups[1].Value.Split(';').SkipLast(1)) {
			int Part(string name) => Regex.Match(entry, @"\\" + name + @"(\d+)") is { Success: true } m ? int.Parse(m.Groups[1].Value) : -1;
			int red = Part("red"), green = Part("green"), blue = Part("blue");
			colors.Add(red < 0 && green < 0 && blue < 0 ? null : Color.FromArgb(Math.Max(0, red), Math.Max(0, green), Math.Max(0, blue)));
		}
		return colors.Count == 0 ? new List<Color?> { null } : colors;
	}

	/// The format of each table row of an RTF document, in document order (from its \trowd ... \cellx definitions).
	public static List<RowFormat> RowFormats(string rtf) {
		var colors = ColorTable(rtf);
		var rows = new List<RowFormat>();
		foreach (Match row in Regex.Matches(rtf, @"\\trowd(.*?)\\row(?![a-z])", RegexOptions.Singleline)) {
			string definitions = row.Groups[1].Value;
			var end = Regex.Match(definitions, @"\\(pard|intbl)(?![a-z])|\\cell(?![a-z])");
			rows.Add(ParseRow(end.Success ? definitions[..end.Index] : definitions, colors));
		}
		return rows;
	}

	// A row from its definitions (what follows \trowd up to the cells' text).
	static RowFormat ParseRow(string definitions, List<Color?> colors) {
		var row = new RowFormat();
		var cell = new CellFormat();
		int side = -1, x = 0, foreColor = -1, backColor = -1, shading = -1;
		bool gapSet = false;
		Color? ColorAt(int index) => index > 0 && index < colors.Count ? colors[index] : null;
		foreach (Match word in ControlWord.Matches(definitions)) {
			string name = word.Groups[1].Value;
			int value = word.Groups[2].Success ? int.Parse(word.Groups[2].Value) : int.MinValue;
			switch (name) {
				case "trql": row.Alignment = 0; break;
				case "trqc": row.Alignment = 1; break;
				case "trqr": row.Alignment = 2; break;
				case "trrh": row.Height = value == int.MinValue ? 0 : value; break;
				case "trgaph": row.Padding = Math.Max(0, value); gapSet = true; break;
				case "trpaddl": if (!gapSet) row.Padding = Math.Max(0, value); break;
				case "clvertalt": cell.VerticalAlign = 0; break;
				case "clvertalc": cell.VerticalAlign = 1; break;
				case "clvertalb": cell.VerticalAlign = 2; break;
				case "clcfpat": foreColor = value; break;
				case "clcbpat": backColor = value; break;
				case "clshdng": shading = value; break;
				case "clbrdrt": side = (int)Side.Top; cell.Borders[side] = new Border(DefaultBorder, null); break;
				case "clbrdrl": side = (int)Side.Left; cell.Borders[side] = new Border(DefaultBorder, null); break;
				case "clbrdrb": side = (int)Side.Bottom; cell.Borders[side] = new Border(DefaultBorder, null); break;
				case "clbrdrr": side = (int)Side.Right; cell.Borders[side] = new Border(DefaultBorder, null); break;
				case "brdrnone" or "brdrnil" or "brdrtbl": if (side >= 0) cell.Borders[side] = new Border(0, null); break;
				case "brdrw": if (side >= 0 && cell.Borders[side].Width > 0) cell.Borders[side] = cell.Borders[side] with { Width = Math.Max(1, value) }; break;
				case "brdrcf": if (side >= 0 && cell.Borders[side].Width > 0) cell.Borders[side] = cell.Borders[side] with { Color = ColorAt(value) }; break;
				case "cellx":
					int edge = value == int.MinValue ? x : value;
					cell.Width = Math.Max(0, edge - x);    // (from the left margin, as RowRtf.Widths measures them)
					cell.Shading = Shade(ColorAt(foreColor), ColorAt(backColor), backColor, shading);
					row.Cells.Add(cell);
					x = edge;
					cell = new CellFormat();
					side = foreColor = backColor = shading = -1;
					break;
			}
		}
		return row;
	}

	// A cell's color from RichEdit's pattern shading (\clcfpat at \clshdng hundredths of a percent over \clcbpat), or Word's
	// plain background (\clcbpat alone).
	static Color? Shade(Color? fore, Color? back, int backIndex, int shading) {
		if (shading <= 0) return backIndex > 0 ? back : null;
		var f = fore ?? Color.Black;
		var b = back ?? Color.White;
		double s = Math.Min(10000, shading) / 10000.0;
		return Color.FromArgb((int)Math.Round(f.R * s + b.R * (1 - s)), (int)Math.Round(f.G * s + b.G * (1 - s)), (int)Math.Round(f.B * s + b.B * (1 - s)));
	}

	/// RTF with Word's cell backgrounds (\clcbpat, without a pattern) as RichEdit shows them: a solid pattern of that color
	/// (RichEdit ignores \clcbpat alone).
	public static string WordShading(string rtf) {
		if (!rtf.Contains(@"\clcbpat", StringComparison.Ordinal)) return rtf;
		return Regex.Replace(rtf, @"(?<before>(?:\\cl[a-z]+-?\d* ?|\\brdr[a-z]+-?\d* ?|\\brsp\d+ ?)*)\\cellx(?<x>-?\d+)", m => {
			string definition = m.Groups["before"].Value;
			if (Regex.IsMatch(definition, @"\\clcfpat\d") || Regex.IsMatch(definition, @"\\clshdng[1-9]")) return m.Value;
			var back = Regex.Match(definition, @"\\clcbpat(\d+) ?");
			if (!back.Success) return m.Value;
			definition = Regex.Replace(definition.Remove(back.Index, back.Length), @"\\clshdng0 ?", "");
			return definition + @"\clcfpat" + back.Groups[1].Value + @"\clshdng10000\cellx" + m.Groups["x"].Value;
		});
	}

	// A row's own words for its format.
	static string RowWords(RowFormat format) =>
		format.Alignment switch { 1 => @"\trqc", 2 => @"\trqr", _ => @"\trql" } + (format.Height != 0 ? @"\trrh" + format.Height : "")
		+ @"\trgaph" + format.Padding + @"\trpaddl" + format.Padding + @"\trpaddr" + format.Padding + @"\trpaddfl3\trpaddfr3";

	// A cell's vertical alignment, shading and borders as RTF (index gives a color's place in the color table).
	static string CellWords(CellFormat cell, Func<Color?, int> index) {
		var sb = new StringBuilder();
		if (cell.VerticalAlign == 1) sb.Append(@"\clvertalc"); else if (cell.VerticalAlign == 2) sb.Append(@"\clvertalb");
		if (cell.Shading is Color shade) sb.Append(@"\clcfpat").Append(index(shade)).Append(@"\clshdng10000");
		string[] names = { "clbrdrt", "clbrdrl", "clbrdrb", "clbrdrr" };
		for (int s = 0; s < 4; s++) {
			var border = cell.Borders[s];
			if (border.Width <= 0) continue;
			sb.Append('\\').Append(names[s]).Append(@"\brdrw").Append(border.Width).Append(@"\brdrs");
			if (border.Color is Color line) sb.Append(@"\brdrcf").Append(index(line));
		}
		return sb.ToString();
	}

	// A color's index in a color table (added to it if it isn't there); 0 for the automatic color.
	static int ColorIndex(List<Color?> colors, Color? color) {
		if (color is not Color c) return 0;
		int found = colors.Count > 1 ? colors.FindIndex(1, k => k is Color known && known.R == c.R && known.G == c.G && known.B == c.B) : -1;
		if (found > 0) return found;
		colors.Add(Color.FromArgb(c.R, c.G, c.B));
		return colors.Count - 1;
	}

	// The color table of a list of colors (index 0, automatic, first).
	static string ColorTableRtf(List<Color?> colors) =>
		@"{\colortbl ;" + string.Concat(colors.Skip(1).Select(c => c is Color k ? $@"\red{k.R}\green{k.G}\blue{k.B};" : ";")) + "}";

	// The definitions of a row (Head up to its last \cellx) rewritten from a format: the row's own words replaced, each
	// cell's borders, shading and vertical alignment replaced, other words (merges, text direction ...) kept.
	static void ApplyRow(RowRtf parsed, RowFormat format) {
		string head = parsed.Head;
		int trowd = head.IndexOf(@"\trowd", StringComparison.Ordinal);
		string rtfHead = head[..trowd], rowWords = head[trowd..];
		rowWords = Regex.Replace(rowWords, @"\\(trq[lcr]|trrh-?\d+|trgaph\d+|trpadd[lr]\d+|trpaddf[lr]\d+) ?", "");
		rowWords = rowWords.TrimEnd() + RowWords(format);
		var colors = ColorTable(rtfHead);
		var definitions = new List<string>();
		int x = 0;
		for (int i = 0; i < parsed.Definitions.Count && i < format.Cells.Count; i++) {
			var cell = format.Cells[i];
			string kept = Regex.Replace(parsed.Definitions[i], @"\\cellx-?\d+$", "");
			kept = Regex.Replace(kept, @"\\clbrdr[tlbr] ?(\\brdr[a-z]+-?\d* ?|\\brsp\d+ ?)*", "");
			kept = Regex.Replace(kept, @"\\(clvertal[tcb]|clcbpat\d+|clcfpat\d+|clshdng\d+) ?", "");
			x += cell.Width;
			definitions.Add(kept.TrimEnd() + CellWords(cell, color => ColorIndex(colors, color)) + @"\cellx" + x);
		}
		for (int i = definitions.Count; i < parsed.Definitions.Count; i++) definitions.Add(parsed.Definitions[i]);
		parsed.Definitions = definitions;
		parsed.Head = WithColorTable(rtfHead, colors) + rowWords;
	}

	// An RTF header with its color table replaced (or added after the font table).
	static string WithColorTable(string rtfHead, List<Color?> colors) {
		string table = ColorTableRtf(colors);
		var existing = Regex.Match(rtfHead, @"\{\\colortbl\s*[^}]*\}");
		if (existing.Success) return rtfHead.Remove(existing.Index, existing.Length).Insert(existing.Index, table);
		var fonts = Regex.Match(rtfHead, @"\{\\fonttbl(?:[^{}]|\{[^{}]*\})*\}");
		int at = fonts.Success ? fonts.Index + fonts.Length : Math.Max(0, rtfHead.IndexOf(' ') + 1);
		return rtfHead.Insert(at, table);
	}

	/// The formats of a table's rows (from the document's RTF).
	public static List<RowFormat> Formats(RichEditBox editor, Table table) {
		var all = RowFormats(((RichTextBox)editor).Rtf ?? "");
		var tables = Find(editor.TextDocument);
		int before = tables.TakeWhile(t => t.Start != table.Start).Sum(t => t.Rows.Count);
		return all.Count >= before + table.Rows.Count ? all.GetRange(before, table.Rows.Count) : table.Rows.Select(r => new RowFormat()).ToList();
	}

	/// Changes the format of the rows of the table starting at tableStart: change gets all the rows' formats (and changes
	/// them); the rows are rewritten (as one undo step when the caller groups the edit). The selection is kept.
	public static void Format(RichEditBox editor, int tableStart, Action<List<RowFormat>> change) {
		var table = Find(editor.TextDocument).First(t => t.Start == tableStart);
		var formats = Formats(editor, table);
		change(formats);
		var selection = editor.TextDocument.Selection2;
		int start = selection.Start, end = selection.End;
		for (int r = table.Rows.Count - 1; r >= 0; r--) {
			var row = Find(editor.TextDocument).First(t => t.Start == tableStart).Rows[r];
			var parsed = RowRtf.Parse(GetRtf(editor, row.Start, row.End));
			ApplyRow(parsed, formats[r]);
			SetRtf(editor, row.Start, row.End, parsed.Build());
		}
		editor.TextDocument.Selection2.SetRange(start, end);
	}

	/// The cells (row, column) of a table that the selection touches, or the caret's cell.
	public static List<(int Row, int Cell)> SelectedCells(RichEditBox editor, Table table) {
		var selection = editor.TextDocument.Selection2;
		int start = selection.Start, end = selection.End;
		var cells = new List<(int, int)>();
		for (int r = 0; r < table.Rows.Count; r++)
			for (int c = 0; c < table.Rows[r].Cells.Count; c++) {
				var cell = table.Rows[r].Cells[c];
				bool touched = end > start ? cell.Start < end && cell.End >= start : start >= cell.Start && start <= cell.End;
				if (touched) cells.Add((r, c));
			}
		return cells;
	}

	/// Sets the borders of a block of cells: preset "none", "all", "outside", "inside", "top", "bottom", "left" or "right".
	/// The sides they share with cells next to them are set too, so a removed line is gone and an added one is whole.
	public static void SetBorders(List<RowFormat> rows, List<(int Row, int Cell)> cells, string preset, Border border) {
		if (cells.Count == 0) return;
		if (border.Color is Color k && k.R == 0 && k.G == 0 && k.B == 0) border = border with { Color = null };    // black is the automatic color
		var chosen = cells.ToHashSet();
		int top = cells.Min(c => c.Row), bottom = cells.Max(c => c.Row);
		void Set(int r, int c, Side side, Border value) {
			if (r < 0 || r >= rows.Count || c < 0 || c >= rows[r].Cells.Count) return;
			rows[r].Cells[c].Borders[(int)side] = value;
			var (nr, nc, other) = side switch { Side.Top => (r - 1, c, Side.Bottom), Side.Bottom => (r + 1, c, Side.Top), Side.Left => (r, c - 1, Side.Right), _ => (r, c + 1, Side.Left) };
			if (nr >= 0 && nr < rows.Count && nc >= 0 && nc < rows[nr].Cells.Count) rows[nr].Cells[nc].Borders[(int)other] = value;
		}
		var none = new Border(0, null);
		string kind = preset.Trim().ToLowerInvariant();
		foreach (var (r, c) in cells) {
			int first = cells.Where(k => k.Row == r).Min(k => k.Cell), last = cells.Where(k => k.Row == r).Max(k => k.Cell);
			bool atTop = r == top || !chosen.Contains((r - 1, c)), atBottom = r == bottom || !chosen.Contains((r + 1, c));
			bool atLeft = c == first, atRight = c == last;
			switch (kind) {
				case "none": foreach (Side s in Enum.GetValues<Side>()) Set(r, c, s, none); break;
				case "all": foreach (Side s in Enum.GetValues<Side>()) Set(r, c, s, border); break;
				case "outside":
					if (atTop) Set(r, c, Side.Top, border);
					if (atBottom) Set(r, c, Side.Bottom, border);
					if (atLeft) Set(r, c, Side.Left, border);
					if (atRight) Set(r, c, Side.Right, border);
					break;
				case "inside":
					if (!atTop) Set(r, c, Side.Top, border);
					if (!atBottom) Set(r, c, Side.Bottom, border);
					if (!atLeft) Set(r, c, Side.Left, border);
					if (!atRight) Set(r, c, Side.Right, border);
					break;
				case "top": if (atTop) Set(r, c, Side.Top, border); break;
				case "bottom": if (atBottom) Set(r, c, Side.Bottom, border); break;
				case "left": if (atLeft) Set(r, c, Side.Left, border); break;
				case "right": if (atRight) Set(r, c, Side.Right, border); break;
				default: throw new ArgumentException("Borders are \"none\", \"all\", \"outside\", \"inside\", \"top\", \"bottom\", \"left\" or \"right\".");
			}
		}
	}

	// ---- HTML for the formats

	static string Hex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

	static string Points(int twips) => (twips / 20.0).ToString("0.##", CultureInfo.InvariantCulture) + "pt";

	static string BorderCss(Border border) => border.Width <= 0 ? "none" : $"{Points(border.Width)} solid {Hex(border.Color ?? Color.Black)}";

	/// The style of a <td> for a cell of a row.
	public static string CellCss(RowFormat row, CellFormat cell) {
		var css = new List<string>();
		string[] sides = { "top", "left", "bottom", "right" };
		var b = cell.Borders;
		if (b.All(x => x == b[0])) css.Add("border: " + BorderCss(b[0]));
		else for (int s = 0; s < 4; s++) css.Add($"border-{sides[s]}: {BorderCss(b[s])}");
		css.Add($"padding: 2px {Points(row.Padding)}");
		if (cell.Width > 0) css.Add("width: " + Points(cell.Width));
		if (cell.Shading is Color shade) css.Add("background-color: " + Hex(shade));
		if (cell.VerticalAlign != 0) css.Add("vertical-align: " + (cell.VerticalAlign == 1 ? "middle" : "bottom"));
		return string.Join("; ", css);
	}

	/// The style of a <table> whose first row is row: centered or right-aligned tables get auto margins.
	public static string TableCss(RowFormat row) =>
		"border-collapse: collapse" + row.Alignment switch { 1 => "; margin-left: auto; margin-right: auto", 2 => "; margin-left: auto; margin-right: 0", _ => "" };

	/// The style of a <tr> (its height), or "".
	public static string RowCss(RowFormat row) => row.Height == 0 ? "" : "height: " + Points(Math.Abs(row.Height));

	// ---- reading HTML styles

	static readonly Dictionary<string, Color> NamedColors = new(StringComparer.OrdinalIgnoreCase) {
		["black"] = Color.Black, ["white"] = Color.White, ["gray"] = Color.Gray, ["grey"] = Color.Gray, ["silver"] = Color.Silver,
		["red"] = Color.Red, ["maroon"] = Color.Maroon, ["green"] = Color.Green, ["lime"] = Color.Lime, ["blue"] = Color.Blue,
		["navy"] = Color.Navy, ["yellow"] = Color.Yellow, ["orange"] = Color.Orange, ["purple"] = Color.Purple, ["teal"] = Color.Teal,
		["aqua"] = Color.Aqua, ["fuchsia"] = Color.Fuchsia, ["olive"] = Color.Olive, ["lightgray"] = Color.LightGray, ["lightgrey"] = Color.LightGray,
	};

	/// A CSS color (#rgb, #rrggbb, rgb(...), a basic name), or null.
	public static Color? CssColor(string value) {
		value = value.Trim();
		if (Regex.Match(value, @"^#([0-9a-f]{6})$", RegexOptions.IgnoreCase) is { Success: true } six)
			return Color.FromArgb(Convert.ToInt32(six.Groups[1].Value[..2], 16), Convert.ToInt32(six.Groups[1].Value[2..4], 16), Convert.ToInt32(six.Groups[1].Value[4..], 16));
		if (Regex.Match(value, @"^#([0-9a-f]{3})$", RegexOptions.IgnoreCase) is { Success: true } three)
			return Color.FromArgb(Convert.ToInt32(new string(three.Groups[1].Value[0], 2), 16), Convert.ToInt32(new string(three.Groups[1].Value[1], 2), 16), Convert.ToInt32(new string(three.Groups[1].Value[2], 2), 16));
		if (Regex.Match(value, @"^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", RegexOptions.IgnoreCase) is { Success: true } rgb)
			return Color.FromArgb(Math.Min(255, int.Parse(rgb.Groups[1].Value)), Math.Min(255, int.Parse(rgb.Groups[2].Value)), Math.Min(255, int.Parse(rgb.Groups[3].Value)));
		return NamedColors.TryGetValue(value, out var named) ? named : null;
	}

	/// A length in twips: CSS's pt, px, in, cm and mm (px when there is no unit), or null.
	public static int? CssLength(string value) {
		var m = Regex.Match(value.Trim(), @"^([\d.]+)\s*(pt|px|in|cm|mm)?$", RegexOptions.IgnoreCase);
		if (!m.Success || !double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double n)) return null;
		return (int)Math.Round(m.Groups[2].Value.ToLowerInvariant() switch { "pt" => n * 20, "in" => n * 1440, "cm" => n * 1440 / 2.54, "mm" => n * 144 / 2.54, _ => n * 15 });
	}

	// A CSS border ("1px solid #000", "none", "0.5pt dotted red"), or null if it says nothing.
	static Border? CssBorder(string value) {
		value = value.Trim();
		if (value == "") return null;
		if (Regex.IsMatch(value, @"\b(none|hidden)\b", RegexOptions.IgnoreCase) || value == "0") return new Border(0, null);
		int width = DefaultBorder;
		Color? color = null;
		foreach (string part in Regex.Split(value, @"\s+(?![^(]*\))")) {
			if (CssLength(part) is int length) width = Math.Max(1, length);
			else if (part.Equals("medium", StringComparison.OrdinalIgnoreCase)) width = 30;
			else if (part.Equals("thick", StringComparison.OrdinalIgnoreCase)) width = 50;
			else if (CssColor(part) is Color c) color = c;
		}
		return new Border(width, color is Color k && k.R == 0 && k.G == 0 && k.B == 0 ? null : color);
	}

	static Dictionary<string, string> Style(string attributes) {
		var style = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		var m = Regex.Match(attributes, @"style\s*=\s*(""([^""]*)""|'([^']*)')", RegexOptions.IgnoreCase);
		if (m.Success)
			foreach (string declaration in (m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value).Split(';')) {
				int colon = declaration.IndexOf(':');
				if (colon > 0) style[declaration[..colon].Trim()] = declaration[(colon + 1)..].Trim();
			}
		return style;
	}

	static string? Attribute(string attributes, string name) =>
		Regex.Match(attributes, @"\b" + name + @"\s*=\s*(""([^""]*)""|'([^']*)'|([^\s>]+))", RegexOptions.IgnoreCase) is { Success: true } m
			? (m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Success ? m.Groups[3].Value : m.Groups[4].Value) : null;

	// A cell's format from its <td>/<th> attributes, over the table's border (tableBorder).
	static CellFormat CssCell(string attributes, Border tableBorder) {
		var style = Style(attributes);
		var cell = new CellFormat();
		var all = style.TryGetValue("border", out var shorthand) ? CssBorder(shorthand) : null;
		string[] sides = { "top", "left", "bottom", "right" };
		for (int s = 0; s < 4; s++)
			cell.Borders[s] = (style.TryGetValue("border-" + sides[s], out var side) ? CssBorder(side) : null) ?? all ?? tableBorder;
		if ((style.TryGetValue("background-color", out var back) || style.TryGetValue("background", out back)) && CssColor(back.Split(' ')[0]) is Color b) cell.Shading = b;
		else if (Attribute(attributes, "bgcolor") is { } bg && CssColor(bg) is Color bgc) cell.Shading = bgc;
		string valign = style.TryGetValue("vertical-align", out var v) ? v : Attribute(attributes, "valign") ?? "";
		cell.VerticalAlign = valign.ToLowerInvariant() switch { "middle" or "center" => 1, "bottom" => 2, _ => 0 };
		if (style.TryGetValue("width", out var w) && CssLength(w) is int width) cell.Width = width;
		else if (Attribute(attributes, "width") is { } aw && CssLength(aw) is int awidth) cell.Width = awidth;
		return cell;
	}

	// The row format of a table (alignment, padding) and row (height) from their attributes.
	static RowFormat CssRow(string tableAttributes, string rowAttributes, List<CellFormat> cells) {
		var table = Style(tableAttributes);
		var row = new RowFormat { Cells = cells };
		string left = table.TryGetValue("margin-left", out var ml) ? ml : "", right = table.TryGetValue("margin-right", out var mr) ? mr : "";
		if (table.TryGetValue("margin", out var margin) && margin.Contains("auto", StringComparison.OrdinalIgnoreCase)) left = right = "auto";
		string align = Attribute(tableAttributes, "align") ?? "";
		row.Alignment = left == "auto" && right == "auto" || align.Equals("center", StringComparison.OrdinalIgnoreCase) ? 1
			: left == "auto" || align.Equals("right", StringComparison.OrdinalIgnoreCase) ? 2 : 0;
		var rowStyle = Style(rowAttributes);
		if (rowStyle.TryGetValue("height", out var h) && CssLength(h) is int height) row.Height = height;
		return row;
	}

	// ================================================================ Word (.docx) tables

	// The built-in .docx converter goes through HTML, whose table styles it only partly understands (it wrote every
	// border black and 3 pt, and left out cell padding); its reader loses the table's alignment and padding. These put
	// the formats in the .docx file itself, in Word's own terms, after it is written, and take them from it after it
	// is read. Tables are matched by order, and only when their rows and cells match.

	static readonly System.Xml.Linq.XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	// (the order Word's schema wants the properties in)
	static readonly string[] TblPrOrder = { "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook" };
	static readonly string[] TrPrOrder = { "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden" };
	static readonly string[] TcPrOrder = { "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark" };

	static System.Xml.Linq.XElement Props(System.Xml.Linq.XElement element, string name) {
		var props = element.Element(W + name);
		if (props == null) { props = new System.Xml.Linq.XElement(W + name); element.AddFirst(props); }
		return props;
	}

	// Puts child in parent (replacing one of its name) where Word's order wants it; null removes it.
	static void SetChild(System.Xml.Linq.XElement parent, string name, System.Xml.Linq.XElement? child, string[] order) {
		parent.Element(W + name)?.Remove();
		if (child == null) return;
		int rank = Array.IndexOf(order, name);
		var before = parent.Elements().LastOrDefault(e => Array.IndexOf(order, e.Name.LocalName) is int r && r >= 0 && r < rank);
		if (before != null) before.AddAfterSelf(child); else parent.AddFirst(child);
	}

	static System.Xml.Linq.XElement Element(string name, params object[] content) => new(W + name, content);
	static System.Xml.Linq.XAttribute Value(string name, object value) => new(W + name, value);

	// The document's top-level tables (not tables in tables) and their rows' cells.
	static List<System.Xml.Linq.XElement> DocxTables(System.Xml.Linq.XDocument document) =>
		document.Descendants(W + "tbl").Where(t => !t.Ancestors(W + "tbl").Any()).ToList();

	static System.Xml.Linq.XElement BorderElement(string side, Border border) => border.Width <= 0
		? Element(side, Value("val", "nil"))
		: Element(side, Value("val", "single"), Value("sz", Math.Max(2, (int)Math.Round(border.Width / 2.5))), Value("space", 0),
			Value("color", border.Color is Color c ? $"{c.R:X2}{c.G:X2}{c.B:X2}" : "auto"));

	/// Writes the formats of the document's tables (one list of row formats per table) into a .docx file's tables.
	public static void TablesToDocx(string path, List<List<RowFormat>> tables) {
		using var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update);
		var entry = zip.GetEntry("word/document.xml");
		if (entry == null) return;
		System.Xml.Linq.XDocument document;
		using (var stream = entry.Open()) document = System.Xml.Linq.XDocument.Load(stream);
		var docxTables = DocxTables(document);
		if (docxTables.Count != tables.Count) return;
		for (int t = 0; t < tables.Count; t++) {
			var rows = docxTables[t].Elements(W + "tr").ToList();
			var formats = tables[t];
			if (rows.Count != formats.Count || formats.Count == 0) continue;
			var first = formats[0];
			var tblPr = Props(docxTables[t], "tblPr");
			SetChild(tblPr, "jc", Element("jc", Value("val", first.Alignment switch { 1 => "center", 2 => "right", _ => "left" })), TblPrOrder);
			SetChild(tblPr, "tblCellMar", Element("tblCellMar",
				Element("left", Value("w", first.Padding), Value("type", "dxa")), Element("right", Value("w", first.Padding), Value("type", "dxa"))), TblPrOrder);
			for (int r = 0; r < rows.Count; r++) {
				var format = formats[r];
				var cells = rows[r].Elements(W + "tc").ToList();
				if (cells.Count != format.Cells.Count) continue;
				var trPr = Props(rows[r], "trPr");
				SetChild(trPr, "trHeight", format.Height == 0 ? null
					: Element("trHeight", Value("val", Math.Abs(format.Height)), Value("hRule", format.Height < 0 ? "exact" : "atLeast")), TrPrOrder);
				for (int c = 0; c < cells.Count; c++) {
					var cell = format.Cells[c];
					var tcPr = Props(cells[c], "tcPr");
					if (cell.Width > 0) SetChild(tcPr, "tcW", Element("tcW", Value("w", cell.Width), Value("type", "dxa")), TcPrOrder);
					SetChild(tcPr, "tcBorders", Element("tcBorders", BorderElement("top", cell.Borders[(int)Side.Top]), BorderElement("left", cell.Borders[(int)Side.Left]),
						BorderElement("bottom", cell.Borders[(int)Side.Bottom]), BorderElement("right", cell.Borders[(int)Side.Right])), TcPrOrder);
					SetChild(tcPr, "shd", cell.Shading is Color shade
						? Element("shd", Value("val", "clear"), Value("color", "auto"), Value("fill", $"{shade.R:X2}{shade.G:X2}{shade.B:X2}")) : null, TcPrOrder);
					SetChild(tcPr, "vAlign", Element("vAlign", Value("val", cell.VerticalAlign switch { 1 => "center", 2 => "bottom", _ => "top" })), TcPrOrder);
				}
			}
			// the column grid, from the first row's widths
			if (docxTables[t].Element(W + "tblGrid") is { } grid && first.Cells.All(c => c.Width > 0) && grid.Elements(W + "gridCol").Count() == first.Cells.Count)
				foreach (var (col, cell) in grid.Elements(W + "gridCol").Zip(first.Cells)) col.SetAttributeValue(W + "w", cell.Width);
		}
		using (var stream = entry.Open()) {
			stream.SetLength(0);
			document.Save(stream);
		}
	}

	static Border? DocxBorder(System.Xml.Linq.XElement? border) {
		if (border == null) return null;
		string kind = (string?)border.Attribute(W + "val") ?? "single";
		if (kind is "nil" or "none") return new Border(0, null);
		int size = int.TryParse((string?)border.Attribute(W + "sz"), out int sz) ? sz : 4;
		string color = (string?)border.Attribute(W + "color") ?? "auto";
		Color? c = color.Length == 6 && !color.Equals("auto", StringComparison.OrdinalIgnoreCase) && CssColor("#" + color) is Color k && (k.R | k.G | k.B) != 0 ? k : null;
		return new Border(Math.Max(1, (int)Math.Round(size * 2.5)), c);
	}

	/// Takes the formats a .docx file's tables have (alignment, padding, row heights, and the borders, shading and
	/// vertical alignment their cells state) onto the document's tables, as read from it.
	public static void TablesFromDocx(RichEditBox editor, string path) {
		System.Xml.Linq.XDocument document;
		using (var zip = System.IO.Compression.ZipFile.OpenRead(path)) {
			var entry = zip.GetEntry("word/document.xml");
			if (entry == null) return;
			using var stream = entry.Open();
			document = System.Xml.Linq.XDocument.Load(stream);
		}
		var docxTables = DocxTables(document);
		var tables = Find(editor.TextDocument);
		if (docxTables.Count != tables.Count) return;
		for (int t = tables.Count - 1; t >= 0; t--) {
			var docxRows = docxTables[t].Elements(W + "tr").ToList();
			if (docxRows.Count != tables[t].Rows.Count) continue;
			var tblPr = docxTables[t].Element(W + "tblPr");
			string? jc = (string?)tblPr?.Element(W + "jc")?.Attribute(W + "val");
			int? padding = int.TryParse((string?)(tblPr?.Element(W + "tblCellMar")?.Element(W + "left") ?? tblPr?.Element(W + "tblCellMar")?.Element(W + "start"))?.Attribute(W + "w"), out int pad) ? pad : null;
			var tableBorders = tblPr?.Element(W + "tblBorders");
			Border? TableSide(string name) => DocxBorder(tableBorders?.Element(W + name));
			Format(editor, tables[t].Start, rows => {
				for (int r = 0; r < rows.Count; r++) {
					var format = rows[r];
					if (jc != null) format.Alignment = jc is "center" ? 1 : jc is "right" or "end" ? 2 : 0;
					if (padding is int p) format.Padding = p;
					var height = docxRows[r].Element(W + "trPr")?.Element(W + "trHeight");
					if (height != null && int.TryParse((string?)height.Attribute(W + "val"), out int h))
						format.Height = (string?)height.Attribute(W + "hRule") == "exact" ? -h : h;
					var cells = docxRows[r].Elements(W + "tc").ToList();
					if (cells.Count != format.Cells.Count) continue;
					for (int c = 0; c < cells.Count; c++) {
						var tcPr = cells[c].Element(W + "tcPr");
						var cell = format.Cells[c];
						var borders = tcPr?.Element(W + "tcBorders");
						Border? Side(string a, string b) => DocxBorder(borders?.Element(W + a) ?? borders?.Element(W + b));
						var top = Side("top", "top") ?? (r == 0 ? TableSide("top") : TableSide("insideH"));
						var left = Side("left", "start") ?? (c == 0 ? TableSide("left") ?? TableSide("start") : TableSide("insideV"));
						var bottom = Side("bottom", "bottom") ?? (r == rows.Count - 1 ? TableSide("bottom") : TableSide("insideH"));
						var right = Side("right", "end") ?? (c == cells.Count - 1 ? TableSide("right") ?? TableSide("end") : TableSide("insideV"));
						if (top is Border bt) cell.Borders[(int)RtfTables.Side.Top] = bt;
						if (left is Border bl) cell.Borders[(int)RtfTables.Side.Left] = bl;
						if (bottom is Border bb) cell.Borders[(int)RtfTables.Side.Bottom] = bb;
						if (right is Border br) cell.Borders[(int)RtfTables.Side.Right] = br;
						string? fill = (string?)tcPr?.Element(W + "shd")?.Attribute(W + "fill");
						if (fill != null) cell.Shading = fill.Length == 6 && !fill.Equals("auto", StringComparison.OrdinalIgnoreCase) && CssColor("#" + fill) is Color f ? f : null;
						string? valign = (string?)tcPr?.Element(W + "vAlign")?.Attribute(W + "val");
						if (valign != null) cell.VerticalAlign = valign switch { "center" => 1, "bottom" => 2, _ => 0 };
					}
				}
			});
		}
	}
}
