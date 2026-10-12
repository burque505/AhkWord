// Real tables for RichEditBox (used by RTE Writer, rte_writer9.cs, through #load).
// RichEdit has tables of its own: in the text, a row is U+FFF9 + paragraph mark, then each cell's text followed by
// U+0007, then U+FFFB + paragraph mark. Tables are created and changed through RTF (\trowd ... \cellx ... \cell \row),
// because RichEditorNET declares TOM's ITextRange2.Row with the wrong type. Rows are changed by rewriting them (a row
// inserted between two rows would end up nested in the next one).
// RichEditorNET's HTML writes each cell as a paragraph between "row" markers and cannot read <table>; ExportHtml and
// LoadHtml do that part.
using System.Drawing;
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
	const string Border = @"\clbrdrt\brdrw10\brdrs\clbrdrl\brdrw10\brdrs\clbrdrb\brdrw10\brdrs\clbrdrr\brdrw10\brdrs ";

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

	/// RTF for a table: cells holds each cell's text (plain), or null for empty cells; widths in twips per column.
	public static string BuildRtf(int rows, IList<int> widths, bool header, Func<int, int, string?>? cells = null) {
		var sb = new StringBuilder(@"{\rtf1\ansi ");
		for (int r = 0; r < rows; r++) {
			sb.Append(@"\trowd\trgaph108\trleft0");
			int x = 0;
			foreach (int width in widths) { x += width; sb.Append(Border).Append(@"\cellx").Append(x); }
			sb.Append(@"\pard\intbl ");
			for (int c = 0; c < widths.Count; c++) {
				string text = Escape(cells?.Invoke(r, c) ?? "");
				// a header cell's mark is bold too, so that text typed into the empty cell is bold
				sb.Append(header && r == 0 ? @"\b " + text + @"\cell\b0 " : text + @"\cell ");
			}
			sb.Append(@"\row ");
		}
		return sb.Append('}').ToString();
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
		var widths = RowWidths(editor);
		var rows = tables.SelectMany(t => t.Rows.Select((row, i) => (Table: t, Row: row, First: i == 0, Last: i == t.Rows.Count - 1))).ToList();
		if (widths.Count != rows.Count) widths = new List<List<int>>();
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
			if (first) sb.Append("<table style=\"border-collapse: collapse\">");
			sb.Append("<tr>");
			for (int c = 0; c < cells.Count; c++) {
				string width = rowIndex < widths.Count && c < widths[rowIndex].Count ? $" width: {widths[rowIndex][c] / 20}pt;" : "";
				sb.Append($"<td style=\"border: 1px solid #000; padding: 2px 6px;{width}\">").Append(cells[c]).Append("</td>");
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
			var rows = ParseTable(tables[t], out bool header, out var widths);
			if (rows.Count == 0) continue;
			int columns = rows.Max(r => r.Count);
			if (widths == null || widths.Count != columns) widths = EqualWidths(columns, widthTwips);
			editor.Select(range.Start, range.End - range.Start);
			editor.SelectedText = "";
			string? Plain(int r, int c) => c < rows[r].Count && IsPlain(rows[r][c].Html) ? WebUtility.HtmlDecode(rows[r][c].Html).Trim() : null;
			int tableStart = InsertRtf(editor, BuildRtf(rows.Count, widths, header, Plain));
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

	static List<List<HtmlCell>> ParseTable(string table, out bool header, out List<int>? widths) {
		var rows = new List<List<HtmlCell>>();
		header = false;
		widths = null;
		foreach (var (start, end) in Elements(table, "tr")) {
			string tr = table[start..end];
			var cells = new List<HtmlCell>();
			var rowWidths = new List<int>();
			foreach (Match cell in Regex.Matches(tr, @"<(td|th)\b([^>]*)>(.*?)</\1>", RegexOptions.IgnoreCase | RegexOptions.Singleline)) {
				if (rows.Count == 0 && cell.Groups[1].Value.Equals("th", StringComparison.OrdinalIgnoreCase)) header = true;
				cells.Add(new HtmlCell(cell.Groups[3].Value.Trim()));
				var w = Regex.Match(cell.Groups[2].Value, @"width\s*[:=]\s*""?([\d.]+)\s*(pt|px)?", RegexOptions.IgnoreCase);
				if (w.Success && float.TryParse(w.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float n))
					rowWidths.Add((int)(w.Groups[2].Value.Equals("pt", StringComparison.OrdinalIgnoreCase) ? n * 20 : n * 15));
			}
			if (rows.Count == 0 && rowWidths.Count == cells.Count && cells.Count > 0) widths = rowWidths;
			if (rows.Count == 0 && Regex.IsMatch(table[..start], @"<thead\b", RegexOptions.IgnoreCase)) header = true;
			if (cells.Count > 0) rows.Add(cells);
		}
		return rows;
	}
}
