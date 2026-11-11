// (iteration 92) Merge Cells, Split Cells and Delete Cells for AhkWord's tables (AhkWord.cs, through #load), as Word has
// them. RichEdit has no commands for these: a table's rows are read as RTF, changed and written back at once
// (RtfTables.RewriteRows: one undo step, the rows of a vertical merge together), and cells are found by position (their
// edges, \cellxN), not by number, so that cells merged across columns are reckoned with.
//   - a merge across columns is one cell as wide as the columns (as RichEdit's and Word's RTF have it);
//   - a merge down rows is \clvmgf on its first row's cell and \clvmrg on the cells below it (RichEdit shows them
//     as one cell; their text is the first's).
// The selected cells are those from the cell where the selection starts to the one where it ends: the rows between, and
// the part of the table between the two cells' outer edges, widened to whole merged cells (as Word does).
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static partial class RtfTables
{
	sealed partial class RowRtf
	{
		/// Rows written after this one (Split Cells into rows).
		public List<RowRtf> After = new();

		/// "clvmgf" (a merge down starts here), "clvmrg" (it goes on here) or null: cell i's merge down.
		public string? MergeOf(int i) {
			var m = Regex.Match(Definitions[i], @"\\(clvmgf|clvmrg)(?![a-z])");
			return m.Success ? m.Groups[1].Value : null;
		}

		public void SetMergeOf(int i, string? word) {
			string definition = Regex.Replace(Definitions[i], @"\\clvm(gf|rg)(?![a-z])", "");
			Definitions[i] = word == null ? definition : @"\" + word + definition;
		}

		/// The cells wholly between left and right (twips).
		public List<int> CellsIn(int left, int right) =>
			Enumerable.Range(0, Definitions.Count).Where(i => { var (a, b) = Bounds(i); return a >= left - SameEdge && b <= right + SameEdge; }).ToList();

		/// The cell whose edges are left and right, or -1.
		public int CellAt(int left, int right) =>
			Enumerable.Range(0, Definitions.Count).FirstOrDefault(i => { var (a, b) = Bounds(i); return Math.Abs(a - left) <= SameEdge && Math.Abs(b - right) <= SameEdge; }, -1);

		static string WithEdge(string definition, int x) => Regex.Replace(definition, @"\\cellx-?\d+$", @"\cellx" + x);

		/// Every cell's paragraph properties written at its start. RichEdit writes them once for a row (\pard\intbl ...
		/// before the first cell's text) and the other cells carry them on: moved, merged or deleted, a cell would take
		/// another cell's alignment.
		public void SpellOutProperties() {
			string current = LeadingProperties == "" ? @"\pard\intbl" : LeadingProperties;
			for (int i = 0; i < Contents.Count; i++) {
				if (!Contents[i].TrimStart().StartsWith(@"\pard", StringComparison.Ordinal)) Contents[i] = current + " " + Contents[i];
				// the last paragraph properties in it (a nested table's, \itap2 and on, aren't the row's)
				foreach (Match m in Regex.Matches(MaskNested(Contents[i]), @"\\pard(?![a-z])(\\[a-z]+-?\d*)*"))
					if (!Regex.IsMatch(m.Value, @"\\itap([2-9]|\d\d)")) current = ParagraphOnly(Contents[i].Substring(m.Index, m.Length));
			}
		}

		/// An empty cell's text: its paragraph properties only.
		public string Empty(int i) => (Contents.Count > i ? Regex.Match(Contents[i].TrimStart(), @"^\\pard(?![a-z])(\\[a-z]+-?\d*)*").Value : "") is { Length: > 0 } p && !Regex.IsMatch(p, @"\\itap([2-9]|\d\d)")
			? ParagraphOnly(p) + " " : ParagraphOnly(LeadingProperties == "" ? @"\pard\intbl" : LeadingProperties) + " ";

		// Paragraph properties without the character formats RichEdit writes after them (\pard\intbl\qc\f2\fs22): those
		// belong to the text that follows, which has its own.
		static string ParagraphOnly(string properties) => Regex.Replace(properties,
			@"\\(f|fs|cf|cb|highlight|lang|langfe|langnp|expnd|expndtw|kerning|up|dn|charscalex|cs|b|i|ul|ulnone|uld|uldb|ulw|strike|super|sub|nosupersub|caps|scaps|v|plain)-?\d*(?![a-z])", "");

		/// Cells first..last made one: as wide as they were together, with text (the first's definition: its borders,
		/// shading, alignment).
		public void Join(int first, int last, string text) {
			int right = Edges()[last];
			string definition = WithEdge(Definitions[first], right);
			Definitions.RemoveRange(first, last - first + 1);
			Definitions.Insert(first, definition);
			Contents.RemoveRange(first, Math.Min(Contents.Count, last + 1) - first);
			Contents.Insert(first, text);
		}

		/// Cell i made count cells, as wide as it in equal parts; the first keeps its text, each its definition.
		public void Divide(int i, int count) {
			var (left, right) = Bounds(i);
			string definition = Definitions[i];
			string empty = Empty(i);
			for (int j = count - 1; j >= 1; j--) {
				Definitions.Insert(i + 1, WithEdge(definition, left + (int)Math.Round((right - left) * (double)(j + 1) / count)));
				Contents.Insert(Math.Min(i + 1, Contents.Count), empty);
			}
			Definitions[i] = WithEdge(definition, left + (int)Math.Round((right - left) / (double)count));
		}

		/// The same row, its cells empty.
		public RowRtf EmptyCopy() {
			var copy = Parse(Build());
			for (int i = 0; i < copy.Contents.Count; i++) copy.Contents[i] = copy.Empty(i);
			return copy;
		}
	}

	/// The selected cells: rows First..Last of a table, and the part Left..Right of them (twips).
	sealed record Block(int First, int Last, int Left, int Right);

	// The table the selection is in and its cells where the selection starts and ends (a table nested in a cell: refused,
	// as the other commands that change its rows and columns).
	static (Table Table, int Row, int Cell, int EndRow, int EndCell) SelectionCells(RichEditBox editor) {
		if (BlockCells(editor) is { } block) return block;
		var selection = editor.TextDocument.Selection2;
		int start = selection.Start, end = Math.Max(selection.Start, selection.End - 1);
		var (table, r, c) = CurrentOuterOnly(editor);
		var last = At(editor, end);
		if (last is not { } l || l.Table.Start != table.Start || Innermost(editor, end) is { Nested: true }) return (table, r, c, r, c);
		return r <= l.Row ? (table, r, c, l.Row, l.Cell) : (table, l.Row, l.Cell, r, c);
	}

	// The block from the two cells, widened to whole cells (merged across) and whole merges down.
	static Block BlockOf(List<RowRtf> rows, int row, int cell, int endRow, int endCell) {
		var (a1, b1) = rows[row].Bounds(cell);
		var (a2, b2) = rows[endRow].Bounds(endCell);
		int first = row, last = endRow, left = Math.Min(a1, a2), right = Math.Max(b1, b2);
		for (bool changed = true; changed;) {
			changed = false;
			for (int k = first; k <= last; k++)
				for (int i = 0; i < rows[k].Definitions.Count; i++) {
					var (a, b) = rows[k].Bounds(i);
					if (a >= right - SameEdge || b <= left + SameEdge) continue;    // (not over the block)
					if (a < left - SameEdge) { left = a; changed = true; }
					if (b > right + SameEdge) { right = b; changed = true; }
				}
			if (first > 0 && rows[first].CellsIn(left, right).Any(i => rows[first].MergeOf(i) == "clvmrg")) { first--; changed = true; }
			if (last + 1 < rows.Count && rows[last + 1].CellsIn(left, right).Any(i => rows[last + 1].MergeOf(i) == "clvmrg")) { last++; changed = true; }
		}
		return new Block(first, last, left, right);
	}

	// The text of the cells the table had when the command started (the story's: what shows in each).
	static Func<int, int, string> CellTexts(RichEditBox editor, Table table) {
		string s = Story(editor.TextDocument);
		return (row, cell) => row < table.Rows.Count && cell < table.Rows[row].Cells.Count ? s[table.Rows[row].Cells[cell].Start..table.Rows[row].Cells[cell].End] : "";
	}

	static bool Blank(string text) => text.Trim(' ', '\r', '\n', '\t', '\v', MergedMark) == "";

	// ================================================================ Merge Cells

	/// Merges the selected cells into one (the caret's cell and the cells it is merged with: nothing to do). Their text
	/// goes into it, each cell's as paragraphs of its own, as Word does.
	public static void MergeCells(RichEditBox editor) {
		var (table, r0, c0, r1, c1) = SelectionCells(editor);
		var texts = CellTexts(editor, table);
		int row = r0, cell = c0;
		bool merged = false;
		RewriteRows(editor, table, rows => {
			foreach (var each in rows) each.SpellOutProperties();
			var block = BlockOf(rows, r0, c0, r1, c1);
			if (IsOneCell(rows, block)) throw new InvalidOperationException(T("status.table.mergeOne", "Select the cells to merge: two or more."));
			cell = Merge(rows, block, texts);
			row = block.First;
			merged = true;
		});
		if (merged) CaretTo(editor, table.Start, row, cell);
	}

	// Whether the block is already one cell merged down.
	static bool IsOneMerge(List<RowRtf> rows, Block block) {
		for (int k = block.First; k <= block.Last; k++) {
			var cells = rows[k].CellsIn(block.Left, block.Right);
			if (cells.Count != 1) return false;
			string? word = rows[k].MergeOf(cells[0]);
			if (word != (k == block.First ? "clvmgf" : "clvmrg")) return false;
		}
		return true;
	}

	// The block's cells made one; returns that cell's number in the block's first row.
	static int Merge(List<RowRtf> rows, Block block, Func<int, int, string> texts) {
		// the text: each cell's that has any (a cell merged down holds its merge's text)
		var parts = new List<string>();
		for (int k = block.First; k <= block.Last; k++)
			foreach (int i in rows[k].CellsIn(block.Left, block.Right))
				if (!Blank(texts(k, i)) && i < rows[k].Contents.Count) parts.Add(rows[k].Contents[i]);
		int firstCell = -1;
		for (int k = block.First; k <= block.Last; k++) {
			var cells = rows[k].CellsIn(block.Left, block.Right);
			if (cells.Count == 0) continue;
			string text = k == block.First && parts.Count > 0 ? string.Join(@"\par ", parts) : rows[k].Empty(cells[0]);
			rows[k].Join(cells[0], cells[^1], text);
			rows[k].SetMergeOf(cells[0], block.Last == block.First ? null : k == block.First ? "clvmgf" : "clvmrg");
			if (k == block.First) firstCell = cells[0];
		}
		return Math.Max(0, firstCell);
	}

	// ================================================================ Split Cells

	/// The caret's cell (the selected cells, merged first) split into columns and rows. A cell merged down splits into
	/// a number of rows that divides the rows it spans (all of them: as before the merge); a cell of one row split into
	/// rows adds rows, the row's other cells merged down over them, as Word does. Text stays in the first cell.
	public static void SplitCells(RichEditBox editor, int columns, int rowCount) {
		columns = Math.Clamp(columns, 1, 63);
		rowCount = Math.Clamp(rowCount, 1, 100);
		if (columns == 1 && rowCount == 1) return;
		var (table, r0, c0, r1, c1) = SelectionCells(editor);
		var texts = CellTexts(editor, table);
		int row = r0, cell = c0;
		RewriteRows(editor, table, rows => {
			foreach (var each in rows) each.SpellOutProperties();
			var block = BlockOf(rows, r0, c0, r1, c1);
			if (!IsOneCell(rows, block)) Merge(rows, block, texts);
			row = block.First;
			cell = Split(rows, block, columns, rowCount);
		});
		CaretTo(editor, table.Start, row, cell);
	}

	/// The span of the cell at the caret: (columns, rows), as Split Cells' dialog starts with.
	public static (int Columns, int Rows) CellSpan(RichEditBox editor) {
		var (table, r0, c0, r1, c1) = SelectionCells(editor);
		var selection = editor.TextDocument.Selection2;
		int start = selection.Start, end = selection.End;
		List<RowRtf>? rows;
		try { rows = ReadRows(editor, table); }    // (reading a table's RTF selects it)
		finally { editor.TextDocument.Selection2.SetRange(start, end); }
		if (rows == null) return (2, 1);
		var block = BlockOf(rows, r0, c0, r1, c1);
		int span = block.Last - block.First + 1;
		// the table's column edges inside the cell: how many columns it covers
		var edges = rows.SelectMany(row => row.Edges()).Where(x => x > block.Left + SameEdge && x < block.Right - SameEdge).Distinct().OrderBy(x => x).ToList();
		int across = 1;
		for (int i = 0; i < edges.Count; i++) if (i == 0 || edges[i] - edges[i - 1] > SameEdge) across++;
		return (Math.Max(2, across), span);
	}

	static bool IsOneCell(List<RowRtf> rows, Block block) =>
		rows[block.First].CellsIn(block.Left, block.Right).Count == 1 && (block.Last == block.First || IsOneMerge(rows, block));

	// The block, one cell, split; returns the first new cell's number in the block's first row.
	static int Split(List<RowRtf> rows, Block block, int columns, int rowCount) {
		int span = block.Last - block.First + 1;
		var cellRows = Enumerable.Range(block.First, span).Select(k => rows[k]).ToList();
		if (span > 1) {
			// its rows in groups of span / rowCount, each merged down (one row each: no merge left)
			if (span % rowCount != 0)
				throw new InvalidOperationException(T("status.table.splitRows", "This cell spans {0} rows: it can be split into a number of rows that divides {0}.", span));
			int group = span / rowCount;
			for (int j = 0; j < span; j++) {
				int i = cellRows[j].CellAt(block.Left, block.Right);
				if (i >= 0) cellRows[j].SetMergeOf(i, group == 1 ? null : j % group == 0 ? "clvmgf" : "clvmrg");
			}
		}
		else if (rowCount > 1) {
			var first = rows[block.First];
			int own = first.CellAt(block.Left, block.Right);
			// the row's other cells merged down over the new rows
			for (int i = 0; i < first.Definitions.Count; i++)
				if (i != own && first.MergeOf(i) == null) first.SetMergeOf(i, "clvmgf");
			for (int n = 1; n < rowCount; n++) {
				var copy = first.EmptyCopy();
				for (int i = 0; i < copy.Definitions.Count; i++) copy.SetMergeOf(i, i == own ? null : "clvmrg");
				first.After.Add(copy);
				cellRows.Add(copy);
			}
		}
		if (columns > 1)
			foreach (var row in cellRows) {
				int i = row.CellAt(block.Left, block.Right);
				if (i >= 0) row.Divide(i, columns);
			}
		return Math.Max(0, rows[block.First].CellAt(block.Left, block.Left + (block.Right - block.Left) / Math.Max(1, columns)));
	}

	// ================================================================ Delete Cells

	/// Deletes the selected cells (the caret's): how = "left" (the cells after them in their rows move left), "up" (the
	/// cells below them move up, empty cells at the bottom), "row" (their rows) or "column" (their columns), as Word's
	/// Delete Cells does.
	public static void DeleteCells(RichEditBox editor, string how) {
		how = (how ?? "").Trim().ToLowerInvariant();
		var (table, r0, c0, r1, c1) = SelectionCells(editor);
		if (how is "row" or "rows") {
			DeleteRows(editor, table, r0, r1, c0);
			return;
		}
		if (how is not ("left" or "up" or "column" or "columns"))
			throw new ArgumentException(T("status.table.deleteHow", "Delete cells how? \"left\", \"up\", \"row\" or \"column\"."));
		if (how.StartsWith("column") && table.Rows.All(row => row.Cells.Count <= 1)) { DeleteTable(editor); return; }
		var texts = CellTexts(editor, table);
		int row = r0, cell = c0;
		RewriteRows(editor, table, rows => {
			foreach (var each in rows) each.SpellOutProperties();
			var block = BlockOf(rows, r0, c0, r1, c1);
			row = block.First;
			cell = Math.Max(0, rows[block.First].CellsIn(block.Left, block.Right).DefaultIfEmpty(0).First() - (how == "left" ? 1 : 0));
			if (how == "left") for (int k = block.First; k <= block.Last; k++) rows[k].DeleteSpan(block.Left, block.Right);
			else if (how.StartsWith("column")) foreach (var each in rows) each.DeleteSpan(block.Left, block.Right);
			else ShiftUp(rows, block, texts);
		});
		CaretTo(editor, table.Start, row, cell);    // (row and cell clamped to the table as it is now)
	}

	// The block's cells emptied by moving the cells below them up (their text; the cells stay): every row from the block
	// down must have the same cells there, none merged down.
	static void ShiftUp(List<RowRtf> rows, Block block, Func<int, int, string> texts) {
		var shape = rows[block.First].CellsIn(block.Left, block.Right).Select(i => rows[block.First].Bounds(i)).ToList();
		for (int k = block.First; k < rows.Count; k++) {
			var cells = rows[k].CellsIn(block.Left, block.Right);
			bool same = cells.Count == shape.Count && cells.Select((i, n) => (Bounds: rows[k].Bounds(i), n)).All(c => Math.Abs(c.Bounds.Left - shape[c.n].Left) <= SameEdge && Math.Abs(c.Bounds.Right - shape[c.n].Right) <= SameEdge);
			if (!same || cells.Any(i => rows[k].MergeOf(i) != null))
				throw new InvalidOperationException(T("status.table.shiftUp", "The cells below aren't the same shape (merged cells): delete the cells shifting the others left, or the whole rows."));
		}
		int height = block.Last - block.First + 1;
		for (int k = block.First; k < rows.Count; k++) {
			var cells = rows[k].CellsIn(block.Left, block.Right);
			var below = k + height < rows.Count ? rows[k + height].CellsIn(block.Left, block.Right) : null;
			for (int n = 0; n < cells.Count; n++)
				rows[k].Contents[cells[n]] = below != null ? rows[k + height].Contents[below[n]] : rows[k].Empty(cells[n]);
		}
	}

	// The table's rows as RewriteRows reads them (null: not as expected).
	static List<RowRtf>? ReadRows(RichEditBox editor, Table table) {
		string rtf = GetRtf(editor, table.Start, table.End);
		var matches = Regex.Matches(MaskNested(rtf), @"\\trowd(?![a-z]).*?\\row(?![a-z])", RegexOptions.Singleline).Cast<Match>().ToList();
		if (matches.Count != table.Rows.Count) return null;
		return matches.Select(m => RowRtf.Parse(rtf.Substring(m.Index, m.Length))).ToList();
	}
	// ================================================================ a block of cells (AhkWord's own selection)

	/// Whether two positions are in cells of different rows of one (outermost) table: a drag between them selects a
	/// block of cells (RichEdit would select the whole rows).
	public static bool SpansRows(RichEditBox editor, int from, int to) =>
		At(editor, from) is { } a && At(editor, to) is { } b && a.Table.Start == b.Table.Start && a.Row != b.Row
		&& Innermost(editor, from) is { Nested: false } && Innermost(editor, to) is { Nested: false };

	/// Selects the cells from row1, cell1 to row2, cell2 (from 1) of the table at the caret as a block (the rows between,
	/// the columns between the two cells): Merge, Split and Delete Cells act on them.
	public static void SelectCells(RichEditBox editor, int row1, int cell1, int row2, int cell2) {
		var table = Innermost(editor, editor.TextDocument.Selection2.Start) is { Nested: false } at ? at.Table
			: Find(editor.TextDocument).FirstOrDefault() ?? throw new InvalidOperationException(T("status.table.none", "The document has no table."));
		int Position(int row, int cell) {
			if (row < 1 || row > table.Rows.Count) throw new ArgumentOutOfRangeException(nameof(row), T("status.table.noRow", "The table has no row {0}.", row));
			var cells = table.Rows[row - 1].Cells;
			if (cell < 1 || cell > cells.Count) throw new ArgumentOutOfRangeException(nameof(cell), T("status.table.noCell", "Row {0} has no cell {1}.", row, cell));
			return cells[cell - 1].Start;
		}
		int from = Position(row1, cell1), to = Position(row2, cell2);
		if (editor is WriterEditBox writer) writer.SelectBlock(from, to);
		else editor.TextDocument.Selection2.SetRange(Math.Min(from, to), Math.Max(from, to));
	}

	/// Empties the selected cells' text (a block's: Delete and Backspace), their paragraph formats kept; one undo step.
	public static void ClearCells(RichEditBox editor) {
		var (table, r0, c0, r1, c1) = SelectionCells(editor);
		RewriteRows(editor, table, rows => {
			foreach (var each in rows) each.SpellOutProperties();
			var block = BlockOf(rows, r0, c0, r1, c1);
			for (int k = block.First; k <= block.Last; k++)
				foreach (int i in rows[k].CellsIn(block.Left, block.Right))
					if (i < rows[k].Contents.Count && rows[k].MergeOf(i) != "clvmrg") rows[k].Contents[i] = rows[k].Empty(i);
		});
		CaretTo(editor, table.Start, r0, c0);
	}

	/// The edges of the cells of the table at a position, in twips: for each row its left edge, then each cell's right
	/// edge (null: not read). Read once when a block of cells is made (reading the RTF moves the selection), for painting it.
	public static List<List<int>>? EdgesOf(RichEditBox editor, int position) {
		if (At(editor, position) is not { } at) return null;
		var selection = editor.TextDocument.Selection2;
		int start = selection.Start, end = selection.End;
		try { return ReadRows(editor, at.Table)?.Select(row => new List<int> { row.Left }.Concat(row.Edges()).ToList()).ToList(); }
		finally { editor.TextDocument.Selection2.SetRange(start, end); }
	}

	const int EM_POSFROMCHAR = 0x00D6;
	[System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessageW")]
	static extern IntPtr SendPoint(IntPtr hWnd, int msg, ref System.Drawing.Point point, IntPtr cp);

	// Where a character is in the window (its line's top left), by its TOM position.
	static System.Drawing.Point PointOf(RichEditBox editor, int cp) {
		var point = new System.Drawing.Point();
		SendPoint(editor.Handle, EM_POSFROMCHAR, ref point, (IntPtr)Math.Max(0, cp));
		return point;
	}

	/// The block's cells on the screen (client rectangles), as Merge, Split and Delete Cells will take them: the rows
	/// between the two positions' cells, the columns between them, widened to whole merged cells. Null: not a block.
	public static List<System.Drawing.Rectangle>? BlockRectangles(RichEditBox editor, int from, int to, List<List<int>>? edgesTwips = null) {
		if (!editor.IsHandleCreated || At(editor, from) is not { } a || At(editor, to) is not { } b || a.Table.Start != b.Table.Start || a.Row == b.Row && a.Cell == b.Cell) return null;
		var table = a.Table;
		string s = Story(editor.TextDocument);
		float pixelsPerTwip = editor.DeviceDpi / 1440f * Math.Max(0.1f, ((RichTextBox)editor).ZoomFactor);
		int gap = (int)Math.Round(108 * pixelsPerTwip);    // (the cell's padding before its text: \trgaph108, RichEdit's and Word's own)
		// each row's cells, left and right edges in pixels: the first's from its text, the others from the edges in twips
		// (the block's, read when it was made), or from where their text starts
		var twips = edgesTwips ?? (editor as WriterEditBox)?.BlockEdges;
		if (twips != null && twips.Count != table.Rows.Count) twips = null;
		var edges = new List<List<(int Left, int Right)>>();
		var tops = new List<int>();
		// the table's left edge, from where its first cells' text starts: a left-aligned one's (centred or right-aligned
		// text starts further in), the leftmost (an indent or a bullet starts further in too); rows of the same table
		// start at the same edge as far as RichEdit draws them
		var starts = Enumerable.Range(0, table.Rows.Count).Where(k => table.Rows[k].Cells.Count > 0).Select(k => {
			int at = table.Rows[k].Cells[0].Start;
			int alignment = 0;
			try { alignment = editor.TextDocument.Range2(at, at).Para.Alignment; } catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException or ArgumentException) { }
			return (Row: k, X: PointOf(editor, at).X, Left: alignment == 0 && s[at] != MergedMark);
		}).ToList();
		var leftAligned = starts.Where(x => x.Left).ToList();
		var anchor = (leftAligned.Count > 0 ? leftAligned : starts).OrderBy(x => x.X).FirstOrDefault();
		int anchorRow = anchor.Row, tableLeft = anchor.X - gap;
		for (int k = 0; k < table.Rows.Count; k++) {
			var row = table.Rows[k];
			var cells = new List<(int, int)>();
			if (twips != null && twips[k].Count == row.Cells.Count + 1) {
				int Pixel(int x) => tableLeft + (int)Math.Round((x - twips[anchorRow][0]) * pixelsPerTwip);
				for (int i = 0; i < row.Cells.Count; i++) cells.Add((Pixel(twips[k][i]), Pixel(twips[k][i + 1])));
			}
			else
				for (int i = 0; i < row.Cells.Count; i++) {
					int x = PointOf(editor, row.Cells[i].Start).X - gap;
					int next = i + 1 < row.Cells.Count ? PointOf(editor, row.Cells[i + 1].Start).X - gap : x + (int)(1440 * pixelsPerTwip);
					cells.Add((x, next));
				}
			edges.Add(cells);
			tops.Add(row.Cells.Count > 0 ? PointOf(editor, row.Cells[0].Start).Y - 1 : PointOf(editor, row.Start).Y);
		}
		int bottomOfTable = PointOf(editor, Math.Min(table.End, Math.Max(0, s.Length - 1))).Y - 1;
		int Bottom(int k) => k + 1 < tops.Count ? tops[k + 1] : bottomOfTable;
		const int near = 3;
		int r0 = Math.Min(a.Row, b.Row), r1 = Math.Max(a.Row, b.Row);
		int L = Math.Min(edges[a.Row][a.Cell].Left, edges[b.Row][b.Cell].Left), R = Math.Max(edges[a.Row][a.Cell].Right, edges[b.Row][b.Cell].Right);
		bool Inside((int Left, int Right) c) => c.Left >= L - near && c.Right <= R + near;
		bool Continues(int k) => Enumerable.Range(0, edges[k].Count).Any(i => Inside(edges[k][i]) && table.Rows[k].Cells[i] is var cell && cell.End == cell.Start + 1 && s[cell.Start] == MergedMark);
		for (bool changed = true; changed;) {
			changed = false;
			for (int k = r0; k <= r1; k++)
				foreach (var c in edges[k]) {
					if (c.Left >= R - near || c.Right <= L + near) continue;
					if (c.Left < L - near) { L = c.Left; changed = true; }
					if (c.Right > R + near) { R = c.Right; changed = true; }
				}
			if (r0 > 0 && Continues(r0)) { r0--; changed = true; }
			if (r1 + 1 < table.Rows.Count && Continues(r1 + 1)) { r1++; changed = true; }
		}
		var rectangles = new List<System.Drawing.Rectangle>();
		for (int k = r0; k <= r1; k++)
			foreach (var c in edges[k].Where(Inside))
				rectangles.Add(System.Drawing.Rectangle.FromLTRB(c.Left + 1, tops[k] + 1, c.Right, Bottom(k)));
		return rectangles;
	}

	// The commands' cells from a block of cells, if there is one (null: RichEdit's selection is the selection).
	static (Table Table, int Row, int Cell, int EndRow, int EndCell)? BlockCells(RichEditBox editor) {
		if (editor is not WriterEditBox { CellBlock: { } block }) return null;
		if (At(editor, block.From) is not { } a || At(editor, block.To) is not { } b || a.Table.Start != b.Table.Start) return null;
		return a.Row <= b.Row ? (a.Table, a.Row, a.Cell, b.Row, b.Cell) : (a.Table, b.Row, b.Cell, a.Row, a.Cell);
	}

	/// Whether a block of cells is selected (AhkWord's own selection, see WriterEditBox.CellBlock).
	public static bool HasBlock(RichEditBox editor) => BlockCells(editor) != null;
}
