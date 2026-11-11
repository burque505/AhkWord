// Iteration 91: tables from data, and data from tables (used by AhkWord.cs through #load).
// RichEdit's tables are text with marks in it (rte_tables2.cs): they can't be bound to data. Instead, data are read into
// rows of cells here and put into a table (InsertTableFromData, InsertTableFromFile, SetTableData: EditorApi), and a
// table's cells are read back out as CSV, TSV or JSON (TableData). The sources:
//   - CSV and TSV text or files: RFC 4180 (quoted fields, "" for a quote, line breaks inside quotes); the delimiter is a
//     tab, ";" (what Excel saves where the decimal mark is a comma) or ",", whichever the first line has most of, or
//     Excel's "sep=;" first line. Files: UTF-8 or UTF-16 (with or without a byte order mark), else the ANSI code page.
//   - JSON: an array of rows (each an array of values), or of objects (their names make the first row).
//   - Excel workbooks (.xlsx, .xlsm), read with the Open XML SDK (no Excel needed; the file may be open in Excel): a
//     sheet by name or number, the cells in use or a range ("A1:D20"), values as Excel shows them (number formats,
//     dates, percentages) in this computer's culture; formulas' last computed values.
// Not read: old .xls workbooks, .ods (save them as .xlsx or CSV).
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public static class TabularData
{
	public const int MaxRows = 10000, MaxColumns = 63;

	// ================================================================ text

	/// Rows of cells from text: format "csv", "tsv", "json" or "" (JSON if it starts with [ or {, else CSV/TSV by its
	/// delimiter). keysRow: JSON objects' names make a first row.
	public static List<List<string>> Parse(string data, string format = "", bool keysRow = true) {
		data ??= "";
		if (data.Length > 0 && data[0] == '\uFEFF') data = data[1..];
		string kind = (format ?? "").Trim().ToLowerInvariant();
		if (kind == "") kind = data.TrimStart() is { Length: > 0 } t && t[0] is '[' or '{' ? "json" : "csv";
		var rows = kind switch {
			"json" => ParseJson(data, keysRow),
			"tsv" or "tab" => ParseCsv(data, '\t'),
			"csv" => ParseCsv(data, null),
			_ => throw new ArgumentException($"\"{format}\" is not a data format: use \"csv\", \"tsv\" or \"json\"."),
		};
		return Tidy(rows);
	}

	/// CSV (or TSV) text as rows; delimiter null: worked out from the first line.
	public static List<List<string>> ParseCsv(string text, char? delimiter) {
		if (delimiter == null) {
			var sep = Regex.Match(text, @"\A\s*sep=(.)\s*?(\r\n|\n|\r)");
			if (sep.Success) { delimiter = sep.Groups[1].Value[0]; text = text[sep.Length..]; }
			else delimiter = GuessDelimiter(text);
		}
		char d = delimiter.Value;
		var rows = new List<List<string>>();
		var row = new List<string>();
		var field = new StringBuilder();
		bool quoted = false, wasQuoted = false;
		for (int i = 0; i < text.Length; i++) {
			char c = text[i];
			if (quoted) {
				if (c == '"') {
					if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
					else quoted = false;
				} else field.Append(c);
			} else if (c == '"' && field.Length == 0 && !wasQuoted) { quoted = wasQuoted = true; }
			else if (c == d) { row.Add(field.ToString()); field.Clear(); wasQuoted = false; }
			else if (c is '\r' or '\n') {
				if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
				row.Add(field.ToString()); field.Clear(); wasQuoted = false;
				rows.Add(row); row = new List<string>();
			} else field.Append(c);
		}
		if (field.Length > 0 || row.Count > 0 || wasQuoted) { row.Add(field.ToString()); rows.Add(row); }
		// (line breaks in quoted fields: \r\n as \n)
		foreach (var r in rows) for (int i = 0; i < r.Count; i++) if (r[i].Contains('\r')) r[i] = r[i].Replace("\r\n", "\n").Replace('\r', '\n');
		return rows;
	}

	// The delimiter the first line (outside quotes) has most of: tab, ";" or ","; "," if none.
	static char GuessDelimiter(string text) {
		int tabs = 0, semis = 0, commas = 0;
		bool quoted = false;
		foreach (char c in text) {
			if (c == '"') quoted = !quoted;
			else if (quoted) continue;
			else if (c is '\r' or '\n') break;
			else if (c == '\t') tabs++;
			else if (c == ';') semis++;
			else if (c == ',') commas++;
		}
		if (tabs > 0 && tabs >= semis && tabs >= commas) return '\t';
		return semis > commas ? ';' : ',';
	}

	/// JSON rows: an array of arrays (or of single values: one column), or an array of objects (or one object).
	public static List<List<string>> ParseJson(string json, bool keysRow) {
		JsonDocument parsed;
		try { parsed = JsonDocument.Parse(json, new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip }); }
		catch (JsonException ex) { throw new ArgumentException("The data is not valid JSON: " + ex.Message); }
		using var doc = parsed;
		var root = doc.RootElement;
		var items = root.ValueKind == JsonValueKind.Array ? root.EnumerateArray().ToList() : new List<JsonElement> { root };
		var rows = new List<List<string>>();
		if (items.Count > 0 && items.All(i => i.ValueKind is JsonValueKind.Object or JsonValueKind.Null)) {
			var names = new List<string>();
			foreach (var item in items.Where(i => i.ValueKind == JsonValueKind.Object))
				foreach (var p in item.EnumerateObject()) if (!names.Contains(p.Name)) names.Add(p.Name);
			if (keysRow) rows.Add(names.ToList());
			foreach (var item in items)
				rows.Add(names.Select(n => item.ValueKind == JsonValueKind.Object && item.TryGetProperty(n, out var v) ? Value(v) : "").ToList());
			return rows;
		}
		foreach (var item in items)
			rows.Add(item.ValueKind == JsonValueKind.Array ? item.EnumerateArray().Select(Value).ToList() : new List<string> { Value(item) });
		return rows;
	}

	/// The names of JSON objects' values, in order (null if the data aren't JSON objects).
	public static List<string>? JsonNames(string data, string format) {
		string kind = (format ?? "").Trim().ToLowerInvariant();
		if (kind is not ("" or "json")) return null;
		string t = (data ?? "").TrimStart('\uFEFF').TrimStart();
		if (t.Length == 0 || t[0] is not ('[' or '{')) return null;
		var rows = ParseJson(t, true);
		using var doc = JsonDocument.Parse(t, new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip });
		var root = doc.RootElement;
		bool objects = root.ValueKind == JsonValueKind.Object || root.ValueKind == JsonValueKind.Array && root.GetArrayLength() > 0
			&& root.EnumerateArray().All(i => i.ValueKind is JsonValueKind.Object or JsonValueKind.Null);
		return objects && rows.Count > 0 ? rows[0] : null;
	}

	static string Value(JsonElement v) => v.ValueKind switch {
		JsonValueKind.String => v.GetString() ?? "",
		JsonValueKind.Null or JsonValueKind.Undefined => "",
		JsonValueKind.True => "true",
		JsonValueKind.False => "false",
		_ => v.GetRawText(),    // (numbers as written; objects and arrays as JSON)
	};

	// Trailing empty rows dropped, then every row as long as the longest; the limits checked.
	static List<List<string>> Tidy(List<List<string>> rows) {
		while (rows.Count > 0 && rows[^1].All(c => c == "")) rows.RemoveAt(rows.Count - 1);
		int columns = rows.Count == 0 ? 0 : rows.Max(r => r.Count);
		foreach (var r in rows) while (r.Count < columns) r.Add("");
		if (rows.Count > MaxRows) throw new ArgumentException($"The data has {rows.Count} rows; a table can have at most {MaxRows}.");
		if (columns > MaxColumns) throw new ArgumentException($"The data has {columns} columns; a table can have at most {MaxColumns}.");
		return rows;
	}

	/// Rows as text: "csv" (comma), "csv;" (semicolon), "tsv", "json" (an array of rows) or "objects" (an array of
	/// objects named by the first row).
	public static string Write(List<List<string>> rows, string format) {
		switch ((format ?? "").Trim().ToLowerInvariant()) {
			case "" or "csv": return Delimited(rows, ',');
			case "csv;": return Delimited(rows, ';');
			case "tsv" or "tab": return Delimited(rows, '\t');
			case "json": return JsonSerializer.Serialize(rows, JsonOptions);
			case "objects":
				if (rows.Count == 0) return "[]";
				var names = rows[0].Select((n, i) => n.Trim() == "" ? $"Column{i + 1}" : n).ToList();
				for (int i = 0; i < names.Count; i++)    // (names made unique: "Name", "Name 2")
					for (int k = 2; names.Take(i).Contains(names[i]); k++) names[i] = rows[0][i] + " " + k;
				var list = rows.Skip(1).Select(r => names.Select((n, i) => (n, v: i < r.Count ? r[i] : "")).ToDictionary(p => p.n, p => p.v)).ToList();
				return JsonSerializer.Serialize(list, JsonOptions);
			default: throw new ArgumentException($"\"{format}\" is not a data format: use \"csv\", \"csv;\", \"tsv\", \"json\" or \"objects\".");
		}
	}

	static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

	static string Delimited(List<List<string>> rows, char d) =>
		string.Join("\r\n", rows.Select(r => string.Join(d, r.Select(c => c.IndexOfAny(new[] { d, '"', '\n', '\r' }) >= 0 || c != c.Trim() && c != "" ? "\"" + c.Replace("\"", "\"\"") + "\"" : c))));

	// ================================================================ files

	/// Rows from a file: .xlsx/.xlsm (sheet: a name or number, "" for the first; range such as "A1:D20", "" for the
	/// cells in use), .json, or CSV/TSV text (.csv, .tsv, .txt and others).
	public static List<List<string>> ReadFile(string path, string sheet = "", string range = "") {
		if (!File.Exists(path)) throw new FileNotFoundException("There is no file " + path, path);
		string ext = Path.GetExtension(path).ToLowerInvariant();
		if (ext is ".xls" or ".ods" or ".xlsb") throw new ArgumentException($"{Path.GetFileName(path)}: this kind of workbook can't be read; save it as .xlsx or CSV.");
		if (ext is ".xlsx" or ".xlsm" or ".xltx" or ".xltm") return Tidy(ReadXlsx(path, sheet, range));
		if ((sheet ?? "") != "" || (range ?? "") != "") throw new ArgumentException("A sheet or range is for Excel workbooks (.xlsx) only.");
		string text = ReadText(path);
		return Parse(text, ext == ".json" ? "json" : ext == ".tsv" ? "tsv" : "");
	}

	/// A text file's text: UTF-8 or UTF-16 by its byte order mark, else UTF-8 if it is valid UTF-8, else the ANSI code
	/// page (what Excel's "CSV" saves). The file may be open in another program.
	public static string ReadText(string path) {
		byte[] bytes;
		using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
			bytes = new byte[stream.Length];
			stream.ReadExactly(bytes);
		}
		if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
		if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
		if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
		try { return new UTF8Encoding(false, true).GetString(bytes); }
		catch (DecoderFallbackException) {
			int codePage = CultureInfo.CurrentCulture.TextInfo.ANSICodePage;
			var ansi = CodePagesEncodingProvider.Instance.GetEncoding(codePage) ?? CodePagesEncodingProvider.Instance.GetEncoding(1252) ?? Encoding.Latin1;
			return ansi.GetString(bytes);
		}
	}

	/// The names of an Excel workbook's sheets, in order.
	public static List<string> SheetNames(string path) {
		using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
		using var book = SpreadsheetDocument.Open(stream, false);
		return book.WorkbookPart?.Workbook?.Sheets?.Elements<Sheet>().Select(s => s.Name?.Value ?? "").ToList() ?? new List<string>();
	}

	static List<List<string>> ReadXlsx(string path, string sheetName, string range) {
		using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
		SpreadsheetDocument book;
		try { book = SpreadsheetDocument.Open(stream, false); }
		catch (Exception ex) when (ex is DocumentFormat.OpenXml.Packaging.OpenXmlPackageException or InvalidDataException or IOException) {
			throw new ArgumentException($"{Path.GetFileName(path)} could not be read as an Excel workbook: {ex.Message}");
		}
		using var _ = book;
		var workbook = book.WorkbookPart ?? throw new ArgumentException($"{Path.GetFileName(path)} has no workbook in it.");
		var sheets = workbook.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
		if (sheets.Count == 0) throw new ArgumentException($"{Path.GetFileName(path)} has no sheets.");
		Sheet sheet;
		sheetName = (sheetName ?? "").Trim();
		if (sheetName == "") sheet = sheets.FirstOrDefault(s => s.State == null || s.State.Value == SheetStateValues.Visible) ?? sheets[0];
		else sheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
			?? (int.TryParse(sheetName, out int n) && n >= 1 && n <= sheets.Count ? sheets[n - 1]
			: throw new ArgumentException($"{Path.GetFileName(path)} has no sheet \"{sheetName}\": its sheets are {string.Join(", ", sheets.Select(s => "\"" + s.Name + "\""))}."));
		if (workbook.GetPartById(sheet.Id!.Value!) is not WorksheetPart part) throw new ArgumentException($"\"{sheet.Name}\" is a chart sheet, not a sheet of cells.");

		var shared = workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(ItemText).ToList() ?? new List<string>();
		var styles = workbook.WorkbookStylesPart?.Stylesheet;
		var formatIds = styles?.CellFormats?.Elements<CellFormat>().Select(f => f.NumberFormatId?.Value ?? 0u).ToList() ?? new List<uint>();
		var codes = styles?.NumberingFormats?.Elements<NumberingFormat>().Where(f => f.NumberFormatId != null)
			.GroupBy(f => f.NumberFormatId!.Value).ToDictionary(g => g.Key, g => g.First().FormatCode?.Value ?? "") ?? new Dictionary<uint, string>();
		bool date1904 = workbook.Workbook.WorkbookProperties?.Date1904?.Value == true;

		(int r1, int c1, int r2, int c2)? limits = null;
		if ((range ?? "").Trim() != "") {
			var m = Regex.Match(range!.Trim().Replace("$", ""), @"^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$");
			if (!m.Success) throw new ArgumentException($"\"{range}\" is not a range: use one such as \"A1:D20\".");
			int ra = int.Parse(m.Groups[2].Value), ca = ColumnNumber(m.Groups[1].Value);
			int rb = m.Groups[3].Success ? int.Parse(m.Groups[4].Value) : ra, cb = m.Groups[3].Success ? ColumnNumber(m.Groups[3].Value) : ca;
			limits = (Math.Min(ra, rb), Math.Min(ca, cb), Math.Max(ra, rb), Math.Max(ca, cb));
		}

		var cells = new Dictionary<(int Row, int Column), string>();
		int lastRow = 0;
		foreach (var row in (part.Worksheet?.Descendants<Row>() ?? Enumerable.Empty<Row>())) {
			int r = row.RowIndex?.Value is uint ri ? (int)ri : lastRow + 1;
			lastRow = r;
			int lastColumn = 0;
			foreach (var cell in row.Elements<Cell>()) {
				int c = cell.CellReference?.Value is string reference ? ColumnNumber(Regex.Match(reference, "^[A-Za-z]+").Value) : lastColumn + 1;
				lastColumn = c;
				if (limits is var (lr1, lc1, lr2, lc2) && (r < lr1 || r > lr2 || c < lc1 || c > lc2)) continue;
				string text = CellText(cell, shared, formatIds, codes, date1904);
				if (text != "") cells[(r, c)] = text;
				if (cells.Count > MaxRows * MaxColumns) throw new ArgumentException("The sheet has too many cells for a table: give a range.");
			}
		}
		if (cells.Count == 0 && limits == null) return new List<List<string>>();
		int top = limits?.r1 ?? cells.Keys.Min(k => k.Row), bottom = limits?.r2 ?? cells.Keys.Max(k => k.Row);
		int left = limits?.c1 ?? cells.Keys.Min(k => k.Column), right = limits?.c2 ?? cells.Keys.Max(k => k.Column);
		if (bottom - top + 1 > MaxRows) throw new ArgumentException($"The sheet's cells take {bottom - top + 1} rows; a table can have at most {MaxRows}: give a range.");
		if (right - left + 1 > MaxColumns) throw new ArgumentException($"The sheet's cells take {right - left + 1} columns; a table can have at most {MaxColumns}: give a range.");
		var rows = new List<List<string>>();
		for (int r = top; r <= bottom; r++)
			rows.Add(Enumerable.Range(left, right - left + 1).Select(c => cells.TryGetValue((r, c), out var t) ? t : "").ToList());
		return rows;
	}

	// A shared string's text: its runs' text, without phonetic hints (rPh).
	static string ItemText(SharedStringItem item) =>
		item.Text?.Text ?? string.Concat(item.Elements<Run>().Select(r => r.Text?.Text ?? ""));

	/// "A" = 1, "Z" = 26, "AA" = 27.
	public static int ColumnNumber(string letters) {
		int n = 0;
		foreach (char ch in letters.ToUpperInvariant()) n = n * 26 + (ch - 'A' + 1);
		return n;
	}

	static string CellText(Cell cell, List<string> shared, List<uint> formatIds, Dictionary<uint, string> codes, bool date1904) {
		string raw = cell.CellValue?.Text ?? "";
		var type = cell.DataType?.Value;
		if (type == CellValues.SharedString) return int.TryParse(raw, out int i) && i >= 0 && i < shared.Count ? shared[i] : "";
		if (type == CellValues.InlineString) return cell.InlineString?.Text?.Text ?? string.Concat(cell.InlineString?.Elements<Run>().Select(r => r.Text?.Text ?? "") ?? Array.Empty<string>());
		if (type == CellValues.Boolean) return raw == "1" ? "TRUE" : raw == "0" ? "FALSE" : raw;
		if (type == CellValues.String || type == CellValues.Error) return raw;
		if (raw == "") return "";
		int style = (int)(cell.StyleIndex?.Value ?? 0);
		uint id = style < formatIds.Count ? formatIds[style] : 0;
		string? code = codes.TryGetValue(id, out var own) ? own : null;
		if (type == CellValues.Date) {
			if (!DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var when)) return raw;
			return FormatDate(when, id, code);
		}
		if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) return raw;
		return FormatNumber(value, id, code, date1904);
	}

	// ================================================================ Excel's number formats

	// The built-in formats Excel doesn't write out (ids 0-49); the dates among them are handled by FormatDate.
	static readonly Dictionary<uint, string> BuiltIn = new() {
		[1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00", [9] = "0%", [10] = "0.00%", [11] = "0.00E+00",
		[12] = "# ?/?", [13] = "# ??/??", [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)", [39] = "#,##0.00;(#,##0.00)",
		[40] = "#,##0.00;[Red](#,##0.00)", [48] = "##0.0E+0", [49] = "@",
		[14] = "short date", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy", [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM",
		[20] = "h:mm", [21] = "h:mm:ss", [22] = "short date h:mm", [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mm:ss.0",
	};

	/// A number as Excel shows it with the format (id, and its code if it is the workbook's own), in this computer's
	/// culture: General as up to 15 significant digits; dates and times; the usual number formats (0, #,##0.00, %, E+00,
	/// sections for negative numbers and zero, text in quotes, currency symbols). Fractions and others: as General.
	public static string FormatNumber(double value, uint id, string? code, bool date1904 = false) {
		var culture = CultureInfo.CurrentCulture;
		code ??= BuiltIn.TryGetValue(id, out var builtIn) ? builtIn : null;
		string General() => value.ToString("G15", culture);
		if (code == null || code.Trim() == "" || code.Equals("General", StringComparison.OrdinalIgnoreCase) || code == "@") return General();
		if (IsDateCode(code)) {
			if (value < 0 || value > 2958465) return General();
			DateTime when;
			try { when = DateTime.FromOADate(value + (date1904 ? 1462 : 0)); } catch (ArgumentException) { return General(); }
			if (code.Contains("[h]", StringComparison.OrdinalIgnoreCase)) {    // (elapsed hours)
				var span = TimeSpan.FromDays(value);
				return ((int)span.TotalHours).ToString(culture) + culture.DateTimeFormat.TimeSeparator + span.ToString(@"mm\:ss", culture).Replace(":", culture.DateTimeFormat.TimeSeparator);
			}
			return FormatDate(when, id, code);
		}
		// the number sections (positive; negative; zero; text): the text section and Excel's own markup out
		var sections = SplitSections(code);
		if (sections.Count > 3) sections = sections.Take(3).ToList();
		var net = new List<string>();
		foreach (string section in sections) {
			if (Regex.IsMatch(section.Replace("\"", ""), @"\?")) return General();    // (fractions, aligned digits)
			net.Add(NetNumberFormat(section));
		}
		try {
			string text = value.ToString(string.Join(";", net), culture);
			return text.Trim();
		}
		catch (FormatException) { return General(); }
	}

	// Format sections split at ";" outside quotes and brackets.
	static List<string> SplitSections(string code) {
		var parts = new List<string>();
		var current = new StringBuilder();
		bool quoted = false, bracket = false;
		for (int i = 0; i < code.Length; i++) {
			char c = code[i];
			if (c == '\\' && !quoted && i + 1 < code.Length) { current.Append(c).Append(code[++i]); continue; }
			if (c == '"') quoted = !quoted;
			else if (!quoted && c == '[') bracket = true;
			else if (!quoted && c == ']') bracket = false;
			if (c == ';' && !quoted && !bracket) { parts.Add(current.ToString()); current.Clear(); continue; }
			current.Append(c);
		}
		parts.Add(current.ToString());
		return parts;
	}

	// One Excel number section as .NET's custom format: [$€-409] as its symbol, [Red] and conditions out, _x (room for a
	// character) as a space, *x (fill) out, quoted text and \x as they are (.NET reads them the same).
	static string NetNumberFormat(string section) {
		var sb = new StringBuilder();
		for (int i = 0; i < section.Length; i++) {
			char c = section[i];
			if (c == '"') {
				int end = section.IndexOf('"', i + 1);
				if (end < 0) end = section.Length - 1;
				sb.Append(section, i, end - i + 1);
				i = end;
			} else if (c == '\\' && i + 1 < section.Length) { sb.Append('\\').Append(section[++i]); }
			else if (c == '[') {
				int end = section.IndexOf(']', i);
				if (end < 0) break;
				string inside = section[(i + 1)..end];
				if (inside.StartsWith('$')) {
					string symbol = inside[1..].Split('-')[0];
					if (symbol != "") sb.Append('"').Append(symbol.Replace("\"", "")).Append('"');
				}
				i = end;
			} else if (c == '_' && i + 1 < section.Length) { sb.Append(' '); i++; }
			else if (c == '*' && i + 1 < section.Length) { i++; }
			else if (c is '@') { }
			else if (char.IsLetter(c) && c is not ('E' or 'e')) sb.Append('\\').Append(c);    // (letters are text in Excel)
			else sb.Append(c);
		}
		return sb.ToString();
	}

	// Whether a format code is a date or time: d, m, y, h or s outside quotes, brackets ([h] counts) and escapes.
	static bool IsDateCode(string code) {
		if (code is "short date" or "short date h:mm") return true;
		string bare = Regex.Replace(code, "\"[^\"]*\"|\\\\.|_.|\\*.", "");
		bare = Regex.Replace(bare, @"\[(?![hms]+\])[^\]]*\]", "", RegexOptions.IgnoreCase);
		bare = SplitSections(bare)[0];
		return Regex.IsMatch(bare, "[dmyhs]", RegexOptions.IgnoreCase) && !Regex.IsMatch(bare, "^General$", RegexOptions.IgnoreCase);
	}

	// A date or time as Excel's code shows it, translated into .NET's terms (m is minutes after h or before s, else the
	// month; AM/PM as tt); Excel's "short date" formats (ids 14 and 22) as this computer's.
	static string FormatDate(DateTime when, uint id, string? code) {
		var culture = CultureInfo.CurrentCulture;
		code ??= BuiltIn.TryGetValue(id, out var builtIn) ? builtIn : "short date";
		if (code == "short date") return when.ToString(culture.DateTimeFormat.ShortDatePattern, culture);
		if (code == "short date h:mm") return when.ToString(culture.DateTimeFormat.ShortDatePattern + " H:mm", culture);
		code = SplitSections(code)[0];
		bool ampm = Regex.IsMatch(code, "AM/PM|A/P", RegexOptions.IgnoreCase);
		// the tokens: letters runs, quoted text, escapes, and the rest
		var tokens = Regex.Matches(code, "\"[^\"]*\"|\\\\.|\\[[^\\]]*\\]|AM/PM|A/P|[yY]+|[mM]+|[dD]+|[hH]+|[sS]+|\\.0+|.", RegexOptions.IgnoreCase)
			.Select(m => m.Value).ToList();
		var sb = new StringBuilder();
		for (int i = 0; i < tokens.Count; i++) {
			string t = tokens[i];
			char first = char.ToLowerInvariant(t[0]);
			if (t.StartsWith('"')) sb.Append('\'').Append(t[1..^1].Replace("'", "\\'")).Append('\'');
			else if (t.StartsWith('\\')) sb.Append(t);
			else if (t.StartsWith('[')) { }    // (colors, locales)
			else if (t.Equals("AM/PM", StringComparison.OrdinalIgnoreCase)) sb.Append("tt");
			else if (t.Equals("A/P", StringComparison.OrdinalIgnoreCase)) sb.Append('t');
			else if (first == 'y') sb.Append(t.Length <= 2 ? "yy" : "yyyy");
			else if (first == 'd') sb.Append(t.Length switch { 1 => "%d", 2 => "dd", 3 => "ddd", _ => "dddd" });
			else if (first == 'h') sb.Append(ampm ? (t.Length == 1 ? "%h" : "hh") : (t.Length == 1 ? "%H" : "HH"));
			else if (first == 's') sb.Append(t.Length == 1 ? "%s" : "ss");
			else if (t.StartsWith('.') && t.Length > 1 && t.Skip(1).All(ch => ch == '0')) sb.Append('.').Append(new string('f', Math.Min(7, t.Length - 1)));
			else if (first == 'm') {
				bool minutes = t.Length <= 2 && (PreviousLetter(tokens, i) is 'h' || NextLetter(tokens, i) is 's');
				sb.Append(minutes ? (t.Length == 1 ? "%m" : "mm") : t.Length switch { 1 => "%M", 2 => "MM", 3 => "MMM", 5 => "MMM", _ => "MMMM" });
			}
			else if (char.IsLetter(t[0])) sb.Append('\\').Append(t);
			else if (t is "/" or ":") sb.Append('\\').Append(t);    // (Excel's are as written)
			else sb.Append(t);
		}
		try { return when.ToString(sb.ToString(), culture).Trim(); }    // (%d and the like: one specifier, anywhere)
		catch (FormatException) { return when.ToString(culture); }
	}

	static char? PreviousLetter(List<string> tokens, int i) {
		for (int k = i - 1; k >= 0; k--) if (char.IsLetter(tokens[k][0]) && !tokens[k].StartsWith('\\')) return char.ToLowerInvariant(tokens[k][0]);
		return null;
	}

	static char? NextLetter(List<string> tokens, int i) {
		for (int k = i + 1; k < tokens.Count; k++) if (char.IsLetter(tokens[k][0]) && !tokens[k].StartsWith('\\')) return char.ToLowerInvariant(tokens[k][0]);
		return null;
	}
}
