	// The experimental officeimo converter (iteration 47) against openxml, made-up text: one document with a heading,
	// bold, italic and colour, centred and right-aligned lines, double spacing, an ordinary list, a numbered style, legal
	// numbering, a page break, a table, a picture, a header and footer with page numbers, and line numbers. Saved as .docx
	// by each converter: what the file holds; opened again by the same converter: what came back; officeimo's report.
	// A comparison: it reports, and checks only that each converter writes and reads a .docx at all.
	public string SelfTest(Action<int> pump) {
		var r = new StringBuilder();
		void Check(string what, object actual, object expected) => r.AppendLine((Equals(actual, expected) ? "ok   " : "FAIL ") + what + ": " + actual + (Equals(actual, expected) ? "" : " (expected " + expected + ")"));
		var doc = Editor.TextDocument;
		string folder = Path.GetDirectoryName(Environment.GetEnvironmentVariable("RT_OUT")!)!;
		string Story() => doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		int Find(string text) => Story().IndexOf(text, StringComparison.Ordinal);
		void Select(string from, string to) { int a = Find(from), b = Find(to); var p = doc.Range2(b, b); p.Expand(4); api.Select(a, p.End - 1 - a); }
		string picture = Path.Combine(folder, "made-up.png");
		using (var bmp = new Bitmap(90, 60)) {
			using (var g = Graphics.FromImage(bmp)) { g.Clear(Color.SteelBlue); g.FillEllipse(Brushes.Orange, 15, 10, 60, 40); }
			bmp.Save(picture);
		}
		// the document
		Run(() => api.View = "Rich");
		Run(() => { api.PrintLayout = true; api.PageView = false; });
		Run(() => api.New());
		var rtf = new StringBuilder(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}{\colortbl;\red200\green0\blue0;}\f0\fs24 ");
		foreach (var paragraph in new[] { "Made-up Title", @"A plain paragraph with {\b bold}, {\i italic} and {\cf1 red} words.", @"\qc Centred line.", @"\qr Right-aligned line.",
			@"\sl480\slmult1 A double-spaced paragraph of made-up text, long enough to wrap onto a second line of the page so that its spacing shows.",
			"Alpha item.", "Bravo item.", "Charlie item.", "Delta item.", "Echo item.", "Foxtrot item.", "Golf after a page break.", "Closing paragraph." })
			rtf.Append(@"\pard " + paragraph + @"\par ");
		rtf.Append("}");
		Run(() => api.Rtf = rtf.ToString());
		Run(() => { api.Select(Find("Made-up Title"), 0); api.SetStyle("Heading 1"); });
		Run(() => { Select("Alpha", "Bravo"); api.NumberedList(); });
		Run(() => { Select("Charlie", "Delta"); api.SetStyle("Numbered Paragraph"); });
		Run(() => { Select("Echo", "Foxtrot"); api.LegalNumbering(); });
		Run(() => { api.Select(Find("Golf after"), 0); api.InsertPageBreak(); });
		Run(() => { api.Select(Find("Closing paragraph"), 0); api.InsertTable(2, 2); });
		Run(() => { api.Select(Find("Closing paragraph"), 0); api.InsertImageFile(picture, "a made-up picture"); });
		Run(() => { api.SetHeader("Made-up header", "", ""); api.ShowHeader = true; api.SetFooter("", "Page {page}", ""); api.ShowFooter = true; });
		Run(() => api.LineNumbers = "continuous");
		pump(600);
		// what the editor has (the original), and again after a round trip
		string Shape() {
			var parts = new List<string>();
			string story = Story();
			foreach (var p in StyleEngine.Paragraphs(doc, 0, doc.Range2(0, 0).StoryLength)) {
				var range = doc.Range2(p.Start, p.End);
				string shown = (range.GetText2(tomConstants.tomIncludeNumbering) ?? "").TrimEnd('\r').Replace("\t", "→");
				if (shown.Trim('\u000C', '￼', ' ') == "") continue;
				var para = range.Para;
				string style = StyleEngine.StyleAt(doc, Styles, p.Start).Name;
				string align = para.Alignment switch { 1 => " centre", 2 => " right", _ => "" };
				string spacing = para.LineSpacingRule == 5 && Math.Abs(para.LineSpacing - 2f) < 0.05f ? " double" : "";
				parts.Add($"{shown.Substring(0, Math.Min(18, shown.Length))}{(style is "Normal" or "" ? "" : $" ({style})")}{align}{spacing}");
			}
			int bold = Find("bold"), italic = Find("italic"), red = Find("red");
			string inline = $"bold {doc.Range2(bold, bold + 4).Font.Bold == -1}, italic {doc.Range2(italic, italic + 6).Font.Italic == -1}, red {(bold >= 0 && red >= 0 ? (doc.Range2(red, red + 3).Font.ForeColor & 0xFF) > 150 : false)}";
			return $"{string.Join(" | ", parts)}\n       {inline}; tables {story.Count(c => c == '￹')}, pictures {story.Count(c => c == '￼')}, page breaks {PageBreaks.Positions(doc).Count}; " +
				$"header {string.Join("|", PageBands.Parts(Bands.Header))} ({Bands.ShowHeader}), footer {string.Join("|", PageBands.Parts(Bands.Footer))} ({Bands.ShowFooter}); line numbers {api.LineNumbers}";
		}
		string original = Shape();
		string originalFile = Path.Combine(folder, "made-up original.rtf");
		Run(() => api.SaveAs(originalFile));
		r.AppendLine("     original:  " + original.Replace("\n", "\n  "));
		string Docx(string file) {
			using var zip = System.IO.Compression.ZipFile.OpenRead(file);
			string Part(string name) => zip.GetEntry(name) is { } e ? new StreamReader(e.Open()).ReadToEnd() : "";
			string body = Part("word/document.xml"), numbering = Part("word/numbering.xml");
			string footers = string.Concat(zip.Entries.Where(e => e.FullName.StartsWith("word/footer")).Select(e => new StreamReader(e.Open()).ReadToEnd()));
			string headers = string.Concat(zip.Entries.Where(e => e.FullName.StartsWith("word/header")).Select(e => new StreamReader(e.Open()).ReadToEnd()));
			int C(string text, string pattern) => Regex.Matches(text, pattern).Count;
			return $"paragraphs {C(body, "<w:p[ >]")}, numbered paragraphs {C(body, "<w:numPr>")} (numbering definitions {C(numbering, "<w:abstractNum ")}), tables {C(body, "<w:tbl>")}, pictures {C(body, "<w:drawing>|<w:pict>")}, " +
				$"centred {C(body, @"<w:jc w:val=""center""")}, right {C(body, @"<w:jc w:val=""(right|end)""")}, bold {C(body, "<w:b/>|<w:b ")}, colour {C(body, "<w:color ")}, double {C(body, @"w:line=""480""")}, page breaks {C(body, @"w:type=""page""|<w:pageBreakBefore")}, " +
				$"heading style {C(body, @"w:pStyle w:val=""Heading1""")}; header parts {zip.Entries.Count(e => e.FullName.StartsWith("word/header"))} ({(headers.Contains("Made-up header") ? "with its text" : "no text")}), " +
				$"footer parts {zip.Entries.Count(e => e.FullName.StartsWith("word/footer"))} (PAGE fields {C(footers, @"PAGE\b")}), line numbering {C(body, "<w:lnNumType")}";
		}
		foreach (var converter in new[] { "openxml", "officeimo" }) {
			r.AppendLine($"     (from the original again: the same {Shape() == original})");
			string file = Path.Combine(folder, $"made-up {converter}.docx");
			string written = "";
			try { Run(() => written = api.SaveAs(file, converter)); }
			catch (Exception ex) { r.AppendLine($"     {converter}: saving failed: {ex.GetType().Name}: {ex.Message.Split('\n')[0]}"); }
			Check($"{converter}: saved a .docx", File.Exists(file) && api.LastConverter == converter, true);
			if (!File.Exists(file)) continue;
			r.AppendLine($"     {converter} wrote: {Docx(file)}");
			if (converter == "officeimo") r.AppendLine($"     {converter} report, writing: {(api.ConverterReport() is { Length: > 0 } w ? w.Replace("\n", "\n         ") : "nothing lost")}");
			try { Run(() => api.Open(file, converter)); }
			catch (Exception ex) { r.AppendLine($"     {converter}: opening failed: {ex.GetType().Name}: {ex.Message.Split('\n')[0]}"); continue; }
			pump(600);
			Check($"{converter}: opened it again", api.LastConverter, converter);
			string back = Shape();
			r.AppendLine($"     {converter} read:  " + back.Replace("\n", "\n  "));
			r.AppendLine($"     {converter}: the same as the original: {back == original}");
			if (converter == "officeimo") r.AppendLine($"     {converter} report, reading: {(api.ConverterReport() is { Length: > 0 } rd ? rd.Replace("\n", "\n         ") : "nothing lost")}");
			// (back to the original for the next converter)
			Run(() => api.Open(originalFile));
			pump(400);
		}
		// the original reopened from its .rtf: what differs from the original as built
		Run(() => api.Open(originalFile));
		pump(400);
		r.AppendLine($"     the original reopened from .rtf: {Shape().Replace("\n", "\n  ")}");
		// cross reads: writing and reading told apart
		foreach (var (file, reader) in new[] { ("made-up officeimo.docx", "openxml"), ("made-up openxml.docx", "officeimo") }) {
			try { Run(() => api.Open(Path.Combine(folder, file), reader)); }
			catch (Exception ex) { r.AppendLine($"     {file} read by {reader}: failed: {ex.Message.Split('\n')[0]}"); continue; }
			pump(600);
			r.AppendLine($"     {file} read by {reader}: {Shape().Replace("\n", "\n  ")}");
			if (reader == "officeimo") r.AppendLine($"     officeimo report, reading openxml's file: {(api.ConverterReport() is { Length: > 0 } rd ? rd.Replace("\n", "\n         ") : "nothing lost")}");
		}
		Editor.Modified = false;
		return r.ToString();
	}
	public void Capture(string path) { }
