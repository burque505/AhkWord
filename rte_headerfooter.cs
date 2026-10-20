// Headers and footers of a document, for AhkWord (AhkWord.cs, through #load; uses the Open XML SDK packages
// that rte_office5.cs lists).
// A document's header and footer each have a left, a center and a right part, with the fields {title} {file} {page}
// {pages} {date} {time} (PageBands). They are printed, and kept in the files:
//   RTF            real \header / \footer groups (one line: left <tab> center <tab> right, at center and right tab
//                  stops), with PAGE, NUMPAGES, DATE and TIME fields; the page size and margins of Page Setup; and the
//                  exact parts in {\*\rtewriterheaderfooter ...}, which other programs ignore
//   Word, OpenDocument
//                  the same, made by Word or LibreOffice from that RTF, or by the built-in converter (header and footer
//                  parts); the exact parts in a custom XML part (.docx) or a user-defined property (.odt)
//   HTML           <meta name="rte-writer-header-footer" content="...">
//   Markdown       <!-- rte-writer-header-footer: ... --> as the last line
// Read back: the exact parts if the file has them and its real header and footer are still the ones AhkWord wrote
// (if they were edited in Word or LibreOffice, those win); otherwise the real header and footer, with Word's and
// LibreOffice's page number, page count, date and time fields as {page} {pages} {date} {time}.
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

/// A document's header and footer: whether each is shown, and its left, center and right parts.
public sealed class PageBands
{
	public bool ShowHeader { get; set; }
	public string[] Header { get; set; } = { "", "", "" };
	public bool ShowFooter { get; set; }
	public string[] Footer { get; set; } = { "", "", "" };
	/// The parts as they were written into the file's real header and footer ({title} and {file} filled in), to tell
	/// whether another program changed them since.
	public string[]? WrittenHeader { get; set; }
	public string[]? WrittenFooter { get; set; }

	public static string[] Parts(string[]? parts) => Enumerable.Range(0, 3).Select(i => parts != null && i < parts.Length ? parts[i] ?? "" : "").ToArray();

	public PageBands Copy() => new() { ShowHeader = ShowHeader, Header = Parts(Header), ShowFooter = ShowFooter, Footer = Parts(Footer) };

	public bool SameAs(PageBands other) => ShowHeader == other.ShowHeader && ShowFooter == other.ShowFooter
		&& Parts(Header).SequenceEqual(Parts(other.Header)) && Parts(Footer).SequenceEqual(Parts(other.Footer));

	bool HeaderEmpty => !ShowHeader || Parts(Header).All(p => p.Trim() == "");
	bool FooterEmpty => !ShowFooter || Parts(Footer).All(p => p.Trim() == "");
	/// The parts that go into a file's real header (null: no header there).
	[System.Text.Json.Serialization.JsonIgnore] public string[]? HeaderToWrite => HeaderEmpty ? null : Parts(Header);
	[System.Text.Json.Serialization.JsonIgnore] public string[]? FooterToWrite => FooterEmpty ? null : Parts(Footer);

	static readonly JsonSerializerOptions Json = new() { DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull };
	public string ToJson() => JsonSerializer.Serialize(this, Json);
	public static PageBands? FromJson(string json) {
		try { return JsonSerializer.Deserialize<PageBands>(json) is { } b ? Normalized(b) : null; }
		catch (JsonException) { return null; }
	}
	static PageBands Normalized(PageBands b) { b.Header = Parts(b.Header); b.Footer = Parts(b.Footer); return b; }
}

/// The page a document is laid out on in other programs (twips): Page Setup's paper and margins.
public sealed record PageLayout(int Width, int Height, int Left, int Right, int Top, int Bottom, bool Landscape)
{
	public int TextWidth => Math.Max(1440, Width - Left - Right);
}

public static class HeaderFooter
{
	const string RtfDestination = "rtewriterheaderfooter";
	const string MetaName = "rte-writer-header-footer";
	const string XmlNamespace = "urn:rte-writer:header-footer";
	const string OdtProperty = "AhkWord header and footer";
	static readonly Regex Field = new(@"\{(title|file|page|pages|date|time)\}", RegexOptions.IgnoreCase);

	// {title} and {file} become text; the others stay fields (filled in by Word and LibreOffice).
	static string[] Filled(string[] parts, string title, string file) =>
		parts.Select(p => Regex.Replace(p, @"\{(title|file)\}", m => m.Groups[1].Value.Equals("title", StringComparison.OrdinalIgnoreCase) ? title : file, RegexOptions.IgnoreCase)).ToArray();

	/// The bands as written into a file (with WrittenHeader / WrittenFooter), for the hidden exact copy.
	static PageBands Stamped(PageBands bands, string title, string file) {
		var b = bands.Copy();
		b.WrittenHeader = bands.HeaderToWrite is { } h ? Filled(h, title, file) : null;
		b.WrittenFooter = bands.FooterToWrite is { } f ? Filled(f, title, file) : null;
		return b;
	}

	// What was read from a file's real header and footer (null: none), against the exact copy (null: none).
	static PageBands? Decide(PageBands? exact, string[]? header, string[]? footer, bool haveReal) {
		if (exact != null) {
			bool same = !haveReal || (Same(exact.WrittenHeader, header) && Same(exact.WrittenFooter, footer));
			if (same) { exact.WrittenHeader = exact.WrittenFooter = null; return exact; }
		}
		if (!haveReal) return null;
		return new PageBands {
			ShowHeader = header != null, Header = PageBands.Parts(header),
			ShowFooter = footer != null, Footer = PageBands.Parts(footer)
		};
	}

	static bool Same(string[]? a, string[]? b) =>
		(a == null || a.All(p => p.Trim() == "")) && (b == null || b.All(p => p.Trim() == ""))
		|| (a != null && b != null && PageBands.Parts(a).Select(Clean).SequenceEqual(PageBands.Parts(b).Select(Clean)));

	static string Clean(string s) => Regex.Replace(s.Replace('\u00A0', ' '), @"\s+", " ").Trim();

	// A line of header text (tabs between the parts) as left, center and right; one part alone goes where its
	// paragraph is aligned.
	static string[]? PartsOf(string line, int align) {
		line = line.Trim('\r', '\n');
		if (Clean(line.Replace('\t', ' ')) == "") return null;
		var pieces = line.Split('\t');
		if (pieces.Length == 1) return align switch { 1 => new[] { "", pieces[0].Trim(), "" }, 2 => new[] { "", "", pieces[0].Trim() }, _ => new[] { pieces[0].Trim(), "", "" } };
		if (pieces.Length == 2) return new[] { pieces[0].Trim(), pieces[1].Trim(), "" };
		return new[] { pieces[0].Trim(), pieces[1].Trim(), string.Join(" ", pieces.Skip(2)).Trim() };
	}

	static string FieldOf(string instruction, string result) {
		var m = Regex.Match(instruction, @"^\s*([A-Z]+)(.*)$", RegexOptions.IgnoreCase);
		string name = m.Success ? m.Groups[1].Value.ToUpperInvariant() : "";
		return name switch {
			"PAGE" => "{page}",
			"NUMPAGES" or "SECTIONPAGES" => "{pages}",
			"DATE" or "CREATEDATE" or "SAVEDATE" or "PRINTDATE" => Regex.IsMatch(m.Groups[2].Value, @"[Hhm]{1,2}:") && !Regex.IsMatch(m.Groups[2].Value, "[dMy]") ? "{time}" : "{date}",
			"TIME" => "{time}",
			"FILENAME" => m.Groups[2].Value.Contains("\\p", StringComparison.OrdinalIgnoreCase) ? "{file}" : "{title}",
			_ => result
		};
	}

	// ================================================================ RTF

	/// RTF with the header and footer (and the page layout) added, for an .rtf file, Word or LibreOffice. Added in front
	/// of the first paragraph, after any other changes to the RTF.
	public static string AddToRtf(string rtf, PageBands bands, PageLayout layout, string title, string file, float fontSize) {
		int body = BodyStart(rtf);
		if (body < 0) return rtf;
		var sb = new StringBuilder();
		sb.Append($@"\paperw{layout.Width}\paperh{layout.Height}\margl{layout.Left}\margr{layout.Right}\margt{layout.Top}\margb{layout.Bottom}\headery720\footery720");
		if (layout.Landscape) sb.Append(@"\landscape");
		sb.Append(@"\sectd");
		if (layout.Landscape) sb.Append(@"\lndscpsxn");
		sb.Append($@"\pgwsxn{layout.Width}\pghsxn{layout.Height}\marglsxn{layout.Left}\margrsxn{layout.Right}\margtsxn{layout.Top}\margbsxn{layout.Bottom}\headery720\footery720 ");
		int size = (int)Math.Round(fontSize * 2);
		void Band(string destination, string[]? parts) {
			if (parts == null) return;
			sb.Append("{\\").Append(destination).Append($@"\pard\plain\f0\fs{size}\tqc\tx{layout.TextWidth / 2}\tqr\tx{layout.TextWidth} ");
			sb.Append(string.Join(@"\tab ", Filled(parts, title, file).Select(RtfText)));
			sb.Append(@"\par}");
		}
		Band("header", bands.HeaderToWrite);
		Band("footer", bands.FooterToWrite);
		sb.Append(@"{\*\").Append(RtfDestination).Append(' ').Append(RtfEscape(Stamped(bands, title, file).ToJson())).Append('}');
		return rtf.Insert(body, sb.ToString());
	}

	// A part as RTF: text, with fields for {page} {pages} {date} {time}.
	static string RtfText(string part) {
		var sb = new StringBuilder();
		int at = 0;
		foreach (Match m in Field.Matches(part)) {
			sb.Append(RtfEscape(part[at..m.Index]));
			string instruction = m.Groups[1].Value.ToLowerInvariant() switch { "page" => "PAGE", "pages" => "NUMPAGES", "date" => "DATE", _ => "TIME" };
			string shown = instruction switch { "PAGE" or "NUMPAGES" => "1", "DATE" => DateTime.Now.ToShortDateString(), _ => DateTime.Now.ToShortTimeString() };
			sb.Append(@"{\field{\*\fldinst ").Append(instruction).Append(@"}{\fldrslt ").Append(RtfEscape(shown)).Append("}}");
			at = m.Index + m.Length;
		}
		sb.Append(RtfEscape(part[at..]));
		return sb.ToString();
	}

	static string RtfEscape(string text) {
		var sb = new StringBuilder();
		foreach (char c in text) {
			if (c is '\\' or '{' or '}') sb.Append('\\').Append(c);
			else if (c > 126 || c < 32) sb.Append(@"\u").Append((short)c).Append('?');
			else sb.Append(c);
		}
		return sb.ToString();
	}

	// Where the document's text starts: its first \pard (or \sectd) outside groups other than the outermost.
	static int BodyStart(string rtf) {
		int depth = 0;
		for (int i = 0; i < rtf.Length; i++) {
			char c = rtf[i];
			if (c == '\\') {
				if (depth == 1 && (string.CompareOrdinal(rtf, i, @"\pard", 0, 5) == 0 || string.CompareOrdinal(rtf, i, @"\sectd", 0, 6) == 0)
					&& (i + 6 > rtf.Length || !char.IsLetter(rtf[i + (rtf[i + 1] == 'p' ? 5 : 6)]))) return i;
				i++;
			} else if (c == '{') depth++;
			else if (c == '}') depth--;
		}
		return -1;
	}

	/// The header and footer of RTF (null if it has neither): AhkWord's exact copy, or the RTF's header and footer.
	public static PageBands? FromRtf(string rtf) {
		PageBands? exact = null;
		var copy = Regex.Match(rtf, @"\{\\\*\\" + RtfDestination + @" ((?:[^\\{}]|\\[\\{}]|\\u-?\d+\?)*)\}");
		if (copy.Success) exact = PageBands.FromJson(RtfUnescape(copy.Groups[1].Value));
		string[]? header = null, footer = null;
		bool any = false;
		foreach (var (name, isHeader) in new[] { ("header", true), ("headerr", true), ("footer", false), ("footerr", false) }) {
			if ((isHeader ? header : footer) != null) continue;
			int start = FindGroup(rtf, name);
			if (start < 0) continue;
			any = true;
			var (text, align) = RtfGroupText(rtf, start);
			var parts = PartsOf(FirstLine(text), align);
			if (isHeader) header = parts; else footer = parts;
		}
		return Decide(exact, header, footer, any);
	}

	// Header text over several paragraphs: the first one with text (the editor's header is one line).
	static string FirstLine(string text) => text.Split('\n').FirstOrDefault(l => Clean(l.Replace('\t', ' ')) != "") ?? "";

	static string RtfUnescape(string s) {
		s = Regex.Replace(s, @"\\u(-?\d+)\?", m => ((char)(short)int.Parse(m.Groups[1].Value)).ToString());
		return Regex.Replace(s, @"\\([\\{}])", "$1");
	}

	/// RTF without its header and footer groups (and AhkWord's copy), which RichEdit has no place for.
	public static string StripFromRtf(string rtf) {
		foreach (string name in new[] { "header", "headerl", "headerr", "headerf", "footer", "footerl", "footerr", "footerf", "*\\" + RtfDestination }) {
			while (true) {
				int start = FindGroup(rtf, name);
				if (start < 0) break;
				int end = GroupEnd(rtf, start);
				if (end < 0) break;
				rtf = rtf.Remove(start, end - start + 1);
			}
		}
		return rtf;
	}

	// The position of "{\name" (not a longer control word), or -1.
	static int FindGroup(string rtf, string name) {
		string open = "{\\" + name;
		for (int i = rtf.IndexOf(open, StringComparison.Ordinal); i >= 0; i = rtf.IndexOf(open, i + 1, StringComparison.Ordinal)) {
			int after = i + open.Length;
			if (after >= rtf.Length || !char.IsLetter(rtf[after])) {
				if (i > 0 && rtf[i - 1] == '\\') continue;    // an escaped brace
				return i;
			}
		}
		return -1;
	}

	static int GroupEnd(string rtf, int start) {
		int depth = 0;
		for (int i = start; i < rtf.Length; i++) {
			char c = rtf[i];
			if (c == '\\') { i++; continue; }
			if (c == '{') depth++;
			else if (c == '}' && --depth == 0) return i;
		}
		return -1;
	}

	static readonly Regex RtfToken = new(@"\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z])|([{}])|([^\\{}\r\n]+)|[\r\n]+");
	static readonly Encoding Latin = Windows1252();

	static Encoding Windows1252() {
		try {
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
			return Encoding.GetEncoding(1252);
		} catch (Exception) {
			return Encoding.Latin1;
		}
	}

	// The text of an RTF group (tabs, paragraph ends as \n, fields as {page} and so on), and the alignment of its first
	// paragraph (0 left, 1 center, 2 right).
	static (string Text, int Align) RtfGroupText(string rtf, int start) {
		int end = GroupEnd(rtf, start);
		var tokens = RtfToken.Matches(rtf[start..(end + 1)]).ToList();
		int align = -1, i = 0;
		string text = Group(tokens, ref i, ref align);
		return (text, Math.Max(0, align));
	}

	static readonly HashSet<string> Skipped = new() { "fonttbl", "colortbl", "stylesheet", "pict", "object", "info", "listtable", "listoverridetable", "shp", "nonshppict" };

	// Reads the group whose "{" is tokens[i]; returns its text and leaves i after its "}".
	static string Group(List<Match> tokens, ref int i, ref int align) {
		i++;    // "{"
		var sb = new StringBuilder();
		bool first = true;
		int skipChars = 0;
		while (i < tokens.Count) {
			var t = tokens[i];
			if (t.Groups[5].Success) {
				if (t.Value == "}") { i++; return sb.ToString(); }
				// a group inside: a field, a skipped destination, or more text
				var next = i + 1 < tokens.Count ? tokens[i + 1] : null;
				string word = next?.Groups[1].Value ?? "";
				bool starred = next?.Groups[4].Value == "*";
				if (word == "field") { sb.Append(FieldGroup(tokens, ref i, ref align)); continue; }
				if (starred || Skipped.Contains(word)) { SkipGroup(tokens, ref i); continue; }
				sb.Append(Group(tokens, ref i, ref align));
				continue;
			}
			i++;
			if (t.Groups[1].Success) {
				string word = t.Groups[1].Value;
				switch (word) {
					case "tab": sb.Append('\t'); break;
					case "par": case "sect": sb.Append('\n'); break;
					case "line": sb.Append(' '); break;
					case "u":
						sb.Append((char)(short)int.Parse(t.Groups[2].Value));
						skipChars = 1;
						continue;
					case "qc": if (align < 0) align = 1; break;
					case "qr": if (align < 0) align = 2; break;
					case "ql": case "qj": if (align < 0) align = 0; break;
				}
				if (first && Skipped.Contains(word)) { /* (handled by the caller) */ }
			} else if (t.Groups[3].Success) {
				if (skipChars > 0) { skipChars--; continue; }
				sb.Append(Latin.GetString(new[] { Convert.ToByte(t.Groups[3].Value, 16) }));
			} else if (t.Groups[4].Success) {
				string symbol = t.Groups[4].Value;
				if (symbol == "~") sb.Append('\u00A0');
				else if (symbol is "\\" or "{" or "}") sb.Append(symbol);
			} else if (t.Groups[6].Success) {
				string s = t.Value;
				if (skipChars > 0 && s.Length > 0) { s = s[1..]; skipChars = 0; }
				sb.Append(s);
			}
			first = false;
			skipChars = t.Groups[6].Success || t.Groups[3].Success ? 0 : skipChars;
		}
		return sb.ToString();
	}

	static void SkipGroup(List<Match> tokens, ref int i) {
		int depth = 0;
		for (; i < tokens.Count; i++) {
			if (tokens[i].Value == "{") depth++;
			else if (tokens[i].Value == "}" && --depth == 0) { i++; return; }
		}
	}

	// {\field{\*\fldinst PAGE}{\fldrslt 1}}: the field's {name}, or its result for other fields.
	static string FieldGroup(List<Match> tokens, ref int i, ref int align) {
		int depth = 0;
		string instruction = "", result = "";
		i++;    // "{"
		depth = 1;
		while (i < tokens.Count && depth > 0) {
			var t = tokens[i];
			if (t.Value == "{") {
				// {\*\fldinst ...} or {\fldrslt ...}
				int j = i + 1;
				if (j < tokens.Count && tokens[j].Groups[4].Value == "*") j++;
				string word = j < tokens.Count ? tokens[j].Groups[1].Value : "";
				if (word == "fldinst") {
					int a = -1;
					int k = i;
					// the instruction's text: the group read as text, without its "\*\fldinst"
					instruction = PlainGroup(tokens, ref k);
					i = k;
					continue;
				}
				if (word == "fldrslt") { result = Group(tokens, ref i, ref align); continue; }
				SkipGroup(tokens, ref i);
				continue;
			}
			if (t.Value == "}") { depth--; i++; continue; }
			i++;
		}
		return FieldOf(instruction, result);
	}

	// A group's plain text (all of it, starred or not).
	static string PlainGroup(List<Match> tokens, ref int i) {
		var sb = new StringBuilder();
		int depth = 0;
		for (; i < tokens.Count; i++) {
			var t = tokens[i];
			if (t.Value == "{") depth++;
			else if (t.Value == "}") { if (--depth == 0) { i++; break; } }
			else if (t.Groups[6].Success) sb.Append(t.Value);
			else if (t.Groups[3].Success) sb.Append(Latin.GetString(new[] { Convert.ToByte(t.Groups[3].Value, 16) }));
		}
		return sb.ToString();
	}

	// ================================================================ .docx

	/// Adds the header and footer to a .docx: header and footer parts and the page layout if it has no header or footer
	/// yet (the built-in converter's), and the exact copy in either case.
	public static void AddToDocx(string path, PageBands bands, PageLayout layout, string title, string file, bool parts) {
		using var package = WordprocessingDocument.Open(path, true);
		var main = package.MainDocumentPart!;
		var body = main.Document.Body!;
		if (parts && !main.HeaderParts.Any() && !main.FooterParts.Any()) {
			var section = body.Elements<W.SectionProperties>().LastOrDefault();
			if (section == null) { section = new W.SectionProperties(); body.Append(section); }
			foreach (var old in section.Elements<W.HeaderReference>().Cast<DocumentFormat.OpenXml.OpenXmlElement>().Concat(section.Elements<W.FooterReference>()).ToList()) old.Remove();
			DocumentFormat.OpenXml.OpenXmlElement? last = null;
			if (bands.HeaderToWrite is { } header) {
				var part = main.AddNewPart<HeaderPart>();
				part.Header = new W.Header(DocxParagraph(Filled(header, title, file), layout, "Header"));
				part.Header.Save();
				last = section.PrependChild(new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(part) });
			}
			if (bands.FooterToWrite is { } footer) {
				var part = main.AddNewPart<FooterPart>();
				part.Footer = new W.Footer(DocxParagraph(Filled(footer, title, file), layout, "Footer"));
				part.Footer.Save();
				var reference = new W.FooterReference { Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(part) };
				last = last == null ? section.PrependChild(reference) : section.InsertAfter(reference, last);
			}
			section.RemoveAllChildren<W.PageSize>();
			section.RemoveAllChildren<W.PageMargin>();
			var size = new W.PageSize { Width = (uint)layout.Width, Height = (uint)layout.Height };
			if (layout.Landscape) size.Orient = W.PageOrientationValues.Landscape;
			var margin = new W.PageMargin { Left = (uint)layout.Left, Right = (uint)layout.Right, Top = layout.Top, Bottom = layout.Bottom, Header = 720, Footer = 720, Gutter = 0 };
			if (last == null) { section.PrependChild(margin); section.PrependChild(size); }
			else { section.InsertAfter(margin, last); section.InsertAfter(size, last); }
			main.Document.Save();
		}
		// the exact copy, in a custom XML part of its own (Word keeps those)
		foreach (var old in main.CustomXmlParts.Where(p => ReadCustomXml(p) != null).ToList()) main.DeletePart(old);
		var custom = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
		using (var stream = custom.GetStream(FileMode.Create, FileAccess.Write))
			new XDocument(new XElement(XName.Get("headerFooter", XmlNamespace), Stamped(bands, title, file).ToJson())).Save(stream);
	}

	static string? ReadCustomXml(CustomXmlPart part) {
		try {
			using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
			var root = XDocument.Load(stream).Root;
			return root?.Name == XName.Get("headerFooter", XmlNamespace) ? root.Value : null;
		} catch { return null; }
	}

	static W.Paragraph DocxParagraph(string[] parts, PageLayout layout, string style) {
		var paragraph = new W.Paragraph(new W.ParagraphProperties(
			new W.Tabs(new W.TabStop { Val = W.TabStopValues.Center, Position = layout.TextWidth / 2 }, new W.TabStop { Val = W.TabStopValues.Right, Position = layout.TextWidth })));
		for (int p = 0; p < 3; p++) {
			if (p > 0) paragraph.Append(new W.Run(new W.TabChar()));
			int at = 0;
			string part = parts[p];
			foreach (Match m in Field.Matches(part)) {
				if (m.Index > at) paragraph.Append(new W.Run(new W.Text(part[at..m.Index]) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
				string instruction = m.Groups[1].Value.ToLowerInvariant() switch { "page" => "PAGE", "pages" => "NUMPAGES", "date" => "DATE", _ => "TIME" };
				paragraph.Append(new W.SimpleField(new W.Run(new W.Text("1"))) { Instruction = " " + instruction + " " });
				at = m.Index + m.Length;
			}
			if (at < part.Length) paragraph.Append(new W.Run(new W.Text(part[at..]) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
		}
		return paragraph;
	}

	/// The header and footer of a .docx (null if it has neither).
	public static PageBands? FromDocx(string path) {
		try {
			using var package = WordprocessingDocument.Open(path, false);
			var main = package.MainDocumentPart!;
			PageBands? exact = null;
			foreach (var custom in main.CustomXmlParts)
				if (ReadCustomXml(custom) is { } json) exact = PageBands.FromJson(json);
			var section = main.Document.Body?.Elements<W.SectionProperties>().LastOrDefault()
				?? main.Document.Body?.Descendants<W.SectionProperties>().FirstOrDefault();
			string[]? Read<TRef>(Func<string, OpenXmlPart?> part) where TRef : W.HeaderFooterReferenceType {
				var references = section?.Elements<TRef>().ToList() ?? new List<TRef>();
				var reference = references.FirstOrDefault(r => r.Type == null || r.Type.Value == W.HeaderFooterValues.Default) ?? references.FirstOrDefault();
				if (reference?.Id?.Value is not string id || part(id) is not { } found) return null;
				var root = found.RootElement;
				if (root == null) return null;
				foreach (var paragraph in root.Descendants<W.Paragraph>()) {
					var (line, align) = DocxText(paragraph);
					if (PartsOf(line, align) is { } parts) return parts;
				}
				return null;
			}
			bool haveReal = section != null && (section.Elements<W.HeaderReference>().Any() || section.Elements<W.FooterReference>().Any());
			var header = Read<W.HeaderReference>(id => main.GetPartById(id));
			var footer = Read<W.FooterReference>(id => main.GetPartById(id));
			return Decide(exact, header, footer, haveReal);
		} catch (Exception ex) when (ex is IOException or InvalidDataException or DocumentFormat.OpenXml.Packaging.OpenXmlPackageException) {
			return null;
		}
	}

	// A paragraph's text, with tabs and fields ({page} ...), and its alignment.
	static (string Text, int Align) DocxText(W.Paragraph paragraph) {
		var sb = new StringBuilder();
		string? instruction = null;    // in a complex field: its instruction so far
		bool inResult = false;
		string result = "";
		foreach (var element in paragraph.Descendants()) {
			switch (element) {
				case W.SimpleField simple:
					sb.Append(FieldOf(simple.Instruction?.Value ?? "", simple.InnerText));
					break;
				case W.FieldChar fieldChar when fieldChar.FieldCharType?.Value == W.FieldCharValues.Begin:
					instruction = ""; inResult = false; result = "";
					break;
				case W.FieldChar fieldChar when fieldChar.FieldCharType?.Value == W.FieldCharValues.Separate:
					inResult = true;
					break;
				case W.FieldChar fieldChar when fieldChar.FieldCharType?.Value == W.FieldCharValues.End:
					if (instruction != null) sb.Append(FieldOf(instruction, result));
					instruction = null; inResult = false;
					break;
				case W.FieldCode code:
					if (instruction != null) instruction += code.Text;
					break;
				case W.Text text when text.Ancestors<W.SimpleField>().Any():
					break;
				case W.Text text:
					if (instruction != null) { if (inResult) result += text.Text; }
					else sb.Append(text.Text);
					break;
				case W.TabChar when element.Parent is W.Run:
					if (instruction == null) sb.Append('\t');
					break;
				case W.PositionalTab:
					sb.Append('\t');
					break;
			}
		}
		var justification = paragraph.ParagraphProperties?.Justification?.Val?.Value;
		int align = justification == W.JustificationValues.Center ? 1 : justification == W.JustificationValues.Right || justification == W.JustificationValues.End ? 2 : 0;
		return (sb.ToString(), align);
	}

	// ================================================================ .odt

	static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", Meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
		Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0", Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

	/// Adds the exact copy to an .odt (its real header and footer come from the RTF LibreOffice or Word converted).
	public static void AddToOdt(string path, PageBands bands, string title, string file) {
		using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
		var entry = zip.GetEntry("meta.xml");
		XDocument document;
		if (entry != null) { using (var stream = entry.Open()) document = XDocument.Load(stream); entry.Delete(); }
		else document = new XDocument(new XElement(Office + "document-meta", new XAttribute(XNamespace.Xmlns + "office", Office), new XAttribute(XNamespace.Xmlns + "meta", Meta),
			new XAttribute(Office + "version", "1.3"), new XElement(Office + "meta")));
		var meta = document.Root!.Element(Office + "meta");
		if (meta == null) { meta = new XElement(Office + "meta"); document.Root.Add(meta); }
		meta.Elements(Meta + "user-defined").Where(e => (string?)e.Attribute(Meta + "name") == OdtProperty).Remove();
		meta.Add(new XElement(Meta + "user-defined", new XAttribute(Meta + "name", OdtProperty), Stamped(bands, title, file).ToJson()));
		var added = zip.CreateEntry("meta.xml", CompressionLevel.Optimal);
		using (var stream = added.Open()) document.Save(stream, SaveOptions.DisableFormatting);
	}

	/// The header and footer of an .odt (null if it has neither): from its first master page.
	public static PageBands? FromOdt(string path) {
		try {
			using var zip = ZipFile.OpenRead(path);
			PageBands? exact = null;
			if (zip.GetEntry("meta.xml") is { } metaEntry) {
				using var stream = metaEntry.Open();
				var value = XDocument.Load(stream).Descendants(Meta + "user-defined").FirstOrDefault(e => (string?)e.Attribute(Meta + "name") == OdtProperty)?.Value;
				if (value != null) exact = PageBands.FromJson(value);
			}
			if (zip.GetEntry("styles.xml") is not { } stylesEntry) return Decide(exact, null, null, false);
			XDocument styles;
			using (var stream = stylesEntry.Open()) styles = XDocument.Load(stream);
			var page = styles.Descendants(Style + "master-page").FirstOrDefault(m => (string?)m.Attribute(Style + "name") == "Standard")
				?? styles.Descendants(Style + "master-page").FirstOrDefault();
			var paragraphStyles = styles.Descendants(Style + "style").Where(s => s.Attribute(Style + "name") != null)
				.GroupBy(s => (string)s.Attribute(Style + "name")!).ToDictionary(g => g.Key, g => g.First());
			string[]? Read(string band) {
				var element = page?.Element(Style + band);
				if (element == null || (string?)element.Attribute(Style + "display") == "false") return null;
				// (the band is one line of three parts: a paragraph's parts go where the paragraphs before left room)
				string[]? result = null;
				foreach (var paragraph in element.Descendants(Text + "p")) {
					string styleName = (string?)paragraph.Attribute(Text + "style-name") ?? "";
					string? align = paragraphStyles.TryGetValue(styleName, out var s) ? (string?)s.Element(Style + "paragraph-properties")?.Attribute(Fo + "text-align") : null;
					if (PartsOf(OdtText(paragraph), align is "center" ? 1 : align is "end" or "right" ? 2 : 0) is not { } parts) continue;
					if (result == null) { result = parts; continue; }
					for (int i = 0; i < 3 && i < parts.Length && i < result.Length; i++)
						if (result[i].Trim() == "" && parts[i].Trim() != "") result[i] = parts[i];
				}
				return result;
			}
			bool haveReal = page?.Element(Style + "header") != null || page?.Element(Style + "footer") != null;
			return Decide(exact, Read("header"), Read("footer"), haveReal);
		} catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException) {
			return null;
		}
	}

	static string OdtText(XElement element) {
		var sb = new StringBuilder();
		foreach (var node in element.Nodes()) {
			if (node is XText text) { sb.Append(text.Value); continue; }
			if (node is not XElement e) continue;
			if (e.Name.Namespace != Text) { sb.Append(OdtText(e)); continue; }
			switch (e.Name.LocalName) {
				case "tab": sb.Append('\t'); break;
				case "s": sb.Append(' ', (int?)e.Attribute(Text + "c") ?? 1); break;
				case "line-break": sb.Append(' '); break;
				case "page-number": sb.Append("{page}"); break;
				case "page-count": sb.Append("{pages}"); break;
				case "date": sb.Append("{date}"); break;
				case "time": sb.Append("{time}"); break;
				case "file-name": sb.Append((string?)e.Attribute(Text + "display") == "full" ? "{file}" : "{title}"); break;
				default: sb.Append(OdtText(e)); break;
			}
		}
		return sb.ToString();
	}

	// ================================================================ HTML and Markdown

	public static string HtmlMeta(PageBands bands) =>
		$"<meta name=\"{MetaName}\" content=\"{System.Net.WebUtility.HtmlEncode(bands.Copy().ToJson())}\">\r\n";

	public static PageBands? FromHtml(string html) {
		var m = Regex.Match(html, "<meta\\s+name=\"" + MetaName + "\"\\s+content=\"([^\"]*)\"", RegexOptions.IgnoreCase);
		return m.Success ? PageBands.FromJson(System.Net.WebUtility.HtmlDecode(m.Groups[1].Value)) : null;
	}

	public static string MarkdownLine(PageBands bands) => $"<!-- {MetaName}: {bands.Copy().ToJson().Replace("--", "\\u002D\\u002D")} -->";

	public static PageBands? FromMarkdown(string markdown) {
		var m = Regex.Match(markdown, @"^<!-- " + MetaName + @": (.*) -->\s*$", RegexOptions.Multiline);
		return m.Success ? PageBands.FromJson(m.Groups[1].Value) : null;
	}
}
