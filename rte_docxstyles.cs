// Paragraph styles in the .docx files the built-in converter (rte_office7.cs's OpenXml) reads and writes, for RTE
// Writer (rte_writer36.cs, through #load). The converter goes through HTML (the PowerTools reading, HtmlToOpenXml
// writing), which keeps only headings, and those with its own look. So:
//   - reading: the .docx's paragraph styles, as Word works them out (the document's defaults, the styles each is based
//     on, the theme's fonts), go into the reader's HTML as RTE Writer's own HTML has them (StyleHtml): each paragraph's
//     style (data-rte-style), its line and page breaks (data-rte-breaks), and the styles' definitions (the
//     rte-writer-styles meta). LibreOffice's .odt files are read the same way (LibreOffice makes a .docx of them).
//   - writing: the .docx gets the document's styles (each style's whole definition, not based on another, so that it
//     reads the same in Word and LibreOffice), and each paragraph its style and its line and page breaks.
// The paragraphs are paired up in order (DocxLineSpacing.Pair); if they don't pair up, the file is left as the
// converter made it.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static class DocxStyles
{
	static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
	static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

	/// RTE Writer's name for a Word style name ("heading 1": Heading 1).
	public static string NameOf(string wordName) =>
		ParagraphStyle.HeadingNumber(wordName) is int n and > 0 ? "Heading " + n
		: string.Equals(wordName, "Normal", StringComparison.OrdinalIgnoreCase) ? "Normal" : wordName.Trim();

	/// The styleId for a style: Normal, Heading1 ... (Word's own), others their name's letters and digits.
	public static string IdOf(string name) {
		if (ParagraphStyle.HeadingNumber(name) is int n and > 0) return "Heading" + n;
		string id = new string(name.Where(char.IsLetterOrDigit).ToArray());
		return id == "" ? "Style" + Math.Abs(name.GetHashCode() % 100000) : id;
	}

	static string WordName(string name) => ParagraphStyle.HeadingNumber(name) is int n and > 0 ? "heading " + n : name;

	// ================================================================ reading

	/// The .docx's paragraph styles, each as Word works it out; Names: styleId to RTE Writer's name.
	public static (StyleSheet Sheet, Dictionary<string, string> Names) Read(string path) {
		using var zip = ZipFile.OpenRead(path);
		XDocument? Load(string name) { var e = zip.GetEntry(name); if (e == null) return null; using var s = e.Open(); return XDocument.Load(s); }
		var styles = Load("word/styles.xml");
		var theme = Load("word/theme/theme1.xml");
		var sheet = new StyleSheet { Set = "", Styles = new() };
		var names = new Dictionary<string, string>();
		if (styles?.Root == null) return (sheet, names);
		var byId = styles.Root.Elements(W + "style").Where(s => (string?)s.Attribute(W + "type") == "paragraph" && s.Attribute(W + "styleId") != null)
			.GroupBy(s => (string)s.Attribute(W + "styleId")!).ToDictionary(g => g.Key, g => g.First());
		foreach (var (id, style) in byId) names[id] = NameOf((string?)style.Element(W + "name")?.Attribute(W + "val") ?? id);
		string ThemeFont(string? which) {
			if (theme?.Root == null || which == null) return "";
			var scheme = theme.Descendants(A + "fontScheme").FirstOrDefault();
			var font = which.StartsWith("major", StringComparison.OrdinalIgnoreCase) ? scheme?.Element(A + "majorFont") : scheme?.Element(A + "minorFont");
			return (string?)font?.Element(A + "latin")?.Attribute("typeface") ?? "";
		}
		var defaults = styles.Root.Element(W + "docDefaults");
		foreach (var (id, element) in byId) {
			// the chain: the style, the ones it is based on (the base first)
			var chain = new List<XElement>();
			for (var (at, depth) = ((string?)id, 0); at != null && depth < 20 && byId.TryGetValue(at, out var s); depth++) {
				chain.Insert(0, s);
				at = (string?)s.Element(W + "basedOn")?.Attribute(W + "val");
			}
			var rPrs = new List<XElement?> { defaults?.Element(W + "rPrDefault")?.Element(W + "rPr") };
			var pPrs = new List<XElement?> { defaults?.Element(W + "pPrDefault")?.Element(W + "pPr") };
			rPrs.AddRange(chain.Select(c => c.Element(W + "rPr")));
			pPrs.AddRange(chain.Select(c => c.Element(W + "pPr")));
			string name = names[id];
			var result = new ParagraphStyle { Name = name, Font = "Times New Roman", Size = 10, WidowControl = false };
			foreach (var rPr in rPrs.Where(r => r != null)) Run(result, rPr!, ThemeFont);
			var format = new DocxLineSpacing.Format();
			string align = "left";
			int? outline = null;
			foreach (var pPr in pPrs.Where(p => p != null)) {
				DocxLineSpacing.MergeFormat(format, pPr!);
				if (pPr!.Element(W + "jc")?.Attribute(W + "val")?.Value is { } jc) align = jc switch { "center" => "center", "right" or "end" => "right", "both" or "distribute" => "justify", _ => "left" };
				if (pPr.Element(W + "outlineLvl")?.Attribute(W + "val")?.Value is { } level && int.TryParse(level, out int lv)) outline = lv;
			}
			result.Align = align;
			result.Before = format.Before / 20f;
			result.After = format.After / 20f;
			result.LeftIndent = (format.Left ?? 0) / 20f;
			result.RightIndent = (format.Right ?? 0) / 20f;
			result.FirstLine = (format.FirstLine ?? 0) / 20f;
			result.LineSpacing = format.Line is int line
				? format.Rule == "auto" ? (Math.Abs(line - 240) <= 1 ? "1" : (line / 240f).ToString("0.##", Invariant))
				: (format.Rule == "atLeast" ? "at least " : "exactly ") + (line / 20f).ToString("0.##", Invariant) + "pt"
				: "1";
			result.KeepWithNext = format.KeepNext;
			result.KeepTogether = format.KeepLines;
			result.PageBreakBefore = format.PageBreakBefore;
			result.WidowControl = format.WidowControl;
			result.Contextual = format.Contextual;
			result.Level = ParagraphStyle.HeadingNumber(name) is int h and > 0 ? h : outline is int o and >= 0 and <= 5 && name != "Normal" ? o + 1 : 0;
			if (element.Element(W + "next")?.Attribute(W + "val")?.Value is { } next && next != id && names.TryGetValue(next, out var nextName)) result.Next = nextName;
			sheet.Styles.Add(result);
		}
		return (sheet, names);
	}

	// A run's properties (rPr) over what a style has so far.
	static void Run(ParagraphStyle style, XElement rPr, Func<string?, string> themeFont) {
		if (rPr.Element(W + "rFonts") is { } fonts) {
			string? font = (string?)fonts.Attribute(W + "ascii") ?? (string?)fonts.Attribute(W + "hAnsi");
			if (font == null && ((string?)fonts.Attribute(W + "asciiTheme") ?? (string?)fonts.Attribute(W + "hAnsiTheme")) is { } t && themeFont(t) is { Length: > 0 } themed) font = themed;
			if (font != null) style.Font = font;
		}
		if (rPr.Element(W + "sz")?.Attribute(W + "val")?.Value is { } size && float.TryParse(size, NumberStyles.Float, Invariant, out float half)) style.Size = half / 2;
		if (rPr.Element(W + "b") is { } b) style.Bold = DocxLineSpacing.On(b);
		if (rPr.Element(W + "i") is { } i) style.Italic = DocxLineSpacing.On(i);
		if (rPr.Element(W + "u") is { } u) style.Underline = (string?)u.Attribute(W + "val") is not ("none" or null) || u.Attribute(W + "val") == null;
		if (rPr.Element(W + "color")?.Attribute(W + "val")?.Value is { } color) style.Color = Regex.IsMatch(color, "^[0-9A-Fa-f]{6}$") && color != "000000" ? "#" + color.ToUpperInvariant() : "";
	}

	/// The built-in reader's HTML (one block per paragraph, paired up in order with the .docx's paragraphs and their
	/// formats): each block gets its paragraph's style and line and page breaks, and the page the styles' definitions,
	/// as RTE Writer's own HTML has them.
	public static void IntoHtml(string path, XElement html, IReadOnlyList<XElement> blocks, IReadOnlyList<DocxLineSpacing.Format> formats) {
		StyleSheet sheet;
		Dictionary<string, string> names;
		try { (sheet, names) = Read(path); } catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException) { return; }
		var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Normal" };
		for (int i = 0; i < blocks.Count && i < formats.Count; i++) {
			var f = formats[i];
			string name = names.TryGetValue(f.Style, out var n) ? n : "Normal";
			used.Add(name);
			// (a heading as a paragraph with its style: the editor's HTML import gives h1-h6 sizes of its own, over the
			// sizes the reader gives the heading's text, which are the file's)
			if (Regex.IsMatch(blocks[i].Name.LocalName, "^h[1-6]$")) {
				if (string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase)) name = "Heading " + blocks[i].Name.LocalName[1];
				blocks[i].Name = blocks[i].Name.Namespace + "p";
			}
			if (!string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase)) blocks[i].SetAttributeValue("data-rte-style", name);
			string breaks = (f.KeepNext ? "n" : "") + (f.KeepLines ? "t" : "") + (f.PageBreakBefore ? "p" : "") + (f.WidowControl ? "" : "o");
			if (breaks != "") blocks[i].SetAttributeValue("data-rte-breaks", breaks);
			// (the PowerTools write page break before as a page break in the paragraph, and a left-to-right mark after it:
			// the paragraph's setting instead)
			if (f.PageBreakBefore)
				foreach (var br in blocks[i].Descendants().Where(e => e.Name.LocalName == "br" && ((string?)e.Attribute("style") ?? "").Contains("page-break-before", StringComparison.OrdinalIgnoreCase)).ToList()) {
					if (br.NextNode is XText mark && mark.Value.StartsWith('‎')) mark.Value = mark.Value[1..];
					br.Remove();
				}
		}
		sheet.Styles.RemoveAll(s => !used.Contains(s.Name));
		XNamespace x = html.Name.Namespace;
		var head = html.Element(x + "head");
		if (head == null) { head = new XElement(x + "head"); html.AddFirst(head); }
		head.Add(new XElement(x + "meta", new XAttribute("name", "rte-writer-styles"),
			new XAttribute("content", Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(sheet))))));
	}

	// ================================================================ writing

	/// A .docx the built-in converter wrote: the document's styles in its styles.xml (replacing the converter's of the
	/// same id), and each paragraph's style and line and page breaks. Left as it is if the paragraphs don't pair up.
	public static void IntoDocx(string path, RichEditBox editor, StyleSheet sheet) {
		var doc = editor.TextDocument;
		XDocument document, styles;
		using (var zip = ZipFile.OpenRead(path)) {
			using (var s = zip.GetEntry("word/document.xml")!.Open()) document = XDocument.Load(s);
			var entry = zip.GetEntry("word/styles.xml");
			if (entry == null) return;
			using (var s = entry.Open()) styles = XDocument.Load(s);
		}
		if (DocxLineSpacing.Pair(doc, document) is not { } pair) return;
		var (paragraphs, starts) = pair;
		var usedStyles = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Normal" };
		for (int i = 0; i < paragraphs.Count; i++) {
			var para = doc.Range2(starts[i], starts[i]).Para;
			var style = StyleEngine.StyleAt(doc, sheet, starts[i]);
			var pPr = paragraphs[i].Element(W + "pPr");
			if (pPr == null) { pPr = new XElement(W + "pPr"); paragraphs[i].AddFirst(pPr); }
			bool list = pPr.Element(W + "numPr") != null;
			if (!(list && style.Name == "Normal")) DocxLineSpacing.PlaceIn(pPr, new XElement(W + "pStyle", new XAttribute(W + "val", IdOf(style.Name))));
			usedStyles.Add(style.Name);
			// the line and page breaks where they aren't the style's
			void Flag(string element, bool on, bool styles) {
				if (on == styles) pPr.Element(W + element)?.Remove();
				else DocxLineSpacing.PlaceIn(pPr, on ? new XElement(W + element) : new XElement(W + element, new XAttribute(W + "val", "0")));
			}
			Flag("keepNext", para.KeepWithNext == tomConstants.tomTrue, style.KeepWithNext);
			Flag("keepLines", para.KeepTogether == tomConstants.tomTrue, style.KeepTogether);
			Flag("pageBreakBefore", para.PageBreakBefore == tomConstants.tomTrue, style.PageBreakBefore);
			// (HtmlToOpenXml writes the HTML's page-break-before as a page break at the paragraph's start: with the
			// paragraph's own setting that would be a second one, an empty page)
			if (para.PageBreakBefore == tomConstants.tomTrue)
				foreach (var run in paragraphs[i].Elements(W + "r").TakeWhile(r => r.Elements().All(e => e.Name == W + "rPr" || e.Name == W + "lastRenderedPageBreak"
					|| (e.Name == W + "br" && (string?)e.Attribute(W + "type") == "page"))).ToList())
					run.Remove();
			Flag("widowControl", StyleEngine.WidowControlOf(para), style.WidowControl);
			Flag("suppressLineNumbers", para.NoLineNumber == tomConstants.tomTrue, false);
		}
		// the styles: each the whole of its definition
		var root = styles.Root!;
		foreach (var style in sheet.Styles.Where(s => usedStyles.Contains(s.Name))) {
			string id = IdOf(style.Name);
			root.Elements(W + "style").Where(e => (string?)e.Attribute(W + "type") == "paragraph" && (string?)e.Attribute(W + "styleId") == id).Remove();
			if (style.Name == "Normal") foreach (var e in root.Elements(W + "style").Where(e => (string?)e.Attribute(W + "type") == "paragraph")) e.Attribute(W + "default")?.Remove();
			root.Add(Definition(style, sheet));
		}
		using var update = ZipFile.Open(path, ZipArchiveMode.Update);
		foreach (var (name, xml) in new[] { ("word/document.xml", document), ("word/styles.xml", styles) }) {
			update.GetEntry(name)!.Delete();
			using var stream = update.CreateEntry(name, CompressionLevel.Optimal).Open();
			xml.Save(stream, SaveOptions.DisableFormatting);
		}
	}

	static int Twips(float points) => (int)Math.Round(points * 20);

	// A style as WordprocessingML: its name, the style after it, its paragraph and run properties, in the schema's order.
	static XElement Definition(ParagraphStyle style, StyleSheet sheet) {
		string id = IdOf(style.Name);
		var element = new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", id));
		if (style.Name == "Normal") element.SetAttributeValue(W + "default", "1");
		element.Add(new XElement(W + "name", new XAttribute(W + "val", WordName(style.Name))));
		string next = style.Next != "" && sheet.Find(style.Next) != null ? IdOf(style.Next) : id;
		element.Add(new XElement(W + "next", new XAttribute(W + "val", next)));
		element.Add(new XElement(W + "qFormat"));
		var pPr = new XElement(W + "pPr");
		if (style.KeepWithNext) pPr.Add(new XElement(W + "keepNext"));
		if (style.KeepTogether) pPr.Add(new XElement(W + "keepLines"));
		if (style.PageBreakBefore) pPr.Add(new XElement(W + "pageBreakBefore"));
		pPr.Add(style.WidowControl ? new XElement(W + "widowControl") : new XElement(W + "widowControl", new XAttribute(W + "val", "0")));
		var spacing = new XElement(W + "spacing", new XAttribute(W + "before", Twips(style.Before)), new XAttribute(W + "after", Twips(style.After)));
		try {
			var (rule, value) = global::LineSpacing.Parse(style.LineSpacing.Trim() == "" ? "1" : style.LineSpacing);
			(int line, string lineRule) = rule switch {
				global::LineSpacing.OneAndHalf => (360, "auto"), global::LineSpacing.Double => (480, "auto"),
				global::LineSpacing.Multiple => ((int)Math.Round(value * 240), "auto"),
				global::LineSpacing.Exactly => (Twips(value), "exact"), global::LineSpacing.AtLeast => (Twips(value), "atLeast"),
				_ => (240, "auto")
			};
			spacing.Add(new XAttribute(W + "line", line), new XAttribute(W + "lineRule", lineRule));
		} catch (ArgumentException) { }
		pPr.Add(spacing);
		var ind = new XElement(W + "ind", new XAttribute(W + "left", Twips(style.LeftIndent)), new XAttribute(W + "right", Twips(style.RightIndent)));
		if (style.FirstLine > 0.05f) ind.Add(new XAttribute(W + "firstLine", Twips(style.FirstLine)));
		else if (style.FirstLine < -0.05f) ind.Add(new XAttribute(W + "hanging", Twips(-style.FirstLine)));
		pPr.Add(ind);
		if (style.Contextual) pPr.Add(new XElement(W + "contextualSpacing"));
		pPr.Add(new XElement(W + "jc", new XAttribute(W + "val", style.Align switch { "center" => "center", "right" => "right", "justify" => "both", _ => "left" })));
		if (style.Level is >= 1 and <= 6) pPr.Add(new XElement(W + "outlineLvl", new XAttribute(W + "val", style.Level - 1)));
		element.Add(pPr);
		var rPr = new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "ascii", style.Font), new XAttribute(W + "hAnsi", style.Font),
			new XAttribute(W + "eastAsia", style.Font), new XAttribute(W + "cs", style.Font)));
		rPr.Add(new XElement(W + "b", new XAttribute(W + "val", style.Bold ? "1" : "0")));
		rPr.Add(new XElement(W + "i", new XAttribute(W + "val", style.Italic ? "1" : "0")));
		if (style.Color is { Length: 7 } color) rPr.Add(new XElement(W + "color", new XAttribute(W + "val", color[1..])));
		int half = (int)Math.Round(style.Size * 2);
		rPr.Add(new XElement(W + "sz", new XAttribute(W + "val", half)), new XElement(W + "szCs", new XAttribute(W + "val", half)));
		rPr.Add(new XElement(W + "u", new XAttribute(W + "val", style.Underline ? "single" : "none")));
		element.Add(rPr);
		return element;
	}
}
