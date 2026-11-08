#:package Docxodus@15.1.0
// Docxodus trial (an experiment kept in its own folder, not part of AhkWord): a .docx reader on Docxodus
// (https://github.com/JSv4/Docxodus, MIT), a maintained descendant of Open-Xml-PowerTools, in place of the
// Codeuctivity PowerTools' converter that the built-in "openxml" reader uses. Docxodus's HTML is made to look like
// the Codeuctivity PowerTools' (lists as <ol>/<li> with their numbers in data-pt-list-item-run spans, tabs as tab
// characters, page breaks as <br style="page-break-before: always">), so everything the built-in reader does with that
// HTML afterwards (OfficeConverters.OpenXmlRead's steps) is done the same way.
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

public static class DocxodusReader
{
	static readonly XNamespace Wn = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	const string PageBreakMark = "";    // (a page break, carried through Docxodus as text: it writes every break alike)

	/// The HTML of a .docx (its copy is made in folder; the file itself is never changed).
	public static XElement Html(string docx, string folder, string title, string generalCss) {
		string copy = Path.Combine(folder, "docxodus-" + Path.GetFileName(docx));
		File.Copy(docx, copy, true);
		MarkPageBreaks(copy);
		var settings = new global::Docxodus.WmlToHtmlConverterSettings {
			PageTitle = title, FabricateCssClasses = false, CssClassPrefix = "pt-", GeneralCss = generalCss, AdditionalCss = "", ImageHandler = Image };
		var html = global::Docxodus.WmlToHtmlConverter.ConvertToHtml(new global::Docxodus.WmlDocument(copy), settings);
		var lists = ListParagraphs(copy);
		Normalize(html, lists);
		return html;
	}

	// ---- before converting

	static void MarkPageBreaks(string path) {
		using var package = WordprocessingDocument.Open(path, true);
		var body = package.MainDocumentPart?.Document?.Body;
		if (body == null) return;
		foreach (var br in body.Descendants<W.Break>().Where(b => b.Type?.Value == W.BreakValues.Page).ToList()) {
			br.InsertAfterSelf(new W.Text(PageBreakMark));
			br.Remove();
		}
		package.MainDocumentPart!.Document.Save();
	}

	// Each paragraph's list (its numbering instance and level; null: none), in the order the converter writes them: its
	// own numbering or its style's (and the styles it is based on), as DocxLineSpacing.Paragraphs takes the paragraphs.
	// (Id: the numbering instance; Definition: its abstract list, which the levels of one list can share across instances)
	static List<(string Id, string Definition, int Level, bool Bullet)?> ListParagraphs(string path) {
		using var zip = System.IO.Compression.ZipFile.OpenRead(path);
		XDocument? Load(string name) { var e = zip.GetEntry(name); if (e == null) return null; using var s = e.Open(); return XDocument.Load(s); }
		var document = Load("word/document.xml")!;
		var styles = Load("word/styles.xml");
		var numbering = Load("word/numbering.xml")?.Root;
		var nums = numbering?.Elements(Wn + "num").Where(n => n.Attribute(Wn + "numId") != null)
			.GroupBy(n => (string)n.Attribute(Wn + "numId")!).ToDictionary(g => g.Key, g => g.First()) ?? new();
		var definitionOf = nums.ToDictionary(n => n.Key, n => (string?)n.Value.Element(Wn + "abstractNumId")?.Attribute(Wn + "val") ?? n.Key);
		var abstracts = numbering?.Elements(Wn + "abstractNum").Where(a => a.Attribute(Wn + "abstractNumId") != null)
			.GroupBy(a => (string)a.Attribute(Wn + "abstractNumId")!).ToDictionary(g => g.Key, g => g.First()) ?? new();
		// (a level's number format: an override of the instance's own, or its definition's; Word's bullets are often
		// letters in a symbol font, "o" in Courier New or "§" in Wingdings, which read as numbers by their text)
		bool IsBullet(string numId, int level) {
			XElement? Level(XElement? holder) => holder?.Elements(Wn + "lvl").FirstOrDefault(l => (int?)l.Attribute(Wn + "ilvl") == level);
			var over = nums.GetValueOrDefault(numId)?.Elements(Wn + "lvlOverride").FirstOrDefault(o => (int?)o.Attribute(Wn + "ilvl") == level);
			var lvl = Level(over) ?? Level(abstracts.GetValueOrDefault(definitionOf.GetValueOrDefault(numId, numId)));
			return (string?)lvl?.Element(Wn + "numFmt")?.Attribute(Wn + "val") == "bullet";
		}
		var byId = styles?.Descendants(Wn + "style").Where(s => (string?)s.Attribute(Wn + "type") == "paragraph" && s.Attribute(Wn + "styleId") != null)
			.GroupBy(s => (string)s.Attribute(Wn + "styleId")!).ToDictionary(g => g.Key, g => g.First()) ?? new();
		string? defaultStyle = (string?)styles?.Descendants(Wn + "style").FirstOrDefault(s => (string?)s.Attribute(Wn + "type") == "paragraph" && (string?)s.Attribute(Wn + "default") is "1" or "true")?.Attribute(Wn + "styleId");
		(string? Id, int? Level) FromStyle(string? id) {
			string? numId = null; int? level = null;
			for (int depth = 0; id != null && depth < 20 && byId.TryGetValue(id, out var style); depth++) {
				var numPr = style.Element(Wn + "pPr")?.Element(Wn + "numPr");
				numId ??= (string?)numPr?.Element(Wn + "numId")?.Attribute(Wn + "val");
				level ??= (int?)numPr?.Element(Wn + "ilvl")?.Attribute(Wn + "val");
				id = (string?)style.Element(Wn + "basedOn")?.Attribute(Wn + "val");
			}
			return (numId, level);
		}
		var list = new List<(string, string, int, bool)?>();
		foreach (var p in document.Root!.Element(Wn + "body")!.Descendants(Wn + "p")) {
			if (p.Ancestors(Wn + "txbxContent").Any() || DocxLineSpacing.InMergedCell(p)) continue;
			var pPr = p.Element(Wn + "pPr");
			var own = pPr?.Element(Wn + "numPr");
			var fromStyle = FromStyle((string?)pPr?.Element(Wn + "pStyle")?.Attribute(Wn + "val") ?? defaultStyle);
			string? numId = (string?)own?.Element(Wn + "numId")?.Attribute(Wn + "val") ?? fromStyle.Id;
			int level = (int?)own?.Element(Wn + "ilvl")?.Attribute(Wn + "val") ?? fromStyle.Level ?? 0;
			list.Add(numId is null or "0" ? null : (numId, definitionOf.GetValueOrDefault(numId, numId), level, IsBullet(numId, level)));
		}
		return list;
	}

	// Pictures as the PowerTools' own handler writes them: a PNG inside the HTML, sized by a style.
	static XElement? Image(global::Docxodus.ImageInfo info) {
		byte[]? png = null;
		try { png = info.Bitmap?.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100)?.ToArray(); } catch { }
		string contentType = "image/png";
		if (png == null && info.ImageBytes != null && info.ContentType is "image/png" or "image/jpeg" or "image/gif") { png = info.ImageBytes; contentType = info.ContentType!; }
		if (png == null) return null;
		var img = new XElement(XNamespace.Get("http://www.w3.org/1999/xhtml") + "img", new XAttribute("src", $"data:{contentType};base64,{Convert.ToBase64String(png)}"));
		if (info.ImgStyleAttribute != null) img.Add(new XAttribute("style", info.ImgStyleAttribute.Value));
		if (info.AltText != null) img.Add(new XAttribute("alt", info.AltText));
		return img;
	}

	// ---- after converting: the PowerTools' shape

	static void Normalize(XElement html, List<(string Id, string Definition, int Level, bool Bullet)?> lists) {
		XNamespace x = html.Name.Namespace;
		// (the direction marks Docxodus writes after a break)
		foreach (var br in html.Descendants(x + "br").ToList()) {
			br.Attribute("data-docx-break")?.Remove();
			if (br.NextNode is XText t) {
				string rest = t.Value.TrimStart('‎', '‏');
				if (rest.Length != t.Value.Length) { if (rest == "") t.Remove(); else t.Value = rest; }
				else if (t.GetType() != typeof(XText) && t.ToString() is "&#x200e;" or "&#x200f;") t.Remove();
			}
		}
		// list numbers: the number's text in a marker span of the PowerTools' kind, the layout around it gone
		var blocks = html.Descendants().Where(e => e.Name.Namespace == x && Regex.IsMatch(e.Name.LocalName, "^(p|h[1-6])$")).ToList();
		bool paired = blocks.Count == lists.Count;
		var info = new Dictionary<XElement, (string Id, string Definition, int Level, bool Bullet)>();
		for (int i = 0; i < blocks.Count; i++) {
			// (a number followed by a tab: one box holding the number and the tab; by a space or nothing: the number's span,
			// then the space's)
			var markers = blocks[i].Descendants(x + "span").Where(s => s.Attribute("data-list-marker") != null && !s.Ancestors(x + "span").Any(a => a.Attribute("data-list-marker") != null)).ToList();
			if (markers.Count == 0) continue;
			string number = string.Concat(markers[0].DescendantNodes().OfType<XText>().Select(t => t.Value)).Trim();
			bool tab = markers[0].Descendants().Any(e => e.Attribute("data-docx-tab") != null);
			bool space = markers.Skip(1).Any(m => m.Value.Contains(' ') || m.Value.Contains((char)0xA0));
			foreach (var extra in markers.Skip(1)) extra.Remove();
			var item = paired ? lists[i] : null;
			if (item is { Bullet: true }) number = "•";    // (a bullet whatever its character)
			if (!Representable(number)) {
				// a number the editor's lists can't show ("Article I.", "Section 1.01"): kept as typed text, as Word's own
				// conversion keeps it, rather than a list without it
				markers[0].ReplaceWith(new XElement(x + "span", number + (tab ? "\t" : space ? " " : "")));
				continue;
			}
			markers[0].ReplaceWith(new XElement(x + "span", new XAttribute("data-pt-list-item-run", "1"), number));
			info[blocks[i]] = item is { } l ? l : ("?", "?", LevelByIndent(blocks[i]), false);
		}
		// tabs: the tab itself, the boxes that lay out the text before it gone
		foreach (var tab in html.Descendants(x + "span").Where(s => s.Attribute("data-docx-tab") != null).ToList()) tab.ReplaceWith(new XText("\t"));
		foreach (var box in html.Descendants(x + "span").Where(s => ((string?)s.Attribute("style") ?? "").Contains("flex", StringComparison.Ordinal)).Reverse().ToList())
			box.ReplaceWith(box.Nodes().ToList());
		// page breaks
		foreach (var text in html.DescendantNodes().OfType<XText>().Where(t => t.Value.Contains(PageBreakMark)).ToList()) {
			var parts = text.Value.Split(PageBreakMark);
			var nodes = new List<XNode>();
			for (int i = 0; i < parts.Length; i++) {
				if (i > 0) nodes.Add(new XElement(x + "br", new XAttribute("style", "page-break-before: always")));
				if (parts[i] != "") nodes.Add(new XText(parts[i]));
			}
			text.ReplaceWith(nodes);
		}
		// the list paragraphs that follow each other in one list, in <ol>s nested by level (a level's list after its item):
		// one list while they share their list definition (a level of it can be another numbering instance), a new one
		// where the definition changes or a top-level item is another instance's (a list numbered again from its start)
		foreach (var container in html.DescendantsAndSelf().Where(e => e.Elements().Any(c => info.ContainsKey(c))).ToList()) {
			var stack = new List<(int Level, XElement List)>();
			string? definition = null, instance = null;
			foreach (var child in container.Elements().ToList()) {
				if (!info.TryGetValue(child, out var item)) { stack.Clear(); definition = instance = null; continue; }
				if (stack.Count == 0 || item.Definition != definition || (item.Id != instance && item.Level == 0)) {
					// (a list that starts below its first level: the levels above it as lists without items of their own)
					var top = new XElement(x + "ol");
					child.AddBeforeSelf(top);
					stack.Clear();
					for (int level = 0; level < item.Level; level++) { stack.Add((level, top)); var inner = new XElement(x + "ol"); top.Add(inner); top = inner; }
					if (Start(child) is int s) top.SetAttributeValue("start", s);
					stack.Add((item.Level, top)); definition = item.Definition; instance = item.Id;
				} else {
					instance = item.Id;
					while (stack.Count > 1 && stack[^1].Level > item.Level) stack.RemoveAt(stack.Count - 1);
					while (stack[^1].Level < item.Level) {    // (a level at a time, those skipped without items)
						int next = stack[^1].Level + 1;
						var sub = new XElement(x + "ol", next == item.Level && Start(child) is int s ? new XAttribute("start", s) : null);
						stack[^1].List.Add(sub);
						stack.Add((next, sub));
					}
				}
				child.Remove();
				stack[^1].List.Add(new XElement(x + "li", child));
			}
		}
		// a list right after another (no paragraph between): RichEdit would number it on from the one before; its first
		// item is marked, and numbered from its own start once the document is loaded (OfficeConverters.RestartsFromMarkers)
		foreach (var list in html.Descendants(x + "ol").Where(l => l.Parent?.Name != x + "ol").ToList()) {
			if (list.ElementsBeforeSelf().LastOrDefault() is not { } before || before.Name != x + "ol" && before.Name != x + "ul") continue;
			if (list.Elements(x + "li").FirstOrDefault()?.Elements().FirstOrDefault() is { } first) first.AddFirst(new XText(OfficeConverters.RestartMark.ToString()));
		}
	}

	// A list number the editor's lists can show: a bullet, "3." "3)" "(c)" "iv." and the like (as ListsFromMarkers reads
	// them), or legal numbering's "1.2.3".
	static bool Representable(string number) =>
		number.Length <= 2 && !char.IsLetterOrDigit(number.FirstOrDefault()) || Regex.IsMatch(number, @"^(\()?([0-9]+|[a-zA-Z]+)([.)])?$") || Regex.IsMatch(number, @"^\d+(\.\d+)+\.?$");

	// (where the paragraphs couldn't be paired with the file's: a level by its indent)
	static int LevelByIndent(XElement block) {
		var m = Regex.Match((string?)block.Attribute("style") ?? "", @"margin-left:\s*([\d.]+)pt");
		return m.Success && double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double left) ? (int)Math.Max(0, Math.Round(left / 18) - 2) : 0;
	}

	// A list's first number, from its first item's marker ("3.", "1.2", "(c)", "iv)"); null for bullets
	static int? Start(XElement block) {
		var marker = block.Descendants().FirstOrDefault(e => e.Attribute("data-pt-list-item-run") != null);
		var m = Regex.Match(marker?.Value.Trim() ?? "", @"([0-9]+|[a-zA-Z]+)[.)]?$");
		if (!m.Success) return null;
		string n = m.Groups[1].Value;
		if (char.IsDigit(n[0])) return int.TryParse(n, out int v) ? v : null;
		if (Regex.IsMatch(n, "^[ivxlcdm]+$", RegexOptions.IgnoreCase) && (n.Length > 1 || n is "i" or "I")) return Roman(n.ToLowerInvariant());
		return n.Length == 1 ? char.ToLowerInvariant(n[0]) - 'a' + 1 : null;
	}

	static int Roman(string s) {
		int total = 0, last = 0;
		for (int i = s.Length - 1; i >= 0; i--) {
			int v = s[i] switch { 'i' => 1, 'v' => 5, 'x' => 10, 'l' => 50, 'c' => 100, 'd' => 500, _ => 1000 };
			total += v < last ? -v : v; last = Math.Max(last, v);
		}
		return total;
	}
}
