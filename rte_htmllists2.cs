// List numbering in HTML for RichEditBox (used by RTE Writer, rte_writer10.cs, through #load).
// Version 2 adds legal numbering (rte_legal.cs): a legal list is <ol class="rte-legal">, its items without their
// number text, and CSS counters draw "1.2.1" in browsers.
// RichEditorNET's HTML keeps the numbering style of a list (<ol type="a">) but not its punctuation, its start number,
// or a restart in the middle of a list. This adds them to the HTML it writes, and reads them back after it loads HTML:
//   start number   <ol start="5">
//   restart        </ol><ol start="..."> where the numbering starts again inside one list
//   punctuation    a class such as <ol type="a" class="rte-lower-alpha-paren"> for "a)", "(a)" or "a", with a
//                  <style> block that draws the markers (li::marker) in browsers; "a." is HTML's own
// Each <li> of the library's HTML is one list paragraph of the document, in the same order; if that ever does not hold,
// the HTML is left as it is.
using System.Text;
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.TextObjectModel2;

public static class HtmlLists
{
	// Legal items: Top is the number of their top level (1 for 1.2.3), Series their numbering series (LegalLists.Numbers).
	sealed record Item(int Level, int Style, int Punctuation, int Start, bool Legal, int Top = 1, int Series = 0, string MarkStyle = "");

	const string LegalClass = "rte-legal";
	static readonly Regex LegalNumber = new(@"\G((?:<[^>]+>)*)\d+(?:\.\d+)*\.?(?:&#9;|\t)");

	static readonly Regex ListTag = new(@"<(/?)(ol|ul|li)\b([^>]*)>", RegexOptions.IgnoreCase);
	static readonly Regex ClassAttribute = new(@"\bclass\s*=\s*""[^""]*\brte-([a-z]+(?:-[a-z]+)?)-(paren|parens|plain)\b[^""]*""", RegexOptions.IgnoreCase);
	static readonly Regex StartAttribute = new(@"\bstart\s*=\s*""?(-?\d+)", RegexOptions.IgnoreCase);

	static string? CounterStyle(int style) =>
		style == tomConstants.tomListNumberAsArabic ? "decimal" : style == tomConstants.tomListNumberAsLCLetter ? "lower-alpha"
		: style == tomConstants.tomListNumberAsUCLetter ? "upper-alpha" : style == tomConstants.tomListNumberAsLCRoman ? "lower-roman"
		: style == tomConstants.tomListNumberAsUCRoman ? "upper-roman" : null;

	static string? PunctuationName(int punctuation) => punctuation switch {
		0 => "paren", 0x10000 => "parens", 0x30000 => "plain", _ => null    // 0x20000, "1.", is HTML's own
	};

	static int PunctuationOf(string name) => name.ToLowerInvariant() switch {
		"paren" => 0, "parens" => tomConstants.tomListParentheses, "plain" => tomConstants.tomListPlain, _ => tomConstants.tomListPeriod
	};

	// ================================================================ save

	/// The editor's HTML with the lists' start numbers, restarts and punctuation added. css gets the <style> rules the
	/// punctuation classes need ("" if none).
	public static string Export(RichEditBox editor, string html, out string css) {
		css = "";
		var items = ListItems(editor.TextDocument, editor.Font);
		var tags = ListTag.Matches(html);
		// (the library writes an <li> with nothing but a list in it where a level is skipped: not an item of the document)
		bool Wrapper(Match tag) => tag.Groups[1].Value == "" && tag.Groups[2].Value.Equals("li", StringComparison.OrdinalIgnoreCase)
			&& Regex.IsMatch(html.Substring(tag.Index + tag.Length, Math.Min(12, html.Length - tag.Index - tag.Length)), @"^\s*<(ol|ul)\b", RegexOptions.IgnoreCase);
		// the page's own font, for list numbers and bullets (browsers draw them in the <li>'s font) and text without one
		string page = "body " + editor.Font.Name.Replace("'", "") + " " + editor.Font.SizeInPoints.ToString(System.Globalization.CultureInfo.InvariantCulture);
		if (tags.Count(t => t.Groups[1].Value == "" && t.Groups[2].Value.Equals("li", StringComparison.OrdinalIgnoreCase) && !Wrapper(t)) != items.Count) {
			css = Rule(page);
			return html;
		}

		var sb = new StringBuilder();
		var open = new List<(string Tag, string Attributes, int OutputPosition, int LastStart, int LastItem)>();    // the open lists
		var classes = new SortedSet<string> { page };
		int position = 0, k = 0;
		foreach (Match tag in tags) {
			sb.Append(html, position, tag.Index - position);
			position = tag.Index + tag.Length;
			string name = tag.Groups[2].Value.ToLowerInvariant();
			bool closing = tag.Groups[1].Value == "/";
			if (name is "ol" or "ul") {
				if (closing) { if (open.Count > 0) open.RemoveAt(open.Count - 1); sb.Append(tag.Value); }
				else { open.Add((name, tag.Groups[3].Value, sb.Length, int.MinValue, -1)); sb.Append(tag.Value); }
				continue;
			}
			if (name != "li" || closing || open.Count == 0 || Wrapper(tag)) { sb.Append(tag.Value); continue; }

			var item = items[k++];
			var list = open[^1];
			// a legal list right after an ordinary one (or the other way round) is a list of its own
			if (list.Tag == "ol" && list.LastItem >= 0 && items[list.LastItem].Legal != item.Legal) {
				sb.Append("</ol><ol").Append(list.Attributes).Append('>');
				open[^1] = list = (list.Tag, list.Attributes, sb.Length - ("<ol" + list.Attributes + ">").Length, int.MinValue, -1);
			}
			string liTag = item.MarkStyle == "" ? tag.Value : tag.Value[..^1] + $" style=\"{item.MarkStyle}\">";
			if (list.Tag == "ol" && item.Legal) {
				// a legal list: the class (and the top level's first number) on the <ol>; the number written in the text
				// goes. A list that carries on from one before it starts at its next number; a restart inside a list
				// closes it and opens another.
				classes.Add(LegalClass);
				string extra = $" class=\"{LegalClass}\"" + (item.Level == 0 && item.Top != 1 ? $" start=\"{item.Top}\" style=\"counter-reset: {LegalClass} {item.Top - 1}\"" : "");
				if (list.LastItem < 0) {
					sb.Insert(list.OutputPosition + 3, extra);
				} else if (item.Level == 0 && item.Series != items[list.LastItem].Series) {
					sb.Append("</ol><ol").Append(list.Attributes).Append(extra).Append('>');
					open[^1] = (list.Tag, list.Attributes, sb.Length - ("<ol" + list.Attributes + extra + ">").Length, item.Start, k - 1);
				}
				open[^1] = (list.Tag, list.Attributes, open[^1].OutputPosition, item.Start, k - 1);
				sb.Append(liTag);
				if (LegalNumber.Match(html, position) is { Success: true } number) {
					sb.Append(number.Groups[1].Value);
					position += number.Length;
				}
				continue;
			}
			if (list.Tag == "ol") {
				string extra = Attributes(item, classes);
				if (list.LastItem < 0) {
					// the first item: its numbering becomes the list's
					if (extra != "") sb.Insert(list.OutputPosition + 3, extra);
				} else if (item.Start != list.LastStart || item.Punctuation != items[list.LastItem].Punctuation) {
					// the numbering starts again inside the list: close it and open another
					sb.Append("</ol><ol").Append(list.Attributes).Append(extra).Append('>');
					open[^1] = (list.Tag, list.Attributes, sb.Length - ("<ol" + list.Attributes + extra + ">").Length, item.Start, k - 1);
					sb.Append(liTag);
					continue;
				}
				open[^1] = (list.Tag, list.Attributes, list.OutputPosition, item.Start, k - 1);
			}
			sb.Append(liTag);
		}
		sb.Append(html, position, html.Length - position);
		css = string.Concat(classes.Select(Rule));
		return sb.ToString();
	}

	// start="..." and class="rte-..." for an item's list, or "" when HTML's defaults do (start 1, "1.").
	static string Attributes(Item item, SortedSet<string> classes) {
		string attributes = item.Start != 1 ? $" start=\"{item.Start}\"" : "";
		if (PunctuationName(item.Punctuation) is string punctuation && CounterStyle(item.Style) is string style) {
			string name = $"rte-{style}-{punctuation}";
			classes.Add(name);
			attributes += $" class=\"{name}\"";
		}
		return attributes;
	}

	static string Rule(string name) {
		if (name.StartsWith("body ")) {
			int last = name.LastIndexOf(' ');
			return $"body {{ font-family: '{name[5..last]}'; font-size: {name[(last + 1)..]}pt; }}\r\n";
		}
		if (name == LegalClass) {
			string top = LegalLists.Format == LegalLists.Periods.None ? "\" \"" : "\". \"";
			string below = LegalLists.Format == LegalLists.Periods.All ? "\". \"" : "\" \"";
			return $"ol.{LegalClass} {{ list-style-type: none; counter-reset: {LegalClass}; }}\r\n"
				+ $"ol.{LegalClass} > li {{ counter-increment: {LegalClass}; }}\r\n"
				+ $"ol.{LegalClass} > li::before {{ content: counters({LegalClass}, \".\") {top}; }}\r\n"
				+ $"ol.{LegalClass} ol.{LegalClass} > li::before {{ content: counters({LegalClass}, \".\") {below}; }}\r\n";
		}
		var m = Regex.Match(name, @"^rte-(.+)-(paren|parens|plain)$");
		string counter = $"counter(list-item, {m.Groups[1].Value})";
		string content = m.Groups[2].Value switch { "paren" => counter + " \") \"", "parens" => "\"(\" " + counter + " \") \"", _ => counter + " \" \"" };
		return $"ol.{name} > li::marker {{ content: {content}; }}\r\n";
	}

	/// A <style> element for css (from Export), or "".
	public static string StyleElement(string css) => css == "" ? "" : "<style>\r\n" + css + "</style>\r\n";

	// ================================================================ load

	/// Loads HTML into the editor (with RichEditorNET), then applies the lists' start numbers and punctuation.
	public static void Load(RichEditBox editor, string html) {
		editor.Html = html;
		// the numbering each <li> should have, from its <ol>: punctuation class and start
		var wanted = new List<(int? Punctuation, int Start, bool Legal)?>();
		var open = new List<(int? Punctuation, int Start, bool Legal)?>();
		var firstOfList = new List<bool>();    // each <li>: the first of its list
		var itemsSeen = new List<int>();       // the open lists' <li> counts
		foreach (Match tag in ListTag.Matches(html)) {
			string name = tag.Groups[2].Value.ToLowerInvariant();
			bool closing = tag.Groups[1].Value == "/";
			if (name is "ol" or "ul") {
				if (closing) { if (open.Count > 0) { open.RemoveAt(open.Count - 1); itemsSeen.RemoveAt(itemsSeen.Count - 1); } continue; }
				itemsSeen.Add(0);
				if (name == "ul") { open.Add(null); continue; }
				var cls = ClassAttribute.Match(tag.Groups[3].Value);
				var start = StartAttribute.Match(tag.Groups[3].Value);
				bool legal = Regex.IsMatch(tag.Groups[3].Value, @"\bclass\s*=\s*""[^""]*\b" + LegalClass + @"\b");
				open.Add((cls.Success ? PunctuationOf(cls.Groups[2].Value) : null, start.Success ? int.Parse(start.Groups[1].Value) : 1, legal));
			} else if (name == "li" && !closing) {
				wanted.Add(open.Count > 0 ? open[^1] : null);
				firstOfList.Add(itemsSeen.Count > 0 && itemsSeen[^1]++ == 0);
			}
		}
		var doc = editor.TextDocument;
		var paragraphs = ListParagraphs(doc);
		if (paragraphs.Count != wanted.Count) return;    // the HTML does not map onto the document's list paragraphs
		bool anyLegal = false;
		for (int i = 0; i < paragraphs.Count; i++) {
			if (wanted[i] is not { } w) continue;
			var para = doc.Range2(paragraphs[i], paragraphs[i]).Para;
			int style = para.ListType & 0xFFFF;
			if (style == tomConstants.tomListNone || style == tomConstants.tomListBullet) continue;
			if (w.Legal) {
				// the first item of a top-level legal <ol> starts a list at its start (LegalLists.Normalize decides
				// whether that goes on from the list before); the others go on
				bool first = firstOfList[i];
				LegalLists.SetLegal(para, Math.Max(0, para.ListLevelIndex), first && para.ListLevelIndex <= 0 ? LegalLists.Loaded + w.Start : 0);
				anyLegal = true;
				continue;
			}
			if (w.Punctuation is int punctuation) para.ListType = style | punctuation;
			para.ListStart = w.Start;
		}
		if (anyLegal) {
			LegalLists.Normalize(doc);
			LegalLists.Renumber(editor);
		}
		editor.ClearUndo();
	}

	// ================================================================ the document's list paragraphs

	static List<int> ListParagraphs(ITextDocument2 doc) {
		var starts = new List<int>();
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var range = doc.Range2(0, 0);
		range.Expand(tomConstants.tomParagraph);
		while (true) {
			if ((range.Para.ListType & 0xFFFF) != tomConstants.tomListNone) starts.Add(range.Start);
			if (range.End >= storyEnd || range.End <= range.Start) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
		return starts;
	}

	static List<Item> ListItems(ITextDocument2 doc, System.Drawing.Font? pageFont = null) {
		var legal = LegalLists.Numbers(LegalLists.Paragraphs(doc));
		return ListParagraphs(doc).Select(start => {
			var para = doc.Range2(start, start).Para;
			bool isLegal = LegalLists.IsLegal(para.ListType) && legal.ContainsKey(start);
			return new Item(para.ListLevelIndex, para.ListType & 0xFFFF, para.ListType & 0x70000, Math.Max(0, para.ListStart), isLegal,
				isLegal ? legal[start].Counters[0] : 1, isLegal ? legal[start].Series : 0, MarkStyle(doc, start, pageFont));
		}).ToList();
	}

	// The font of an item's number, as CSS for its <li> (RichEdit draws numbers in the paragraph mark's format; a legal
	// number is the item's first characters): "" where it is the page's.
	static string MarkStyle(ITextDocument2 doc, int start, System.Drawing.Font? pageFont) {
		if (pageFont == null) return "";
		var paragraph = doc.Range2(start, start);
		paragraph.Expand(tomConstants.tomParagraph);
		var para = paragraph.Para;
		int at = LegalLists.IsLegal(para.ListType) ? paragraph.Start : Math.Max(paragraph.Start, paragraph.End - 1);
		var font = doc.Range2(at, at + 1).Font;
		var css = new List<string>();
		if (font.Size > 0 && Math.Abs(font.Size - pageFont.SizeInPoints) > 0.01) css.Add($"font-size: {font.Size.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt");
		if (!string.IsNullOrEmpty(font.Name) && !font.Name.Equals(pageFont.Name, StringComparison.OrdinalIgnoreCase)) css.Add($"font-family: '{font.Name.Replace("'", "")}'");
		return string.Join("; ", css);
	}
}
