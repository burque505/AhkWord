// Cross-references to numbered paragraphs, for AhkWord (AhkWord.cs, through #load): "see paragraph 3(b)", with the
// number kept up to date as the paragraphs are renumbered.
//
// RichEdit has no bookmarks or fields of its own (it drops Word's bookmarks, and keeps only a REF field's result), so
// both ends are hidden text (CFE_HIDDEN, RTF \v), which RichEdit keeps through editing, undo, copying and its RTF:
//   the target     [[xref:_Ref123456789]] in front of the numbered paragraph's text (after a legal item's or a numbered
//                  style's number, which is text: LegalLists.Prefix)
//   a reference    [[xref=_Ref123456789|w]]3(b)[[/xref]]: the number shown between two markers; after the bar, how it
//                  is shown: "r" the paragraph's own number ("(b)"), "w" with the numbers of the levels above ("3(b)")
// A trailing period is left out ("1." is "1": "paragraph 1 above").
// Files: RTF (and so Word and LibreOffice files) gets Word's own bookmarks and REF fields (\r or \w, and \h), which Word
// and LibreOffice update themselves, and they are read back as markers; HTML gets <a id> and <a href="#..."> links,
// Markdown <a id> and [3(b)](#_Ref123456789) links; plain text only the numbers.
using System.Text;
using System.Text.RegularExpressions;
using EllipticBit.RichEditorNET.TextObjectModel2;
using W = DocumentFormat.OpenXml.Wordprocessing;

public static class CrossRefs
{
	const string Id = @"[A-Za-z_][A-Za-z0-9_]{0,39}";
	public static readonly Regex TargetMark = new(@"\[\[xref:(" + Id + @")\]\]");
	public static readonly Regex Reference = new(@"\[\[xref=(" + Id + @")(?:\|(\w*))?\]\]([^\r\[]*?)\[\[/xref\]\]");
	// (and the footnotes' markers, which hold their notes: rte_footnotes.cs; and the tables of contents': rte_toc.cs)
	static readonly Regex AnyMark = new(@"\[\[xref:" + Id + @"\]\]|\[\[xref=" + Id + @"(?:\|\w*)?\]\]|\[\[/xref\]\]|\[\[fn\]\]|\[\[fn:(?:[^\]\r]|\](?!\]))*\]\]|" + Toc.Marks);

	/// What a reference shows when its target is gone (deleted with its paragraph) or no longer numbered.
	public const string Missing = "(reference not found)";    // (no "[": Reference wouldn't match it)

	/// A numbered paragraph that can be referred to: where it starts, its number (as the editor shows it, without a
	/// trailing period), the full number (with the levels above: "3(b)"), the start of its text, and its target's name
	/// (null: none yet).
	public sealed record Target(int Start, string Number, string Full, string Text, string? Name);

	static string Clean(string number) => number.Trim().TrimEnd('.').Trim();

	// The numbers of a paragraph's levels as one, as a legal citation has them: a letter or roman sub-level in
	// parentheses ("23" and "e" as "23(e)", "(a)" as it is), a number after a period ("2" and "1" as "2.1").
	static string Joined(IEnumerable<string> numbers) {
		var sb = new StringBuilder();
		foreach (string number in numbers.Select(Clean).Where(n => n != "")) {
			if (sb.Length > 0 && !number.StartsWith('(')) {
				if (char.IsDigit(number[0])) sb.Append('.').Append(number);
				else sb.Append('(').Append(number.TrimEnd(')')).Append(')');
				continue;
			}
			sb.Append(number);
		}
		return sb.ToString();
	}

	// A letter or roman number typed in front of a paragraph, then a space or a tab ("e. ", "(iii)<tab>"): a sub-point
	// of the numbered paragraph before it, as a document can have them (not list items). Its number and length.
	static readonly Regex TypedSub = new(@"^(\((?:[a-z]{1,2}|[A-Z]{1,2}|[ivxlcdm]{1,6}|[IVXLCDM]{1,6})\)|(?:[a-z]{1,2}|[A-Z]{1,2}|[ivxlcdm]{1,6}|[IVXLCDM]{1,6})[.)])[ \t]+");

	/// Whether text has any reference or target (quick: no paragraphs read).
	public static bool Any(string story) => story.Contains("[[xref", StringComparison.Ordinal) || story.Contains("[[fn", StringComparison.Ordinal) || Toc.Any(story);

	/// The document's numbered paragraphs (legal items, numbered styles' paragraphs, ordinary numbered lists), in order.
	public static List<Target> Targets(ITextDocument2 doc) {
		var paragraphs = LegalLists.Paragraphs(doc);
		var numbers = LegalLists.Numbers(paragraphs);
		var targets = new List<Target>();
		// (an ordinary list's numbers above each level, for the full number)
		var above = new string[10];
		Target? parent = null;    // the numbered paragraph a typed sub-point ("e. ") goes under (iteration 50)
		foreach (var p in paragraphs) {
			string? number = null, full = null;
			int textAt = 0;
			bool sub = false;
			if (parent != null && (p.ListType & 0xFFFF) == tomConstants.tomListNone && TypedSub.Match(p.Text) is { Success: true } point
				&& !StyleEngine.NameOf(p.Style).StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) {
				// a sub-point typed in front of the text, under the numbered paragraph before it: "23(e)"
				number = Clean(point.Groups[1].Value);
				full = Joined(new[] { parent.Full, number });
				textAt = point.Length;
				sub = true;
			} else if (numbers.TryGetValue(p.Start, out var n)) {
				int level = Math.Min(p.Level, 9);
				number = Clean(n.Number);
				// a numbered style's levels each have a format of their own ("1." over "(a)": "1(a)"); legal numbers are full
				if (LegalLists.StyleFormat(p.Style, level) is { } format && format != NumberFormats.Legal)
					full = Joined(Enumerable.Range(0, level + 1).Select(k => NumberFormats.Render(LegalLists.StyleFormat(p.Style, k) ?? NumberFormats.Legal, Math.Max(1, n.Counters[k]))));
				else full = number;
				var m = LegalLists.Prefix.Match(p.Text);
				textAt = m.Success ? m.Length : 0;
			} else if ((p.ListType & 0xFFFF) is int kind && kind != tomConstants.tomListNone && kind != tomConstants.tomListBullet && p.ListType > 0) {
				// an ordinary list: RichEdit's own number, as it shows it
				string shown = doc.Range2(p.Start, p.End).GetText2(tomConstants.tomIncludeNumbering) ?? "";
				int tab = shown.IndexOf('\t');
				number = Clean(tab > 0 ? shown[..tab] : "");
				if (number == "") continue;
				int level = Math.Min(p.Level, 9);
				above[level] = number;
				for (int k = level + 1; k < above.Length; k++) above[k] = "";
				full = Joined(Enumerable.Range(0, level + 1).Select(k => above[k] ?? ""));
			} else if (LegalLists.Prefix.Match(p.Text) is { Success: true } typed) {
				// a number typed in front of the text, and a tab ("4.<tab>"), as a list can come back from another program
				number = full = Clean(typed.Value.TrimEnd('\t'));
				textAt = typed.Length;
			}
			if (number == null || number == "") {
				if (AnyMark.Replace(p.Text, "").Trim() != "") parent = null;    // (other text ends a paragraph's sub-points; blank lines don't)
				continue;
			}
			string text = AnyMark.Replace(p.Text[Math.Min(textAt, p.Text.Length)..], "").Trim();
			var mark = TargetMark.Match(p.Text);
			targets.Add(new Target(p.Start, number, full ?? number, text, mark.Success ? mark.Groups[1].Value : null));
			if (!sub) parent = targets[^1];
		}
		return targets;
	}

	/// The target's name, its marker put in first if it has none.
	public static string EnsureTarget(ITextDocument2 doc, Target target) {
		if (target.Name != null) return target.Name;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		string name;
		do name = "_Ref" + Random.Shared.Next(100_000_000, 999_999_999); while (story.Contains(name, StringComparison.Ordinal));
		var paragraph = doc.Range2(target.Start, target.Start);
		paragraph.Expand(tomConstants.tomParagraph);
		Hidden(doc, MarkAt(paragraph), "[[xref:" + name + "]]");
		return name;
	}

	// Where a paragraph's target marker goes: after its number where that is text ("2.1<tab>"); at the end of its text
	// in a list RichEdit numbers itself (RichEdit draws the number with the paragraph's first character, and a hidden
	// one hides it: none shown); else at its start.
	internal static int MarkAt(ITextRange2 paragraph) {
		string text = paragraph.Text ?? "";
		var m = LegalLists.Prefix.Match(text.TrimEnd('\r'));
		if (m.Success) return paragraph.Start + m.Length;
		if (TypedSub.Match(text) is { Success: true } point) return paragraph.Start + point.Length;    // (a typed sub-point: after its letter)
		int type = paragraph.Para.ListType;
		if (type > 0 && (type & 0xFFFF) != tomConstants.tomListNone && !LegalLists.IsLegal(type))
			return paragraph.Start + text.TrimEnd('\r').Length;
		return paragraph.Start;
	}

	// Puts text at a position, hidden; returns where it ends.
	internal static int Hidden(ITextDocument2 doc, int at, string text) {
		var range = doc.Range2(at, at);
		range.Text = text;
		range.SetRange(at, at + text.Length);
		range.Font.Hidden = tomConstants.tomTrue;
		return at + text.Length;
	}

	/// What a reference to a target shows: its number, or (how = "w") its full number.
	public static string Shown(Target? target, string how) => target == null ? Missing : how == "w" ? target.Full : target.Number;

	/// Inserts a reference to a target at a range (replacing its text); returns where it ends.
	public static int Insert(ITextDocument2 doc, ITextRange2 at, Target target, string how) {
		string name = EnsureTarget(doc, target);
		if (at.End > at.Start) at.Text = "";
		int start = at.Start;
		// (the target's marker can have gone in before the range: the range moved with the text)
		int end = Hidden(doc, start, "[[xref=" + name + "|" + how + "]]");
		string shown = Shown(target, how);
		var number = doc.Range2(end, end);
		number.Text = shown;
		number.SetRange(end, end + shown.Length);
		number.Font.Hidden = tomConstants.tomFalse;
		return Hidden(doc, end + shown.Length, "[[/xref]]");
	}

	/// Brings every reference's number up to date (with the targets' numbers now). Returns how many changed.
	public static int Update(ITextDocument2 doc) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (!story.Contains("[[xref=", StringComparison.Ordinal)) return 0;
		var byName = new Dictionary<string, Target>();
		foreach (var t in Targets(doc)) if (t.Name != null) byName.TryAdd(t.Name, t);
		var marked = TargetMark.Matches(story).Select(m => m.Groups[1].Value).ToHashSet(StringComparer.Ordinal);
		int changed = 0;
		foreach (Match m in Reference.Matches(story).Reverse()) {
			string name = m.Groups[1].Value;
			// (a target whose paragraph has no number just now, as while a list is being changed or a file loaded: left as
			// it is; "not found" only once its marker is gone)
			if (!byName.ContainsKey(name) && marked.Contains(name)) continue;
			string wanted = Shown(byName.GetValueOrDefault(name), m.Groups[2].Value);
			if (m.Groups[3].Value == wanted) continue;
			var text = doc.Range2(m.Groups[3].Index, m.Groups[3].Index + m.Groups[3].Length);
			text.Text = wanted;
			text.SetRange(m.Groups[3].Index, m.Groups[3].Index + wanted.Length);
			text.Font.Hidden = tomConstants.tomFalse;    // (an empty one took the marker's format)
			changed++;
		}
		return changed;
	}

	/// The reference at a position (the name of its target), or null.
	public static string? ReferenceAt(string story, int position) {
		foreach (Match m in Reference.Matches(story))
			if (position >= m.Index && position <= m.Index + m.Length) return m.Groups[1].Value;
		return null;
	}

	/// Where a target's paragraph is (its marker), or -1.
	public static int TargetPosition(string story, string name) {
		foreach (Match m in TargetMark.Matches(story)) if (m.Groups[1].Value == name) return m.Index;
		return -1;
	}

	/// After loading (HTML, Markdown, Word's RTF): the markers hidden, and a target's marker put after its paragraph's
	/// number where it went in front of it. Returns how many markers there are.
	public static int AfterLoad(ITextDocument2 doc) {
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (!Any(story)) return 0;
		if (Footnotes.Any(story)) { Footnotes.AfterLoad(doc); story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? ""; }
		var marks = AnyMark.Matches(story);
		foreach (Match m in marks) {
			var range = doc.Range2(m.Index, m.Index + m.Length);
			if (range.Font.Hidden != tomConstants.tomTrue) range.Font.Hidden = tomConstants.tomTrue;
		}
		// a target at a paragraph's start: where MarkAt puts it (after "1.2<tab>", which Renumber writes at the start; at
		// the end in a list RichEdit numbers, whose number it would hide)
		foreach (Match m in TargetMark.Matches(story).Reverse()) {
			var paragraph = doc.Range2(m.Index, m.Index);
			paragraph.StartOf(tomConstants.tomParagraph, 0);
			if (paragraph.Start != m.Index) continue;
			doc.Range2(m.Index, m.Index + m.Length).Text = "";
			paragraph = doc.Range2(m.Index, m.Index);
			paragraph.Expand(tomConstants.tomParagraph);
			Hidden(doc, MarkAt(paragraph), m.Value);
		}
		Toc.AfterLoad(doc);    // (iteration 52)
		return marks.Count;
	}

	/// Text with spaces for the markers (every position as it was).
	public static string Blanked(string text) => Any(text) ? AnyMark.Replace(text, m => new string(' ', m.Length)) : text;

	/// Text without the markers (plain text: the numbers only).
	public static string PlainText(string text) => Any(text) ? AnyMark.Replace(text, "") : text;

	// ================================================================ RTF (Word's bookmarks and REF fields)

	/// RTF as RichEdit wrote it, with Word's bookmarks and REF fields for the markers.
	public static string ToWordRtf(string rtf) {
		if (!Any(rtf)) return rtf;
		rtf = Toc.ToWordRtf(Toc.SplitHidden(rtf));    // (iteration 52: the tables of contents; each marker in a hidden run of its own)
		const string Words = @"((?:\\[a-z]+-?\d* ?)*)", V = @"\\v(?![a-z0-9]) ?";
		// a reference: its number in the field's result (its format inside; the hidden state as it was, after the group)
		rtf = Regex.Replace(rtf, V + Words + @"\[\[xref=(" + Id + @")(?:\|(\w*))?\]\]" + Words + @"\\v0 ?([^{}]*?)" + V + Words + @"\[\[/xref\]\]" + Words + @"\\v0 ?",
			m => m.Groups[1].Value + m.Groups[4].Value + @"{\field{\*\fldinst REF " + m.Groups[2].Value + (m.Groups[3].Value == "w" ? @" \\w" : @" \\r") + @" \\h }{\fldrslt "
				+ m.Groups[5].Value + "}}" + Footnotes.Delimited(m.Groups[6].Value + m.Groups[7].Value));
		rtf = Regex.Replace(rtf, V + Words + @"\[\[xref:(" + Id + @")\]\]" + Words + @"\\v0 ?",
			m => m.Groups[1].Value + @"{\*\bkmkstart " + m.Groups[2].Value + @"}{\*\bkmkend " + m.Groups[2].Value + "}" + Footnotes.Delimited(m.Groups[3].Value));
		return rtf;
	}

	/// Word's (and LibreOffice's) RTF with markers for its REF fields and the bookmarks they refer to, before RichEdit
	/// loads it (it drops both).
	public static string FromWordRtf(string rtf) {
		rtf = Toc.FromWordRtf(rtf);    // (iteration 52: Word's TOC fields)
		if (!rtf.Contains(@"\fldinst", StringComparison.Ordinal) || !rtf.Contains("REF", StringComparison.Ordinal)) return rtf;
		var names = new HashSet<string>(StringComparer.Ordinal);
		var sb = new StringBuilder();
		int at = 0;
		for (int i = rtf.IndexOf(@"{\field", StringComparison.Ordinal); i >= 0; i = rtf.IndexOf(@"{\field", i + 1, StringComparison.Ordinal)) {
			if (i < at || (i + 7 < rtf.Length && char.IsLetter(rtf[i + 7]))) continue;
			int end = GroupEnd(rtf, i);
			if (end < 0) break;
			string field = rtf[i..(end + 1)];
			int inst = field.IndexOf(@"\fldinst", StringComparison.Ordinal), result = field.IndexOf(@"{\fldrslt", StringComparison.Ordinal);
			if (inst < 0 || result < inst) continue;
			string instruction = Regex.Replace(field[(inst + 8)..result].Replace("\r", "").Replace("\n", "").Replace(@"\\", "\u0001"), @"\\[a-z]+-?\d* ?|[{}]|\\\*", "").Replace('\u0001', '\\');
			var reference = Regex.Match(instruction, @"^\s*REF\s+(" + Id + @")(.*)$", RegexOptions.Singleline);
			if (!reference.Success) continue;
			int resultEnd = GroupEnd(field, result);
			if (resultEnd < 0) continue;
			string shown = field[(result + 9)..resultEnd];
			if (shown.StartsWith(' ')) shown = shown[1..];    // (the space ending "\fldrslt" isn't text)
			string name = reference.Groups[1].Value;
			names.Add(name);
			string how = Regex.IsMatch(reference.Groups[2].Value, @"\\w\b") ? "w" : "r";
			sb.Append(rtf, at, i - at);
			sb.Append(@"{\v [[xref=").Append(name).Append('|').Append(how).Append(@"]]}{").Append(shown).Append(@"}{\v [[/xref]]}");
			at = end + 1;
		}
		if (names.Count == 0) return rtf;
		sb.Append(rtf, at, rtf.Length - at);
		rtf = sb.ToString();
		// the bookmarks referred to: a target's marker where they start
		rtf = Regex.Replace(rtf, @"\{\\\*\\bkmkstart (" + Id + @")\}", m => names.Contains(m.Groups[1].Value) ? @"{\v [[xref:" + m.Groups[1].Value + "]]}" : m.Value);
		return Regex.Replace(rtf, @"\{\\\*\\bkmkend (" + Id + @")\}", m => names.Contains(m.Groups[1].Value) ? "" : m.Value);
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

	// ================================================================ .docx (the built-in reader, which goes through HTML)

	/// A .docx with markers (as text) for its REF fields and the bookmarks they refer to, for a reader that drops both:
	/// the path of a marked copy in folder, or the file itself if it has no REF fields.
	public static string MarkedDocx(string path, string folder) {
		string copy = Path.Combine(folder, "xrefs-" + Path.GetFileName(path));
		File.Copy(path, copy, true);
		bool any = false;
		using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(copy, true)) {
			var body = package.MainDocumentPart?.Document?.Body;
			if (body == null) return path;
			var names = new HashSet<string>(StringComparer.Ordinal);
			bool toc = Toc.MarkDocx(package);    // (iteration 52: its tables of contents, first: their fields go across paragraphs)
			static W.Run Marker(string text) => new(new W.Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
			// simple fields: <w:fldSimple w:instr="REF _Ref1 \h">result</w:fldSimple>
			foreach (var simple in body.Descendants<W.SimpleField>().ToList()) {
				var m = Regex.Match(simple.Instruction?.Value ?? "", @"^\s*REF\s+(" + Id + @")(.*)$");
				if (!m.Success) continue;
				names.Add(m.Groups[1].Value);
				string how = Regex.IsMatch(m.Groups[2].Value, @"\\w\b") ? "w" : "r";
				var parent = simple.Parent!;
				parent.InsertBefore(Marker($"[[xref={m.Groups[1].Value}|{how}]]"), simple);
				foreach (var child in simple.ChildElements.ToList()) { child.Remove(); parent.InsertBefore(child, simple); }
				parent.InsertBefore(Marker("[[/xref]]"), simple);
				simple.Remove();
			}
			// complex fields: begin, instruction, separate, result, end (each a run's w:fldChar)
			foreach (var paragraph in body.Descendants<W.Paragraph>().ToList()) {
				var runs = paragraph.Descendants<W.Run>().ToList();
				for (int i = 0; i < runs.Count; i++) {
					if (runs[i].GetFirstChild<W.FieldChar>()?.FieldCharType?.Value != W.FieldCharValues.Begin) continue;
					int depth = 0, separate = -1, end = -1;
					var instruction = new StringBuilder();
					for (int k = i; k < runs.Count && end < 0; k++) {
						var type = runs[k].GetFirstChild<W.FieldChar>()?.FieldCharType?.Value;
						if (type == W.FieldCharValues.Begin) depth++;
						else if (type == W.FieldCharValues.End && --depth == 0) end = k;
						else if (type == W.FieldCharValues.Separate && depth == 1) separate = k;
						else if (depth == 1 && separate < 0) foreach (var code in runs[k].Elements<W.FieldCode>()) instruction.Append(code.Text);
					}
					if (end < 0) break;
					var m = Regex.Match(instruction.ToString(), @"^\s*REF\s+(" + Id + @")(.*)$");
					if (!m.Success) { i = end; continue; }
					names.Add(m.Groups[1].Value);
					string how = Regex.IsMatch(m.Groups[2].Value, @"\\w\b") ? "w" : "r";
					// the begin run becomes the reference's start, the end run its end; the instruction goes
					runs[i].RemoveAllChildren();
					runs[i].Append(new W.Text($"[[xref={m.Groups[1].Value}|{how}]]") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
					for (int k = i + 1; k <= (separate < 0 ? end - 1 : separate); k++) runs[k].Remove();
					runs[end].RemoveAllChildren();
					runs[end].Append(new W.Text("[[/xref]]"));
					i = end;
				}
			}
			// links to a _Ref bookmark (the built-in converter's, from the HTML's <a href="#_Ref...">): references too
			foreach (var link in body.Descendants<W.Hyperlink>().ToList()) {
				if (link.Anchor?.Value is not { } anchor || !anchor.StartsWith("_Ref", StringComparison.Ordinal)) continue;
				names.Add(anchor);
				var parent = link.Parent!;
				parent.InsertBefore(Marker($"[[xref={anchor}|{(link.Tooltip?.Value == "xref w" ? "w" : "r")}]]"), link);
				foreach (var child in link.ChildElements.ToList()) { child.Remove(); parent.InsertBefore(child, link); }
				parent.InsertBefore(Marker("[[/xref]]"), link);
				link.Remove();
			}
			foreach (var bookmark in body.Descendants<W.BookmarkStart>().ToList())
				if (bookmark.Name?.Value is { } name && names.Contains(name)) bookmark.InsertAfterSelf(Marker($"[[xref:{name}]]"));
			bool notes = Footnotes.MarkDocx(package);    // (iteration 51: its footnotes too)
			any = names.Count > 0 || notes || toc;
			if (any) package.MainDocumentPart!.Document.Save();
		}
		return any ? copy : path;
	}

	/// Bookmarks for the targets in a .docx made from the HTML (whose converter drops empty anchors): each at the start
	/// of the first paragraph, after the one before's, whose text has the target's (name, the paragraph's text after its
	/// number).
	public static void BookmarksIntoDocx(string path, List<(string Name, string Text)> targets) {
		if (targets.Count == 0) return;
		using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, true);
		var body = package.MainDocumentPart?.Document?.Body;
		if (body == null) return;
		var have = body.Descendants<W.BookmarkStart>().Select(b => b.Name?.Value).ToHashSet();
		int id = body.Descendants<W.BookmarkStart>().Select(b => int.TryParse(b.Id?.Value, out int n) ? n : 0).DefaultIfEmpty(0).Max() + 1;
		var paragraphs = body.Descendants<W.Paragraph>().ToList();
		static string Plain(string s) => Regex.Replace(s, @"\s+", " ").Trim();
		int from = 0;
		foreach (var (name, text) in targets) {
			if (have.Contains(name)) continue;
			string wanted = Plain(text);
			if (wanted.Length > 40) wanted = wanted[..40];
			int found = wanted == "" ? -1 : paragraphs.FindIndex(from, p => Plain(p.InnerText).Contains(wanted, StringComparison.Ordinal));
			if (found < 0) continue;
			var paragraph = paragraphs[found];
			// (a table of contents' entry has its heading's text: the heading is further on)
			if (paragraph.Descendants<W.Hyperlink>().Any(l => l.Tooltip?.Value?.StartsWith("toc:", StringComparison.Ordinal) == true)) {
				found = paragraphs.FindIndex(found + 1, p => Plain(p.InnerText).Contains(wanted, StringComparison.Ordinal)
					&& !p.Descendants<W.Hyperlink>().Any(l => l.Tooltip?.Value?.StartsWith("toc:", StringComparison.Ordinal) == true));
				if (found < 0) continue;
				paragraph = paragraphs[found];
			}
			var start = new W.BookmarkStart { Name = name, Id = id.ToString() };
			var end = new W.BookmarkEnd { Id = id.ToString() };
			id++;
			if (paragraph.ParagraphProperties is { } properties) properties.InsertAfterSelf(start); else paragraph.PrependChild(start);
			start.InsertAfterSelf(end);
			from = found + 1;
		}
		package.MainDocumentPart!.Document.Save();
	}

	// ================================================================ HTML and Markdown

	/// HTML from the editor with anchors for the targets and links for the references.
	public static string ToHtml(string html) {
		if (!Any(html)) return html;
		html = Toc.ToHtml(html);    // (iteration 52)
		// (the full number's "w" also in the title: the built-in .docx converter keeps it as the link's tooltip)
		html = Reference.Replace(html, m => m.Groups[2].Value == "w" ? $"<a href=\"#{m.Groups[1].Value}\" data-xref=\"w\" title=\"xref w\">{m.Groups[3].Value}</a>"
			: $"<a href=\"#{m.Groups[1].Value}\" data-xref=\"r\">{m.Groups[3].Value}</a>");
		return TargetMark.Replace(html, m => $"<a id=\"{m.Groups[1].Value}\"></a>");
	}

	/// HTML with markers for the anchors and links ToHtml wrote (and Word's _Ref anchors), before it is loaded.
	public static string FromHtml(string html) {
		html = Toc.FromHtml(html);    // (iteration 52)
		if (!html.Contains("<a", StringComparison.OrdinalIgnoreCase)) return html;
		var names = new HashSet<string>(StringComparer.Ordinal);
		html = Regex.Replace(html, @"<a\s+href=""#(" + Id + @")""(?:\s+data-xref=""(\w*)"")?(?:\s+title=""xref (\w*)"")?\s*>(.*?)</a>", m => {
			bool ours = m.Groups[2].Success || m.Groups[3].Success || m.Groups[1].Value.StartsWith("_Ref", StringComparison.Ordinal);
			if (!ours) return m.Value;
			names.Add(m.Groups[1].Value);
			string how = m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Success ? m.Groups[3].Value : "r";
			return $"[[xref={m.Groups[1].Value}|{how}]]{m.Groups[4].Value}[[/xref]]";
		}, RegexOptions.IgnoreCase | RegexOptions.Singleline);
		if (names.Count == 0) return html;
		return Regex.Replace(html, @"<a\s+(?:id|name)=""(" + Id + @")""\s*>\s*</a>", m => names.Contains(m.Groups[1].Value) ? $"[[xref:{m.Groups[1].Value}]]" : m.Value, RegexOptions.IgnoreCase);
	}

	// Markdown escapes the markers' brackets and underscores: "\[\[xref:\_Ref1\]\]".
	const string E = @"\\?";
	static readonly Regex MarkdownTarget = new(E + @"\[" + E + @"\[xref:((?:\\?[A-Za-z0-9_]){1,40})" + E + @"\]" + E + @"\]");
	static readonly Regex MarkdownReference = new(E + @"\[" + E + @"\[xref=((?:\\?[A-Za-z0-9_]){1,40})(?:" + E + @"\|(\w*))?" + E + @"\]" + E + @"\](.*?)" + E + @"\[" + E + @"\[/xref" + E + @"\]" + E + @"\]");

	/// Markdown from the editor with <a id> anchors for the targets and [number](#name) links for the references.
	public static string ToMarkdown(string markdown) {
		if (!Any(markdown.Replace("\\", ""))) return markdown;
		markdown = Toc.ToMarkdown(markdown);    // (iteration 52)
		markdown = MarkdownReference.Replace(markdown, m => {
			string name = m.Groups[1].Value.Replace("\\", "");
			return $"[{m.Groups[3].Value}](#{name}{(m.Groups[2].Value == "w" ? " \"xref w\"" : "")})";
		});
		return MarkdownTarget.Replace(markdown, m => $"<a id=\"{m.Groups[1].Value.Replace("\\", "")}\"></a>");
	}

	/// Markdown with markers for its anchors and links to them, before it is loaded (escaped, so that Markdown keeps them
	/// as text).
	public static string FromMarkdown(string markdown) {
		markdown = Toc.FromMarkdown(markdown);    // (iteration 52)
		if (!markdown.Contains("](#", StringComparison.Ordinal)) return markdown;
		var names = new HashSet<string>(StringComparer.Ordinal);
		markdown = Regex.Replace(markdown, @"\[([^\]\r\n]*)\]\(#(" + Id + @")(?:\s+""xref (\w*)"")?\)", m => {
			if (!m.Groups[3].Success && !m.Groups[2].Value.StartsWith("_Ref", StringComparison.Ordinal)) return m.Value;
			names.Add(m.Groups[2].Value);
			return $@"\[\[xref={Escaped(m.Groups[2].Value)}\|{(m.Groups[3].Success ? m.Groups[3].Value : "r")}\]\]{m.Groups[1].Value}\[\[/xref\]\]";
		});
		if (names.Count == 0) return markdown;
		return Regex.Replace(markdown, @"<a\s+(?:id|name)=""(" + Id + @")""\s*>\s*</a>", m => names.Contains(m.Groups[1].Value) ? $@"\[\[xref:{Escaped(m.Groups[1].Value)}\]\]" : m.Value, RegexOptions.IgnoreCase);
	}

	static string Escaped(string name) => name.Replace("_", "\\_");
}
