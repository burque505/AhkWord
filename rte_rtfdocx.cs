// (iteration 87) RtfDocx: the built-in .docx writer that needs no Word. It reads the RTF AhkWord gives Word (RichEdit's
// RTF with AhkWord's styles, legal lists, fields, footnotes, header, footer, sections and line numbering added: the
// "finish" step of OfficeConverters.Export) and writes the .docx from it with the Open XML SDK, the way Word does with
// that RTF: no HTML step, no patching afterwards.
//
// What it follows: RTF's own rules. Character and paragraph properties belong to the group they are set in; \pard and
// \plain reset them; a paragraph's properties are the ones in force at its \par. Each style of the style sheet becomes a
// Word style holding the properties its RTF entry gives; a paragraph and its runs then get only the properties that
// differ from their style (as Word writes them). Lists: the list table's lists (\ls) as Word numbering; RichEdit's own
// (\pn) lists as numbering too, a new instance wherever the number the editor showed ({\pntext}) doesn't follow on.
// Tables: rows by \trowd/\cellx, a grid made of all the rows' cell edges, merged cells, borders, shading. Fields stay
// fields (TOC, PAGEREF, PAGE ...); a HYPERLINK inside one paragraph becomes a Word hyperlink. Footnotes and endnotes,
// headers and footers (per section), sections (page size, margins, columns, breaks, first page), bookmarks, pictures
// (PNG, JPEG, EMF, WMF) with their alt text.

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

public static class RtfDocx {
	/// Writes rtf as a .docx at output; returns output.
	public static string Write(string rtf, string output) {
		if (Environment.GetEnvironmentVariable("RT_DUMPRTF") is { Length: > 0 } dump) File.WriteAllText(dump, rtf, Encoding.Latin1);
		var root = Parse(rtf);
		using (var package = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document))
			new Writer(package).Run(root);
		return output;
	}

	// ================================================================ the RTF as a tree of groups

	abstract class Item { }
	sealed class Grp : Item { public readonly List<Item> Kids = new(); }
	sealed class Cw : Item { public string Word = ""; public int? N; }
	sealed class Txt : Item { public string S = ""; }
	sealed class Hex : Item { public byte B; }
	sealed class Uni : Item { public int Code; }
	sealed class Bin : Item { public byte[] Data = Array.Empty<byte>(); }

	static Grp Parse(string s) {
		var root = new Grp();
		var stack = new Stack<(Grp Group, int Uc)>();
		Grp cur = root;
		int uc = 1, skip = 0, i = 0, n = s.Length;
		var text = new StringBuilder();
		void Flush() { if (text.Length > 0) { cur.Kids.Add(new Txt { S = text.ToString() }); text.Clear(); } }
		void Add(Item item) { if (skip > 0) { skip--; return; } Flush(); cur.Kids.Add(item); }
		while (i < n) {
			char c = s[i];
			if (c == '{') { Flush(); skip = 0; var g = new Grp(); cur.Kids.Add(g); stack.Push((cur, uc)); cur = g; i++; continue; }
			if (c == '}') { Flush(); skip = 0; if (stack.Count == 0) break; (cur, uc) = stack.Pop(); i++; continue; }
			if (c == '\r' || c == '\n') { i++; continue; }
			if (c != '\\') {
				if (skip > 0) { skip--; i++; continue; }
				if (c >= 0x80) { Flush(); cur.Kids.Add(new Hex { B = (byte)c }); }
				else text.Append(c);
				i++; continue;
			}
			// a control word or symbol
			if (i + 1 >= n) break;
			char d = s[i + 1];
			if (char.IsAsciiLetter(d)) {
				int start = i + 1, j = start;
				while (j < n && char.IsAsciiLetter(s[j])) j++;
				string word = s[start..j];
				int? param = null;
				if (j < n && (s[j] == '-' || char.IsAsciiDigit(s[j]))) {
					int k = j + 1;
					while (k < n && char.IsAsciiDigit(s[k])) k++;
					if (int.TryParse(s.AsSpan(j, k - j), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int v)) param = v;
					else param = s[j] == '-' ? int.MinValue : int.MaxValue;
					j = k;
				}
				if (j < n && s[j] == ' ') j++;
				i = j;
				if (word == "bin" && param is int count and > 0) {
					var data = Encoding.Latin1.GetBytes(s.Substring(i, Math.Min(count, n - i)));
					i += count;
					Add(new Bin { Data = data });
					continue;
				}
				if (word == "uc") { uc = Math.Max(0, param ?? 1); continue; }
				if (word == "u") {
					if (skip > 0) { skip--; continue; }
					Flush();
					int code = param ?? 0;
					if (code < 0) code += 65536;
					cur.Kids.Add(new Uni { Code = code });
					skip = uc;
					continue;
				}
				Add(new Cw { Word = word, N = param });
				continue;
			}
			if (d == '\'') {
				i += 2;
				byte b = 0;
				if (i + 2 <= n && byte.TryParse(s.AsSpan(i, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var hb)) { b = hb; i += 2; }
				Add(new Hex { B = b });
				continue;
			}
			i += 2;
			if (d == '\\' || d == '{' || d == '}') {
				if (skip > 0) { skip--; continue; }
				text.Append(d);
				continue;
			}
			if (d == '\r' || d == '\n') { Add(new Cw { Word = "par" }); continue; }
			Add(new Cw { Word = d.ToString() });
		}
		Flush();
		return root.Kids.OfType<Grp>().FirstOrDefault() ?? root;
	}

	static string? DestOf(Grp g) {
		if (g.Kids.Count == 0) return null;
		if (g.Kids[0] is Cw { Word: "*" }) return g.Kids.Count > 1 && g.Kids[1] is Cw w ? w.Word : "*";
		return g.Kids[0] is Cw first ? first.Word : null;
	}

	static bool Starred(Grp g) => g.Kids.Count > 0 && g.Kids[0] is Cw { Word: "*" };

	// ================================================================ formats

	sealed class Border {
		public string Style = "single"; public int Width = 10; public int Color = -1; public int Space;
		public Border Clone() => (Border)MemberwiseClone();
		public string Key => $"{Style},{Width},{Color},{Space}";
	}

	sealed class CharFmt {
		public int Font = -1, Size = 24, Color, Highlight, Back = -1, Vert, Spacing, Scale = 100, Position;
		public bool B, I, Strike, DStrike, Caps, SCaps, Hidden, Outline, Shadow, Emboss, Imprint;
		public string Ul = "";
		public CharFmt Clone() => (CharFmt)MemberwiseClone();
		public string Key => $"{Font}|{Size}|{Color}|{Highlight}|{Back}|{Vert}|{Spacing}|{Scale}|{Position}|{B}{I}{Strike}{DStrike}{Caps}{SCaps}{Hidden}{Outline}{Shadow}{Emboss}{Imprint}|{Ul}";
	}

	sealed class TabStop { public int Pos; public char Kind = 'l'; public char Leader = 'n'; public string Key => $"{Pos}{Kind}{Leader}"; }

	sealed class PnDef {
		public string Kind = "body", Fmt = "dec", Before = "", After = "";
		public int Start = 1, Font = -1;
		public string Key(Func<int, string> fontName) => $"{Kind}|{Fmt}|{Before}|{After}|{fontName(Font)}";
	}

	sealed class ParaFmt {
		public int Style, Left, Right, First, Before, After, Sl, Ls, Ilvl, Outline = -1, Shading = -1;
		public int Suffix;    // (iteration 87: AhkWord's \rtesuffix, what follows the list number: 1 a space, 2 nothing)
		public bool SlMult, Keep, KeepNext, PageBefore, Widow, Contextual, InTable;
		public int Itap;    // (the table level: 1 a table, 2 a table nested in its cell, ...; 0 with \intbl means 1)
		public char Align = 'l';
		public List<TabStop> Tabs = new();
		public char TabKind = 'l', TabLeader = 'n';
		public PnDef? Pn;
		public Border?[] Borders = new Border?[4];    // top, left, bottom, right
		public ParaFmt Clone() {
			var p = (ParaFmt)MemberwiseClone();
			p.Tabs = new List<TabStop>(Tabs);
			p.Borders = (Border?[])Borders.Clone();
			return p;
		}
	}

	sealed class CellDef {
		public int Right, CbPat = -1, CfPat = -1, Shading;
		public char VAlign = 't';
		public string? VMerge;
		public bool HMergeFirst, HMerge;
		public Border?[] Borders = new Border?[4];    // top, left, bottom, right
	}

	sealed class RowDef {
		public int Gap, Left, Height;
		public int? PadL, PadR;
		public char Align = 'l';
		public bool Header, Keep;
		public List<CellDef> Cells = new();
		public RowDef Snapshot() { var r = (RowDef)MemberwiseClone(); r.Cells = new List<CellDef>(Cells); return r; }
	}

	sealed class SectFmt {
		public int PgW = 12240, PgH = 15840, ML = 1800, MR = 1800, MT = 1440, MB = 1440, HeaderY = 720, FooterY = 720, Gutter;
		public bool Landscape, TitlePg, ColLine, PgnRestart;
		public string Break = "page";
		public int Cols = 1, ColSpace = 720, PgnStart = 1, LineMod, LineX, LineStart = 1;
		public string? PgnFmt, LineRestart, VAlign;
		public List<(bool Header, string Type, string Id)> Refs = new();
		public SectFmt Clone() { var s = (SectFmt)MemberwiseClone(); s.Refs = new(); return s; }
	}

	sealed class StyleDef {
		public int Number; public string Name = ""; public string Id = ""; public int BasedOn = -1, Next = -1;
		public ParaFmt P = new(); public CharFmt C = new();
	}

	sealed class FontInfo { public string Name = ""; public int Charset = 0; public int CodePage = 1252; }

	sealed class Note { public long Id; public bool End; }

	sealed class PicData { public byte[] Data = Array.Empty<byte>(); public string Type = "png"; public long Cx, Cy; public string Alt = ""; }

	// ---- the paragraph being written: its pieces in order
	enum K { Run, FldBegin, FldSep, FldEnd, BmStart, BmEnd, HlStart, HlEnd, NoteRef, NoteMark, Pic }
	sealed class Inline {
		public K Kind; public CharFmt Fmt = null!; public string Key = "";
		public readonly List<(char Type, string Text)> Pieces = new();    // t text, T tab, b line break, p page break, c column break, h no-break hyphen, s soft hyphen
		public string Text = ""; public bool Lock; public int Id; public Note? Note; public PicData? Pic; public string? Url, Anchor;
	}

	sealed class Story {
		public OpenXmlPart Part = null!;
		public readonly List<OpenXmlElement> Blocks = new();
		public readonly List<Inline> Line = new();
		public string? PnText;
		// the tables being written: [0] the table in the text, [1] a table nested in its cell, ...
		public readonly List<TableLevel> Levels = new() { new TableLevel() };
	}

	sealed class TableLevel {
		public readonly List<(RowDef Def, List<List<OpenXmlElement>> Cells)> Rows = new();
		public List<List<OpenXmlElement>> Cells = new();
		public List<OpenXmlElement> Cell = new();
		public bool Open => Rows.Count > 0 || Cells.Count > 0 || Cell.Count > 0;
	}

	// ================================================================ the writer

	sealed class Writer {
		readonly WordprocessingDocument package;
		readonly MainDocumentPart main;
		Story story = null!;
		CharFmt c = new();
		ParaFmt p = new();
		SectFmt docSect = new(), sect = new();
		RowDef row = new();
		CellDef cell = new();
		Border? border;    // the border the next \brdr... words describe
		readonly List<FontInfo?> fonts = new();
		readonly List<string?> colors = new();
		readonly Dictionary<int, StyleDef> styles = new();
		int defaultFont, ansiCodePage = 1252, defTab = 720;
		bool facing;
		// lists
		readonly List<W.AbstractNum> abstracts = new();
		readonly Dictionary<int, int> abstractOfListId = new();
		readonly Dictionary<(int Abstract, int Level), List<int>> levelTabs = new();    // (the list levels' number tabs)
		readonly Dictionary<int, int> abstractOfNum = new();
		readonly List<W.NumberingInstance> nums = new();
		readonly HashSet<int> numIds = new();
		readonly Dictionary<string, int> pnAbstracts = new();
		(string Key, int NumId, int Next)? lastPn;
		int nextPnNum = 1000;
		// notes, bookmarks, pictures
		FootnotesPart? footnotes;
		EndnotesPart? endnotes;
		long nextFootnote = 1, nextEndnote = 1;
		Note? pendingNote;
		readonly Dictionary<string, int> openBookmarks = new();
		int nextBookmark = 1;
		uint nextPicture = 1;

		public Writer(WordprocessingDocument package) {
			this.package = package;
			main = package.AddMainDocumentPart();
			main.Document = new W.Document(new W.Body());
		}

		public void Run(Grp root) {
			try { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); } catch { }
			story = new Story { Part = main };
			sect = docSect.Clone();
			WalkKids(root);
			FinishStory(story);
			var body = main.Document.Body!;
			foreach (var block in story.Blocks) body.Append(block);
			if (!story.Blocks.Any(b => b is W.Paragraph) || story.Blocks.LastOrDefault() is W.Table) body.Append(new W.Paragraph());
			body.Append(SectPr(sect));
			WriteStyles();
			WriteNumbering();
			WriteSettings();
			if (footnotes != null) footnotes.Footnotes.Save();
			if (endnotes != null) endnotes.Endnotes.Save();
			main.Document.Save();
		}

		// ---- walking the tree

		void Walk(Grp g) {
			var (savedC, savedP) = (c, p);
			c = c.Clone(); p = p.Clone();
			try { WalkGroup(g); }
			finally { (c, p) = (savedC, savedP); }
		}

		// a group's content in a state of its own (footnotes, headers: their handler is the caller)
		// (Its last paragraph, without a \par of its own, ends here with the group's formats, not after them with the
		// text's: a footnote's last paragraph took the body paragraph's spacing and its superscript)
		void WalkInside(Grp g) {
			var (savedC, savedP) = (c, p);
			c = c.Clone(); p = p.Clone();
			try {
				WalkKids(g);
				if (story.Line.Count > 0) EndParagraph();
			}
			finally { (c, p) = (savedC, savedP); }
		}

		void WalkGroup(Grp g) {
			string? dest = DestOf(g);
			switch (dest) {
				case "fonttbl": FontTable(g); return;
				case "colortbl": ColorTable(g); return;
				case "stylesheet": StyleSheet(g); return;
				case "listtable": ListTable(g); return;
				case "listoverridetable": ListOverrides(g); return;
				case "pn": PnGroup(g); return;
				case "pntext": story.PnText = PlainText(g); return;
				case "field": Field(g); return;
				case "pict": Picture(g); return;
				case "shppict": WalkKids(g); return;
				case "footnote": Footnote(g); return;
				case "header": case "headerl": case "headerr": case "headerf":
				case "footer": case "footerl": case "footerr": case "footerf":
					HeaderFooter(g, dest); return;
				case "bkmkstart": BookmarkStart(PlainText(g).Trim()); return;
				case "bkmkend": BookmarkEnd(PlainText(g).Trim()); return;
				case "nesttableprops": {
					// a nested table's row: its \trowd...\cellx and \nestrow, kept apart from the outer row's definition
					var saved = (row, cell);
					row = new RowDef(); cell = new CellDef();
					try { WalkKids(g); }
					finally { (row, cell) = saved; }
					return;
				}
				case "nonesttables": return;    // (the text for readers without nested tables)
				case "info": case "listtext": case "nonshppict": case "fldinst": case "ftnsep": case "ftnsepc": case "aftnsep": case "aftnsepc":
				case "revtbl": case "rsidtbl": case "filetbl": case "xmlnstbl": case "object": case "colorschememapping": case "themedata":
				case "latentstyles": case "datastore": case "pgdsctbl": case "generator": case "mmathPr": case "userprops": case "docvar":
					return;
			}
			if (Starred(g)) return;    // (a destination this writer doesn't know: RTF says skip it)
			WalkKids(g);
		}

		void WalkKids(Grp g) {
			var bytes = new List<byte>();
			void FlushBytes() { if (bytes.Count > 0) { Emit(Decode(bytes.ToArray(), c.Font)); bytes.Clear(); } }
			foreach (var item in g.Kids) {
				if (item is Hex h) { bytes.Add(h.B); continue; }
				FlushBytes();
				switch (item) {
					case Grp sub when DestOf(sub) == "pn": PnGroup(sub); break;    // (the enclosing group's paragraphs)
					case Grp sub: Walk(sub); break;
					case Txt t: Emit(t.S); break;
					case Uni u: Emit(((char)u.Code).ToString()); break;
					case Cw w: Control(w.Word, w.N); break;
				}
			}
			FlushBytes();
		}

		// ---- text

		void Emit(string s) {
			if (s.Length == 0) return;
			if (Charset(c.Font) == 2) s = new string(s.Select(ch => ch < 0x100 ? (char)(0xF000 + ch) : ch).ToArray());
			int from = 0;
			for (int k = 0; k < s.Length; k++) {
				if (s[k] != '\t') continue;
				if (k > from) Piece('t', s[from..k]);
				Piece('T');
				from = k + 1;
			}
			if (from < s.Length) Piece('t', s[from..]);
		}

		void Piece(char type, string text = "") {
			string key = c.Key;
			var line = story.Line;
			if (line.Count > 0 && line[^1] is { Kind: K.Run } last && last.Key == key) {
				if (type == 't' && last.Pieces.Count > 0 && last.Pieces[^1].Type == 't') last.Pieces[^1] = ('t', last.Pieces[^1].Text + text);
				else last.Pieces.Add((type, text));
				return;
			}
			var run = new Inline { Kind = K.Run, Fmt = c.Clone(), Key = key };
			run.Pieces.Add((type, text));
			line.Add(run);
		}

		void Mark(K kind, Action<Inline>? set = null) {
			var inline = new Inline { Kind = kind, Fmt = c.Clone(), Key = c.Key };
			set?.Invoke(inline);
			story.Line.Add(inline);
		}

		string Decode(byte[] data, int font) {
			int cs = Charset(font);
			if (cs == 2) return new string(data.Select(b => (char)b).ToArray());    // (Emit maps a symbol font's bytes to U+F0xx)
			int cp = fonts.ElementAtOrDefault(font < 0 ? defaultFont : font)?.CodePage ?? ansiCodePage;
			try { return Encoding.GetEncoding(cp).GetString(data); }
			catch { return Encoding.GetEncoding(1252).GetString(data); }
		}

		int Charset(int font) => fonts.ElementAtOrDefault(font < 0 ? defaultFont : font)?.Charset ?? 0;
		string FontName(int font) => fonts.ElementAtOrDefault(font < 0 ? defaultFont : font)?.Name ?? "";

		string PlainText(Grp g) {
			var sb = new StringBuilder();
			var bytes = new List<byte>();
			void Rec(Grp group) {
				foreach (var item in group.Kids) {
					if (item is Hex h) { bytes.Add(h.B); continue; }
					if (bytes.Count > 0) { sb.Append(Encoding.GetEncoding(ansiCodePage).GetString(bytes.ToArray())); bytes.Clear(); }
					switch (item) {
						case Txt t: sb.Append(t.S); break;
						case Uni u: sb.Append((char)u.Code); break;
						case Cw { Word: "tab" }: sb.Append('\t'); break;
						case Grp sub when DestOf(sub) is not ("*" or "bkmkstart" or "bkmkend" or "fldinst" or "datafield"): Rec(sub); break;
					}
				}
			}
			Rec(g);
			if (bytes.Count > 0) sb.Append(Encoding.GetEncoding(ansiCodePage).GetString(bytes.ToArray()));
			return sb.ToString();
		}

		// ---- control words in the text

		void Control(string w, int? n) {
			switch (w) {
				case "par": EndParagraph(); return;
				case "sect": EndParagraph(section: true); return;
				case "cell": EndParagraph(cellEnd: true, level: 1); return;
				case "nestcell": EndParagraph(cellEnd: true, level: Math.Max(2, p.Itap)); return;
				case "row": EndRow(1); return;
				case "nestrow": EndRow(Math.Max(2, story.Levels.Count)); return;
				case "pard": p = new ParaFmt(); story.PnText = null; return;
				case "plain": c = new CharFmt(); return;
				case "sectd": sect = docSect.Clone(); return;
				case "trowd": row = new RowDef(); cell = new CellDef(); return;
				case "tab": Piece('T'); return;
				case "line": Piece('b'); return;
				case "page": Piece('p'); return;
				case "column": Piece('c'); return;
				case "~": Piece('t', "\u00A0"); return;
				case "_": Piece('h'); return;
				case "-": Piece('s'); return;
				case "emdash": Piece('t', "\u2014"); return;
				case "endash": Piece('t', "\u2013"); return;
				case "emspace": Piece('t', "\u2003"); return;
				case "enspace": Piece('t', "\u2002"); return;
				case "qmspace": Piece('t', "\u2005"); return;
				case "bullet": Piece('t', "\u2022"); return;
				case "lquote": Piece('t', "\u2018"); return;
				case "rquote": Piece('t', "\u2019"); return;
				case "ldblquote": Piece('t', "\u201C"); return;
				case "rdblquote": Piece('t', "\u201D"); return;
				case "zwj": Piece('t', "\u200D"); return;
				case "zwnj": Piece('t', "\u200C"); return;
				case "chftn":
					if (story.Part is FootnotesPart or EndnotesPart) Mark(K.NoteMark, i => i.Note = new Note { End = story.Part is EndnotesPart });
					else { var note = new Note(); pendingNote = note; Mark(K.NoteRef, i => i.Note = note); }
					return;
				case "chpgn":
					Mark(K.FldBegin, i => i.Text = " PAGE "); Mark(K.FldSep); Piece('t', "1"); Mark(K.FldEnd);
					return;
				// document
				case "deff": defaultFont = n ?? 0; return;
				case "ansicpg": ansiCodePage = n ?? 1252; return;
				case "deftab": defTab = n ?? 720; return;
				case "facingp": facing = true; return;
				case "paperw": docSect.PgW = sect.PgW = n ?? 12240; return;
				case "paperh": docSect.PgH = sect.PgH = n ?? 15840; return;
				case "margl": docSect.ML = sect.ML = n ?? 1800; return;
				case "margr": docSect.MR = sect.MR = n ?? 1800; return;
				case "margt": docSect.MT = sect.MT = n ?? 1440; return;
				case "margb": docSect.MB = sect.MB = n ?? 1440; return;
				case "gutter": docSect.Gutter = sect.Gutter = n ?? 0; return;
				case "landscape": docSect.Landscape = sect.Landscape = true; return;
			}
			if (CharWord(c, w, n) || ParaWord(p, w, n) || SectWord(w, n) || TableWord(w, n) || BorderWord(w, n)) return;
		}

		static bool On(int? n) => n is null or not 0;

		bool CharWord(CharFmt f, string w, int? n) {
			switch (w) {
				case "f": f.Font = n ?? -1; return true;
				case "fs": f.Size = n ?? 24; return true;
				case "b": f.B = On(n); return true;
				case "i": f.I = On(n); return true;
				case "strike": f.Strike = On(n); return true;
				case "striked": f.DStrike = On(n); return true;
				case "caps": f.Caps = On(n); return true;
				case "scaps": f.SCaps = On(n); return true;
				case "v": f.Hidden = On(n); return true;
				case "outl": f.Outline = On(n); return true;
				case "shad": f.Shadow = On(n); return true;
				case "embo": f.Emboss = On(n); return true;
				case "impr": f.Imprint = On(n); return true;
				case "ul": f.Ul = On(n) ? "single" : ""; return true;
				case "ulnone": f.Ul = ""; return true;
				case "uldb": f.Ul = "double"; return true;
				case "uld": f.Ul = "dotted"; return true;
				case "uldash": f.Ul = "dash"; return true;
				case "uldashd": f.Ul = "dotDash"; return true;
				case "uldashdd": f.Ul = "dotDotDash"; return true;
				case "ulwave": f.Ul = "wave"; return true;
				case "ulth": f.Ul = "thick"; return true;
				case "ulw": f.Ul = "words"; return true;
				case "ulldash": f.Ul = "dashLong"; return true;
				case "uldb2": f.Ul = "double"; return true;
				case "cf": f.Color = n ?? 0; return true;
				case "highlight": f.Highlight = n ?? 0; return true;
				case "cb": case "chcbpat": f.Back = n is null or 0 ? -1 : n.Value; return true;
				case "super": f.Vert = 1; return true;
				case "sub": f.Vert = 2; return true;
				case "nosupersub": f.Vert = 0; return true;
				case "up": f.Position = n ?? 6; return true;
				case "dn": f.Position = -(n ?? 6); return true;
				case "expnd": f.Spacing = (n ?? 0) * 5; return true;
				case "expndtw": f.Spacing = n ?? 0; return true;
				case "charscalex": f.Scale = n ?? 100; return true;
				case "lang": case "langfe": case "langnp": case "langfenp": case "noproof": case "cs": case "kerning": case "loch": case "hich": case "dbch": case "af": case "afs": case "ab": case "ai": case "alang": case "insrsid": case "charrsid": case "rtlch": case "ltrch":
					return true;
			}
			return false;
		}

		bool ParaWord(ParaFmt f, string w, int? n) {
			switch (w) {
				case "s": f.Style = n ?? 0; return true;
				case "ql": f.Align = 'l'; return true;
				case "qc": f.Align = 'c'; return true;
				case "qr": f.Align = 'r'; return true;
				case "qj": f.Align = 'j'; return true;
				case "qd": f.Align = 'd'; return true;
				case "li": case "lin": f.Left = n ?? 0; return true;
				case "ri": case "rin": f.Right = n ?? 0; return true;
				case "fi": f.First = n ?? 0; return true;
				case "sb": f.Before = n ?? 0; return true;
				case "sa": f.After = n ?? 0; return true;
				case "sl": f.Sl = n ?? 0; if (f.Sl == 0) f.SlMult = false; return true;
				case "slmult": f.SlMult = On(n); return true;
				case "keep": f.Keep = On(n); return true;
				case "keepn": f.KeepNext = On(n); return true;
				case "pagebb": f.PageBefore = On(n); return true;
				case "widctlpar": f.Widow = true; return true;
				case "nowidctlpar": f.Widow = false; return true;
				case "contextualspace": f.Contextual = true; return true;
				case "intbl": f.InTable = true; return true;
				case "itap": f.Itap = n ?? 1; f.InTable = f.Itap > 0; return true;
				case "ls": f.Ls = n ?? 0; return true;
				case "ilvl": f.Ilvl = n ?? 0; return true;
				case "rtesuffix": f.Suffix = n ?? 0; return true;
				case "outlinelevel": f.Outline = n ?? -1; return true;
				case "tqr": f.TabKind = 'r'; return true;
				case "tqc": f.TabKind = 'c'; return true;
				case "tqdec": f.TabKind = 'd'; return true;
				case "tldot": f.TabLeader = 'd'; return true;
				case "tlhyph": f.TabLeader = 'h'; return true;
				case "tlul": f.TabLeader = 'u'; return true;
				case "tlth": f.TabLeader = 't'; return true;
				case "tleq": f.TabLeader = 'e'; return true;
				case "tlmdot": f.TabLeader = 'm'; return true;
				case "tx": case "tb":
					f.Tabs.RemoveAll(t => t.Pos == (n ?? 0));
					f.Tabs.Add(new TabStop { Pos = n ?? 0, Kind = w == "tb" ? 'b' : f.TabKind, Leader = f.TabLeader });
					f.TabKind = 'l'; f.TabLeader = 'n';
					return true;
				case "jclisttab": return true;
				case "cbpat": f.Shading = n ?? -1; return true;
				case "brdrt": border = f.Borders[0] = new Border(); return true;
				case "brdrl": border = f.Borders[1] = new Border(); return true;
				case "brdrb": border = f.Borders[2] = new Border(); return true;
				case "brdrr": border = f.Borders[3] = new Border(); return true;
				case "box": border = new Border(); for (int k = 0; k < 4; k++) f.Borders[k] = border; return true;
				case "nooverflow": case "nocwrap": case "aspalpha": case "aspnum": case "faauto": case "adjustright": case "rtlpar": case "ltrpar": case "sbauto": case "saauto": case "nosnaplinegrid": case "pararsid": case "yts": case "noline": case "hyphpar": case "wrapdefault":
					return true;
			}
			return false;
		}

		bool SectWord(string w, int? n) {
			switch (w) {
				case "sbknone": sect.Break = "continuous"; return true;
				case "sbkpage": sect.Break = "page"; return true;
				case "sbkcol": sect.Break = "column"; return true;
				case "sbkeven": sect.Break = "even"; return true;
				case "sbkodd": sect.Break = "odd"; return true;
				case "cols": sect.Cols = Math.Max(1, n ?? 1); return true;
				case "colsx": sect.ColSpace = n ?? 720; return true;
				case "linebetcol": sect.ColLine = true; return true;
				case "pgwsxn": sect.PgW = n ?? sect.PgW; return true;
				case "pghsxn": sect.PgH = n ?? sect.PgH; return true;
				case "marglsxn": sect.ML = n ?? sect.ML; return true;
				case "margrsxn": sect.MR = n ?? sect.MR; return true;
				case "margtsxn": sect.MT = n ?? sect.MT; return true;
				case "margbsxn": sect.MB = n ?? sect.MB; return true;
				case "guttersxn": sect.Gutter = n ?? 0; return true;
				case "headery": sect.HeaderY = n ?? 720; return true;
				case "footery": sect.FooterY = n ?? 720; return true;
				case "lndscpsxn": sect.Landscape = true; return true;
				case "titlepg": sect.TitlePg = true; return true;
				case "pgnstarts": sect.PgnStart = n ?? 1; return true;
				case "pgnrestart": sect.PgnRestart = true; return true;
				case "pgncont": sect.PgnRestart = false; return true;
				case "pgndec": sect.PgnFmt = null; return true;
				case "pgnucrm": sect.PgnFmt = "upperRoman"; return true;
				case "pgnlcrm": sect.PgnFmt = "lowerRoman"; return true;
				case "pgnucltr": sect.PgnFmt = "upperLetter"; return true;
				case "pgnlcltr": sect.PgnFmt = "lowerLetter"; return true;
				case "linemod": sect.LineMod = n ?? 0; return true;
				case "linex": sect.LineX = n ?? 0; return true;
				case "linestarts": sect.LineStart = n ?? 1; return true;
				case "linerestart": sect.LineRestart = "newSection"; return true;
				case "lineppage": sect.LineRestart = "newPage"; return true;
				case "linecont": sect.LineRestart = "continuous"; return true;
				case "vertalt": sect.VAlign = null; return true;
				case "vertalc": sect.VAlign = "center"; return true;
				case "vertalj": sect.VAlign = "both"; return true;
				case "vertalb": sect.VAlign = "bottom"; return true;
				case "endnhere": case "sectunlocked": case "sftnbj": case "srauth": case "srdate": case "sectrsid": case "sectlinegrid": case "pgnhn": case "pgnhnsc":
					return true;
			}
			return false;
		}

		bool TableWord(string w, int? n) {
			switch (w) {
				case "trgaph": row.Gap = n ?? 0; return true;
				case "trleft": row.Left = n ?? 0; return true;
				case "trql": row.Align = 'l'; return true;
				case "trqc": row.Align = 'c'; return true;
				case "trqr": row.Align = 'r'; return true;
				case "trrh": row.Height = n ?? 0; return true;
				case "trhdr": row.Header = true; return true;
				case "trkeep": row.Keep = true; return true;
				case "trpaddl": row.PadL = n; return true;
				case "trpaddr": row.PadR = n; return true;
				case "clbrdrt": border = cell.Borders[0] = new Border(); return true;
				case "clbrdrl": border = cell.Borders[1] = new Border(); return true;
				case "clbrdrb": border = cell.Borders[2] = new Border(); return true;
				case "clbrdrr": border = cell.Borders[3] = new Border(); return true;
				case "trbrdrt": case "trbrdrl": case "trbrdrb": case "trbrdrr": case "trbrdrh": case "trbrdrv": border = new Border(); return true;
				case "clcbpat": cell.CbPat = n ?? -1; return true;
				case "clcfpat": cell.CfPat = n ?? -1; return true;
				case "clshdng": cell.Shading = n ?? 0; return true;
				case "clvertalt": cell.VAlign = 't'; return true;
				case "clvertalc": cell.VAlign = 'c'; return true;
				case "clvertalb": cell.VAlign = 'b'; return true;
				case "clvmgf": cell.VMerge = "restart"; return true;
				case "clvmrg": cell.VMerge = "continue"; return true;
				case "clmgf": cell.HMergeFirst = true; return true;
				case "clmrg": cell.HMerge = true; return true;
				case "cellx": cell.Right = n ?? 0; row.Cells.Add(cell); cell = new CellDef(); return true;
				case "trpaddt": case "trpaddb": case "trpaddfl": case "trpaddfr": case "trpaddft": case "trpaddfb": case "trautofit": case "trftsWidth": case "trwWidth": case "trwWidthB": case "trwWidthA": case "trftsWidthB": case "trftsWidthA":
				case "tblind": case "tblindtype": case "clwWidth": case "clftsWidth": case "clpadl": case "clpadr": case "clpadt": case "clpadb": case "clpadfl": case "clpadfr": case "clpadft": case "clpadfb": case "irowband": case "irow": case "lastrow": case "trspdl": case "trspdr": case "trspdfl": case "trspdfr":
					return true;
			}
			return false;
		}

		bool BorderWord(string w, int? n) {
			if (border == null) return w.StartsWith("brdr", StringComparison.Ordinal);
			switch (w) {
				case "brdrs": border.Style = "single"; return true;
				case "brdrdb": border.Style = "double"; return true;
				case "brdrth": border.Style = "thick"; return true;
				case "brdrdot": border.Style = "dotted"; return true;
				case "brdrdash": border.Style = "dashed"; return true;
				case "brdrdashsm": border.Style = "dashSmallGap"; return true;
				case "brdrdashd": border.Style = "dotDash"; return true;
				case "brdrdashdd": border.Style = "dotDotDash"; return true;
				case "brdrtriple": border.Style = "triple"; return true;
				case "brdrwavy": border.Style = "wave"; return true;
				case "brdrinset": border.Style = "inset"; return true;
				case "brdroutset": border.Style = "outset"; return true;
				case "brdrnone": case "brdrnil": case "brdrtbl": border.Style = "none"; return true;
				case "brdrw": border.Width = n ?? 10; return true;
				case "brdrcf": border.Color = n ?? -1; return true;
				case "brsp": border.Space = n ?? 0; return true;
			}
			return w.StartsWith("brdr", StringComparison.Ordinal);
		}

		// ---- the header tables

		void FontTable(Grp g) {
			foreach (var entry in g.Kids.OfType<Grp>()) {
				var info = new FontInfo();
				int number = -1;
				var name = new StringBuilder();
				foreach (var item in entry.Kids) {
					switch (item) {
						case Cw { Word: "f" } f: number = f.N ?? -1; break;
						case Cw { Word: "fcharset" } cs: info.Charset = cs.N ?? 0; break;
						case Cw { Word: "cpg" } cpg: info.CodePage = cpg.N ?? 1252; break;
						case Txt t: name.Append(t.S); break;
						case Hex h: name.Append((char)h.B); break;
						case Uni u: name.Append((char)u.Code); break;
					}
				}
				if (number < 0) continue;
				info.Name = name.ToString().Split(';')[0].Trim();
				if (!entry.Kids.OfType<Cw>().Any(k => k.Word == "cpg")) info.CodePage = CodePageOf(info.Charset);
				while (fonts.Count <= number) fonts.Add(null);
				fonts[number] = info;
			}
		}

		int CodePageOf(int charset) => charset switch {
			0 => 1252, 1 => ansiCodePage, 2 => 1252, 77 => 10000, 128 => 932, 129 => 949, 130 => 1361, 134 => 936, 136 => 950, 161 => 1253, 162 => 1254, 163 => 1258,
			177 => 1255, 178 => 1256, 186 => 1257, 204 => 1251, 222 => 874, 238 => 1250, 255 => 437, _ => ansiCodePage };

		void ColorTable(Grp g) {
			int r = -1, gr = -1, b = -1;
			foreach (var item in g.Kids.Skip(1)) {
				switch (item) {
					case Cw { Word: "red" } x: r = x.N ?? 0; break;
					case Cw { Word: "green" } x: gr = x.N ?? 0; break;
					case Cw { Word: "blue" } x: b = x.N ?? 0; break;
					case Txt t:
						foreach (char ch in t.S) {
							if (ch != ';') continue;
							colors.Add(r < 0 && gr < 0 && b < 0 ? null : $"{Math.Max(0, r):X2}{Math.Max(0, gr):X2}{Math.Max(0, b):X2}");
							r = gr = b = -1;
						}
						break;
				}
			}
		}

		string? ColorHex(int index) => index > 0 && index < colors.Count ? colors[index] : null;

		void StyleSheet(Grp g) {
			foreach (var entry in g.Kids.OfType<Grp>()) {
				if (Starred(entry) || entry.Kids.OfType<Cw>().Any(k => k.Word is "ts" or "ds")) continue;    // (character, table and section styles: not written)
				var def = new StyleDef();
				var savedBorder = border;
				var name = new StringBuilder();
				foreach (var item in entry.Kids) {
					switch (item) {
						case Cw { Word: "s" } s: def.Number = s.N ?? 0; break;
						case Cw { Word: "sbasedon" } x: def.BasedOn = x.N ?? -1; break;
						case Cw { Word: "snext" } x: def.Next = x.N ?? -1; break;
						case Cw x: if (!CharWord(def.C, x.Word, x.N)) ParaWord(def.P, x.Word, x.N); break;
						case Txt t: name.Append(t.S); break;
						case Hex h: name.Append((char)h.B); break;
					}
				}
				border = savedBorder;
				def.P.Style = def.Number;
				def.Name = name.ToString().Split(';')[0].Trim();
				if (def.Name == "") def.Name = "Style " + def.Number;
				styles[def.Number] = def;
			}
			var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
			foreach (var def in styles.Values.OrderBy(d => d.Number)) {
				string id = def.Number == 0 ? "Normal" : StyleIdOf(def.Name);
				string unique = id;
				for (int k = 2; !used.Add(unique); k++) unique = id + k;
				def.Id = unique;
			}
		}

		static string StyleIdOf(string name) {
			var m = Regex.Match(name, @"^toc (\d)$", RegexOptions.IgnoreCase);
			if (m.Success) return "TOC" + m.Groups[1].Value;
			var sb = new StringBuilder();
			foreach (var word in Regex.Split(name, @"[^A-Za-z0-9]+").Where(x => x.Length > 0))
				sb.Append(char.ToUpperInvariant(word[0])).Append(word[1..]);
			return sb.Length > 0 ? sb.ToString() : "Style";
		}

		StyleDef? StyleOf(int number) => styles.TryGetValue(number, out var s) ? s : null;

		// ---- lists

		void ListTable(Grp g) {
			foreach (var list in g.Kids.OfType<Grp>().Where(x => DestOf(x) == "list")) {
				int listId = list.Kids.OfType<Cw>().FirstOrDefault(k => k.Word == "listid")?.N ?? 0;
				bool hybrid = list.Kids.OfType<Cw>().Any(k => k.Word == "listhybrid");
				bool simple = list.Kids.OfType<Cw>().Any(k => k.Word == "listsimple" && On(k.N));
				var levels = list.Kids.OfType<Grp>().Where(x => DestOf(x) == "listlevel").ToList();
				int id = abstracts.Count;
				var abs = new W.AbstractNum { AbstractNumberId = id };
				abs.Append(new W.Nsid { Val = ((uint)listId).ToString("X8") });    // (the RTF's list id, as Word keeps it: Word tells lists apart by it)
				abs.Append(new W.MultiLevelType { Val = hybrid ? W.MultiLevelValues.HybridMultilevel : simple || levels.Count == 1 ? W.MultiLevelValues.SingleLevel : W.MultiLevelValues.Multilevel });
				abs.Append(new W.TemplateCode { Val = "FFFFFFFF" });
				for (int k = 0; k < levels.Count && k < 9; k++) abs.Append(Level(levels[k], k, id));
				// (all nine levels, as Word writes them: a list the RTF gives fewer levels came back from Word differently)
				for (int k = levels.Count; k < 9; k++) abs.Append(new W.Level(new W.NumberingFormat { Val = W.NumberFormatValues.Decimal }, new W.LevelText { Val = "" }, new W.LevelJustification { Val = W.LevelJustificationValues.Left }) { LevelIndex = k, TemplateCode = "FFFFFFFF" });
				abstracts.Add(abs);
				abstractOfListId[listId] = id;
			}
		}

		W.Level Level(Grp g, int index, int abstractId) {
			var cf = new CharFmt { Font = int.MinValue };
			var pf = new ParaFmt();
			int nfc = 0, jc = 0, follow = 0, start = 1;
			bool legal = false, noRestart = false;
			var units = new List<(bool IsByte, int Value)>();
			var savedBorder = border;
			foreach (var item in g.Kids) {
				switch (item) {
					case Cw { Word: "levelnfc" or "levelnfcn" } x: nfc = x.N ?? 0; break;
					case Cw { Word: "leveljc" or "leveljcn" } x: jc = x.N ?? 0; break;
					case Cw { Word: "levelfollow" } x: follow = x.N ?? 0; break;
					case Cw { Word: "levelstartat" } x: start = x.N ?? 1; break;
					case Cw { Word: "levellegal" } x: legal = On(x.N); break;
					case Cw { Word: "levelnorestart" } x: noRestart = On(x.N); break;
					case Cw x: if (!CharWord(cf, x.Word, x.N)) ParaWord(pf, x.Word, x.N); break;
					case Grp sub when DestOf(sub) == "leveltext":
						foreach (var u in sub.Kids) {
							if (u is Hex h) units.Add((true, h.B));
							else if (u is Txt t) foreach (char ch in t.S) units.Add((false, ch));
							else if (u is Uni un) units.Add((false, un.Code));
						}
						break;
				}
			}
			border = savedBorder;
			int font = cf.Font == int.MinValue ? -1 : cf.Font;
			var text = new StringBuilder();
			if (units.Count > 0) {
				int length = units[0].Value;
				var bytes = new List<byte>();
				// (a symbol font's bullet stays its byte, U+00B7 in Symbol, as Word keeps it: U+F0B7 came back from Word's RTF as a \u
				// escape, which RichEdit couldn't read in a list level, and the whole list then lost its numbering)
				void FlushBytes() { if (bytes.Count > 0) { text.Append(Decode(bytes.ToArray(), font)); bytes.Clear(); } }
				for (int k = 1; k <= length && k < units.Count; k++) {
					var (isByte, value) = units[k];
					if (isByte && value < 9) { FlushBytes(); text.Append('%').Append(value + 1); }
					else if (isByte) bytes.Add((byte)value);
					else { FlushBytes(); text.Append((char)value); }
				}
				FlushBytes();
			}
			var level = new W.Level { LevelIndex = index, TemplateCode = "FFFFFFFF" };    // (as Word writes a list it read from RTF)
			level.Append(new W.StartNumberingValue { Val = start });
			level.Append(new W.NumberingFormat { Val = NumberFormat(nfc) });
			if (noRestart) level.Append(new W.LevelRestart { Val = 0 });
			if (legal) level.Append(new W.IsLegalNumberingStyle());
			if (follow == 1) level.Append(new W.LevelSuffix { Val = W.LevelSuffixValues.Space });
			else if (follow == 2) level.Append(new W.LevelSuffix { Val = W.LevelSuffixValues.Nothing });
			level.Append(new W.LevelText { Val = text.ToString() });
			level.Append(new W.LevelJustification { Val = jc == 1 ? W.LevelJustificationValues.Center : jc == 2 ? W.LevelJustificationValues.Right : W.LevelJustificationValues.Left });
			var ppr = new W.PreviousParagraphProperties();
			if (pf.Tabs.Count > 0) ppr.Append(new W.Tabs(pf.Tabs.Select(t => new W.TabStop { Val = W.TabStopValues.Number, Position = t.Pos })));
			levelTabs[(abstractId, index)] = pf.Tabs.Select(t => t.Pos).ToList();
			if (pf.Left != 0 || pf.First != 0) ppr.Append(Indent(pf.Left, pf.Right, pf.First));
			if (ppr.HasChildren) level.Append(ppr);
			if (font >= 0) {
				string name = FontName(font);
				level.Append(new W.NumberingSymbolRunProperties(new W.RunFonts { Ascii = name, HighAnsi = name, ComplexScript = name, Hint = W.FontTypeHintValues.Default }));
			}
			return level;
		}

		static W.NumberFormatValues NumberFormat(int nfc) => nfc switch {
			1 => W.NumberFormatValues.UpperRoman, 2 => W.NumberFormatValues.LowerRoman, 3 => W.NumberFormatValues.UpperLetter, 4 => W.NumberFormatValues.LowerLetter,
			5 => W.NumberFormatValues.Ordinal, 6 => W.NumberFormatValues.CardinalText, 7 => W.NumberFormatValues.OrdinalText, 22 => W.NumberFormatValues.DecimalZero,
			23 => W.NumberFormatValues.Bullet, 255 => W.NumberFormatValues.None, _ => W.NumberFormatValues.Decimal };

		void ListOverrides(Grp g) {
			foreach (var o in g.Kids.OfType<Grp>().Where(x => DestOf(x) == "listoverride")) {
				int listId = o.Kids.OfType<Cw>().FirstOrDefault(k => k.Word == "listid")?.N ?? 0;
				int ls = o.Kids.OfType<Cw>().FirstOrDefault(k => k.Word == "ls")?.N ?? 0;
				if (ls <= 0 || !abstractOfListId.TryGetValue(listId, out int abs)) continue;
				var num = new W.NumberingInstance(new W.AbstractNumId { Val = abs }) { NumberID = ls };
				abstractOfNum[ls] = abs;
				int level = 0;
				foreach (var lfo in o.Kids.OfType<Grp>().Where(x => DestOf(x) == "lfolevel")) {
					var words = lfo.Kids.OfType<Cw>().ToList();
					if (words.Any(k => k.Word == "listoverridestartat") && words.FirstOrDefault(k => k.Word == "levelstartat") is { } at)
						num.Append(new W.LevelOverride(new W.StartOverrideNumberingValue { Val = at.N ?? 1 }) { LevelIndex = level });
					level++;
				}
				nums.Add(num);
				numIds.Add(ls);
			}
		}

		void PnGroup(Grp g) {
			var pn = new PnDef();
			foreach (var item in g.Kids) {
				switch (item) {
					case Cw { Word: "pnlvlbody" }: pn.Kind = "body"; break;
					case Cw { Word: "pnlvlblt" }: pn.Kind = "blt"; pn.Fmt = "blt"; break;
					case Cw { Word: "pnlvlcont" }: pn.Kind = "cont"; break;
					case Cw { Word: "pnlvl" } x: pn.Kind = "lvl" + (x.N ?? 1); break;
					case Cw { Word: "pndec" }: pn.Fmt = "dec"; break;
					case Cw { Word: "pnucrm" }: pn.Fmt = "ucrm"; break;
					case Cw { Word: "pnlcrm" }: pn.Fmt = "lcrm"; break;
					case Cw { Word: "pnucltr" }: pn.Fmt = "ucltr"; break;
					case Cw { Word: "pnlcltr" }: pn.Fmt = "lcltr"; break;
					case Cw { Word: "pnord" }: pn.Fmt = "ord"; break;
					case Cw { Word: "pncard" }: pn.Fmt = "card"; break;
					case Cw { Word: "pnordt" }: pn.Fmt = "ordt"; break;
					case Cw { Word: "pnstart" } x: pn.Start = x.N ?? 1; break;
					case Cw { Word: "pnf" } x: pn.Font = x.N ?? -1; break;
					case Grp sub when DestOf(sub) == "pntxtb": pn.Before = PnText(sub, pn.Font); break;
					case Grp sub when DestOf(sub) == "pntxta": pn.After = PnText(sub, pn.Font); break;
				}
			}
			p.Pn = pn;
		}

		// (iteration 87) A list level's suffix (1 space, 2 nothing), where it has none: what follows its numbers.
		void SuffixInto(int? abstractId, int level, int suffix) {
			if (abstractId is not int id || abstracts.FirstOrDefault(a => a.AbstractNumberId?.Value == id) is not { } abs) return;
			if (abs.Elements<W.Level>().FirstOrDefault(l => l.LevelIndex?.Value == level) is not { } lvl || lvl.GetFirstChild<W.LevelSuffix>() != null) return;
			var suff = new W.LevelSuffix { Val = suffix == 1 ? W.LevelSuffixValues.Space : W.LevelSuffixValues.Nothing };
			if (lvl.GetFirstChild<W.LevelText>() is { } text) text.InsertBeforeSelf(suff); else lvl.Append(suff);
		}

		string PnText(Grp g, int font) {
			var sb = new StringBuilder();
			var bytes = new List<byte>();
			foreach (var item in g.Kids.Skip(1)) {
				if (item is Hex h) { bytes.Add(h.B); continue; }
				if (bytes.Count > 0) { sb.Append(Decode(bytes.ToArray(), font)); bytes.Clear(); }
				if (item is Txt t) sb.Append(t.S);
				else if (item is Uni u) sb.Append((char)u.Code);
			}
			if (bytes.Count > 0) sb.Append(Decode(bytes.ToArray(), font));
			return sb.ToString();    // (a symbol font's byte as it is, as in the list levels)
		}

		// The numbering of a paragraph: (numId, level), or null.
		(int Id, int Level)? Numbering(ParaFmt f, string? shown) {
			if (f.Ls > 0 && numIds.Contains(f.Ls)) {
				if (f.Suffix > 0) SuffixInto(nums.FirstOrDefault(n => n.NumberID?.Value == f.Ls)?.AbstractNumId?.Val?.Value, Math.Clamp(f.Ilvl, 0, 8), f.Suffix);
				return (f.Ls, Math.Clamp(f.Ilvl, 0, 8));
			}
			if (f.Pn is not { } pn || pn.Kind == "cont") return null;
			string key = pn.Key(FontName) + (f.Suffix > 0 ? "|suffix" + f.Suffix : "");
			if (!pnAbstracts.TryGetValue(key, out int abs)) {
				abs = abstracts.Count;
				var a = new W.AbstractNum { AbstractNumberId = abs };
				a.Append(new W.Nsid { Val = (0x2A000000 + abs).ToString("X8") });
				a.Append(new W.MultiLevelType { Val = W.MultiLevelValues.SingleLevel });
				var level = new W.Level { LevelIndex = 0 };
				level.Append(new W.StartNumberingValue { Val = 1 });
				level.Append(new W.NumberingFormat { Val = pn.Fmt switch {
					"blt" => W.NumberFormatValues.Bullet, "ucrm" => W.NumberFormatValues.UpperRoman, "lcrm" => W.NumberFormatValues.LowerRoman,
					"ucltr" => W.NumberFormatValues.UpperLetter, "lcltr" => W.NumberFormatValues.LowerLetter, "ord" => W.NumberFormatValues.Ordinal,
					"card" => W.NumberFormatValues.CardinalText, "ordt" => W.NumberFormatValues.OrdinalText, _ => W.NumberFormatValues.Decimal } });
				if (f.Suffix > 0) level.Append(new W.LevelSuffix { Val = f.Suffix == 1 ? W.LevelSuffixValues.Space : W.LevelSuffixValues.Nothing });
				level.Append(new W.LevelText { Val = pn.Fmt == "blt" ? (pn.Before.Length > 0 ? pn.Before : "\u2022") : pn.Before + "%1" + pn.After });
				level.Append(new W.LevelJustification { Val = W.LevelJustificationValues.Left });
				if (pn.Font >= 0) {
					string name = FontName(pn.Font);
					level.Append(new W.NumberingSymbolRunProperties(new W.RunFonts { Ascii = name, HighAnsi = name, ComplexScript = name, Hint = W.FontTypeHintValues.Default }));
				}
				a.Append(level);
				abstracts.Add(a);
				pnAbstracts[key] = abs;
			}
			int number = pn.Fmt == "blt" ? 1 : ShownNumber(shown, pn) ?? pn.Start;
			if (lastPn is { } last && last.Key == key && (pn.Fmt == "blt" || number == last.Next)) {
				lastPn = (key, last.NumId, number + 1);
				return (last.NumId, 0);
			}
			int id = nextPnNum++;
			var num = new W.NumberingInstance(new W.AbstractNumId { Val = abs }) { NumberID = id };
			// (a start of 1 said too where an earlier list had this definition: Word numbers such lists on, one after the other)
			if (number != 1 || nums.Any(n => n.AbstractNumId?.Val?.Value == abs)) num.Append(new W.LevelOverride(new W.StartOverrideNumberingValue { Val = number }) { LevelIndex = 0 });
			nums.Add(num);
			numIds.Add(id);
			lastPn = (key, id, number + 1);
			return (id, 0);
		}

		static int? ShownNumber(string? shown, PnDef pn) {
			if (string.IsNullOrEmpty(shown)) return null;
			string s = shown.Replace("\t", "").Trim();
			if (pn.Before.Length > 0 && s.StartsWith(pn.Before, StringComparison.Ordinal)) s = s[pn.Before.Length..];
			if (pn.After.Length > 0 && s.EndsWith(pn.After, StringComparison.Ordinal)) s = s[..^pn.After.Length];
			s = s.Trim();
			if (s.Length == 0) return null;
			switch (pn.Fmt) {
				case "dec": case "ord": return int.TryParse(Regex.Match(s, @"\d+").Value, out int d) ? d : null;
				case "ucltr": case "lcltr":
					if (!Regex.IsMatch(s, "^[A-Za-z]+$")) return null;
					return (char.ToLowerInvariant(s[0]) - 'a' + 1) + 26 * (s.Length - 1);
				case "ucrm": case "lcrm": {
					int total = 0, prev = 0;
					foreach (char ch in s.ToUpperInvariant().Reverse()) {
						int v = ch switch { 'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50, 'C' => 100, 'D' => 500, 'M' => 1000, _ => -1 };
						if (v < 0) return null;
						total += v < prev ? -v : v;
						prev = Math.Max(prev, v);
					}
					return total;
				}
			}
			return null;
		}

		// ---- fields, bookmarks, pictures, notes, headers

		void Field(Grp g) {
			var inst = g.Kids.OfType<Grp>().FirstOrDefault(x => DestOf(x) == "fldinst");
			var result = g.Kids.OfType<Grp>().FirstOrDefault(x => DestOf(x) == "fldrslt");
			bool locked = g.Kids.OfType<Cw>().Any(k => k.Word == "fldlock");
			string instr = inst == null ? "" : PlainText(inst).Trim();
			var link = Regex.Match(instr, @"^HYPERLINK\s+(?<l>\\l\s+)?(?:""(?<t>[^""]*)""|(?<t>\S+))(?<rest>.*)$", RegexOptions.Singleline);
			if (link.Success && result != null && !Breaks(result)) {
				string target = link.Groups["t"].Value;
				var sub = Regex.Match(link.Groups["rest"].Value, @"\\l\s+""([^""]*)""");
				string? anchor = link.Groups["l"].Success ? target : target.StartsWith('#') ? target[1..] : sub.Success ? sub.Groups[1].Value : null;
				string? url = link.Groups["l"].Success || target.StartsWith('#') ? null : target;
				Mark(K.HlStart, i => { i.Url = url; i.Anchor = url == null ? anchor : sub.Success ? sub.Groups[1].Value : null; });
				Walk(result);
				Mark(K.HlEnd);
				return;
			}
			Mark(K.FldBegin, i => { i.Text = " " + instr + " "; i.Lock = locked; });
			Mark(K.FldSep);
			if (result != null) Walk(result);
			Mark(K.FldEnd);
		}

		static bool Breaks(Grp g) => g.Kids.Any(k => k is Cw { Word: "par" or "sect" or "cell" or "row" or "nestcell" or "nestrow" } || k is Grp sub && DestOf(sub) is not ("footnote" or "pntext" or "listtext") && Breaks(sub));

		void BookmarkStart(string name) {
			if (name.Length == 0 || openBookmarks.ContainsKey(name)) return;
			int id = nextBookmark++;
			openBookmarks[name] = id;
			Mark(K.BmStart, i => { i.Id = id; i.Text = name; });
		}

		void BookmarkEnd(string name) {
			if (!openBookmarks.Remove(name, out int id)) return;
			Mark(K.BmEnd, i => i.Id = id);
		}

		void Picture(Grp g) {
			var pic = new PicData();
			int picw = 0, pich = 0, goalW = 0, goalH = 0, scaleX = 100, scaleY = 100;
			var hex = new StringBuilder();
			byte[]? binary = null;
			foreach (var item in g.Kids) {
				switch (item) {
					case Cw { Word: "pngblip" }: pic.Type = "png"; break;
					case Cw { Word: "jpegblip" }: pic.Type = "jpeg"; break;
					case Cw { Word: "emfblip" }: pic.Type = "emf"; break;
					case Cw { Word: "wmetafile" }: pic.Type = "wmf"; break;
					case Cw { Word: "dibitmap" or "wbitmap" }: pic.Type = "dib"; break;
					case Cw { Word: "picw" } x: picw = x.N ?? 0; break;
					case Cw { Word: "pich" } x: pich = x.N ?? 0; break;
					case Cw { Word: "picwgoal" } x: goalW = x.N ?? 0; break;
					case Cw { Word: "pichgoal" } x: goalH = x.N ?? 0; break;
					case Cw { Word: "picscalex" } x: scaleX = x.N ?? 100; break;
					case Cw { Word: "picscaley" } x: scaleY = x.N ?? 100; break;
					case Txt t: hex.Append(t.S); break;
					case Bin b: binary = b.Data; break;
					case Grp sub when DestOf(sub) == "picprop":
						foreach (var sp in sub.Kids.OfType<Grp>().Where(x => DestOf(x) == "sp")) {
							var sn = sp.Kids.OfType<Grp>().FirstOrDefault(x => DestOf(x) == "sn");
							var sv = sp.Kids.OfType<Grp>().FirstOrDefault(x => DestOf(x) == "sv");
							if (sn != null && sv != null && PlainText(sn).Trim() == "wzDescription") pic.Alt = PlainText(sv);
						}
						break;
				}
			}
			if (pic.Type == "dib") return;
			string digits = new string(hex.ToString().Where(Uri.IsHexDigit).ToArray());
			if (digits.Length % 2 == 1) digits = digits[..^1];
			try { pic.Data = binary ?? Convert.FromHexString(digits); }
			catch { return; }
			if (pic.Data.Length == 0) return;
			long twW = goalW > 0 ? goalW : picw * 1440L / 2540, twH = goalH > 0 ? goalH : pich * 1440L / 2540;
			twW = twW * scaleX / 100; twH = twH * scaleY / 100;
			if (twW <= 0 || twH <= 0) { twW = 1440; twH = 1440; }
			pic.Cx = twW * 635; pic.Cy = twH * 635;
			if (pic.Type == "wmf") pic.Data = PlaceableWmf(pic.Data, (int)twW, (int)twH);
			Mark(K.Pic, i => i.Pic = pic);
		}

		// RTF keeps a metafile without the "placeable" header a .wmf file has
		static byte[] PlaceableWmf(byte[] data, int width, int height) {
			var header = new byte[22];
			BitConverter.GetBytes(0x9AC6CDD7u).CopyTo(header, 0);
			BitConverter.GetBytes((short)Math.Min(width, short.MaxValue)).CopyTo(header, 10);
			BitConverter.GetBytes((short)Math.Min(height, short.MaxValue)).CopyTo(header, 12);
			BitConverter.GetBytes((short)1440).CopyTo(header, 14);
			ushort sum = 0;
			for (int k = 0; k < 20; k += 2) sum ^= BitConverter.ToUInt16(header, k);
			BitConverter.GetBytes(sum).CopyTo(header, 20);
			return header.Concat(data).ToArray();
		}

		void Footnote(Grp g) {
			bool end = g.Kids.OfType<Cw>().Any(k => k.Word == "ftnalt");
			var note = pendingNote;
			pendingNote = null;
			if (note == null) { note = new Note(); Mark(K.NoteRef, i => i.Note = note); }
			note.End = end;
			OpenXmlPart part;
			if (end) {
				if (endnotes == null) { endnotes = main.AddNewPart<EndnotesPart>(); endnotes.Endnotes = new W.Endnotes(Separator<W.Endnote>(-1, true), Separator<W.Endnote>(0, false)); }
				note.Id = nextEndnote++;
				part = endnotes;
			} else {
				if (footnotes == null) { footnotes = main.AddNewPart<FootnotesPart>(); footnotes.Footnotes = new W.Footnotes(Separator<W.Footnote>(-1, true), Separator<W.Footnote>(0, false)); }
				note.Id = nextFootnote++;
				part = footnotes;
			}
			var saved = (story, lastPn);
			story = new Story { Part = part };
			WalkInside(g);
			FinishStory(story);
			var blocks = story.Blocks.Count > 0 ? story.Blocks : new List<OpenXmlElement> { new W.Paragraph() };
			if (end) endnotes!.Endnotes.Append(new W.Endnote(blocks) { Id = note.Id });
			else footnotes!.Footnotes.Append(new W.Footnote(blocks) { Id = note.Id });
			(story, lastPn) = saved;
		}

		static T Separator<T>(long id, bool first) where T : OpenXmlCompositeElement, new() {
			var note = new T();
			note.Append(new W.Paragraph(new W.ParagraphProperties(new W.SpacingBetweenLines { After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }),
				new W.Run(first ? new W.SeparatorMark() : new W.ContinuationSeparatorMark())));
			if (note is W.Footnote f) { f.Id = id; f.Type = first ? W.FootnoteEndnoteValues.Separator : W.FootnoteEndnoteValues.ContinuationSeparator; }
			if (note is W.Endnote e) { e.Id = id; e.Type = first ? W.FootnoteEndnoteValues.Separator : W.FootnoteEndnoteValues.ContinuationSeparator; }
			return note;
		}

		void HeaderFooter(Grp g, string word) {
			bool header = word.StartsWith("header", StringComparison.Ordinal);
			OpenXmlPart part = header ? main.AddNewPart<HeaderPart>() : main.AddNewPart<FooterPart>();
			var saved = (story, lastPn);
			story = new Story { Part = part };
			WalkInside(g);
			FinishStory(story);
			if (story.Blocks.Count == 0 || story.Blocks[^1] is W.Table) story.Blocks.Add(new W.Paragraph());
			if (part is HeaderPart hp) { hp.Header = new W.Header(story.Blocks); hp.Header.Save(); }
			else if (part is FooterPart fp) { fp.Footer = new W.Footer(story.Blocks); fp.Footer.Save(); }
			string type = word[^1] switch { 'f' => "first", 'l' => "even", _ => "default" };
			sect.Refs.RemoveAll(r => r.Header == header && r.Type == type);
			sect.Refs.Add((header, type, main.GetIdOfPart(part)));
			(story, lastPn) = saved;
		}

		// ---- paragraphs, rows, tables

		// level: the table level the paragraph is in (0 the text, 1 a table, 2 a table nested in a cell, ...)
		void EndParagraph(bool section = false, bool cellEnd = false, int level = -1) {
			var st = story;
			if (level < 0) level = section ? 0 : p.Itap > 0 ? p.Itap : p.InTable ? 1 : 0;
			var para = BuildParagraph(st, p, c, section ? SectPr(sect) : null);
			st.Line.Clear();
			st.PnText = null;
			ToLevel(st, level);
			if (level > 0) {
				var lv = st.Levels[level - 1];
				lv.Cell.Add(para);
				if (cellEnd) { lv.Cells.Add(lv.Cell); lv.Cell = new(); }
			} else {
				st.Blocks.Add(para);
			}
			if (section) sect = sect.Clone();
		}

		// the tables deeper than the level end (each one goes into the cell it is nested in), and levels up to it open
		void ToLevel(Story st, int level) {
			while (st.Levels.Count > Math.Max(1, level)) {
				var inner = st.Levels[^1];
				st.Levels.RemoveAt(st.Levels.Count - 1);
				if (inner.Cell.Count > 0 || inner.Cells.Count > 0) EndRowOf(inner, new RowDef());
				if (BuildTable(inner) is { } nested) st.Levels[^1].Cell.Add(nested);
			}
			if (level == 0) { if (BuildTable(st.Levels[0]) is { } table) st.Blocks.Add(table); }
			while (st.Levels.Count < level) st.Levels.Add(new TableLevel());
		}

		void EndRow(int level) {
			var st = story;
			ToLevel(st, level);
			EndRowOf(st.Levels[level - 1], row.Snapshot());
		}

		static void EndRowOf(TableLevel lv, RowDef def) {
			if (lv.Cell.Count > 0) { lv.Cells.Add(lv.Cell); lv.Cell = new(); }
			lv.Rows.Add((def, lv.Cells));
			lv.Cells = new();
		}

		void FinishStory(Story st) {
			if (st.Line.Count > 0) EndParagraph();
			ToLevel(st, 1);
			if (st.Levels[0].Cell.Count > 0 || st.Levels[0].Cells.Count > 0) EndRowOf(st.Levels[0], row.Snapshot());
			ToLevel(st, 0);
		}

		W.Table? BuildTable(TableLevel lv) {
			if (lv.Rows.Count == 0) return null;
			var rows = lv.Rows.ToList();
			lv.Rows.Clear();
			// every row's cells with their left and right edges; a row with more cells than definitions gets more
			var laid = new List<(RowDef Def, List<(int Left, int Right, CellDef Def, List<OpenXmlElement> Blocks)> Cells)>();
			foreach (var (def, cells) in rows) {
				var defs = def.Cells.ToList();
				while (defs.Count < cells.Count) defs.Add(new CellDef { Right = (defs.Count > 0 ? defs[^1].Right : def.Left) + 1440 });
				var list = new List<(int, int, CellDef, List<OpenXmlElement>)>();
				int left = def.Left;
				for (int k = 0; k < defs.Count; k++) {
					int right = Math.Max(defs[k].Right, left + 15);
					list.Add((left, right, defs[k], k < cells.Count ? cells[k] : new List<OpenXmlElement>()));
					left = right;
				}
				if (list.Count > 0) laid.Add((def, list));
			}
			var edges = laid.SelectMany(r => r.Cells.SelectMany(x => new[] { x.Left, x.Right })).Distinct().OrderBy(x => x).ToList();
			if (edges.Count < 2) return null;
			int tableLeft = edges[0], tableWidth = edges[^1] - edges[0];
			var first = laid[0].Def;
			var table = new W.Table();
			var tblPr = new W.TableProperties();
			tblPr.Append(new W.TableWidth { Width = tableWidth.ToString(CultureInfo.InvariantCulture), Type = W.TableWidthUnitValues.Dxa });
			if (first.Align != 'l') tblPr.Append(new W.TableJustification { Val = first.Align == 'c' ? W.TableRowAlignmentValues.Center : W.TableRowAlignmentValues.Right });
			if (tableLeft != 0) tblPr.Append(new W.TableIndentation { Width = tableLeft, Type = W.TableWidthUnitValues.Dxa });
			tblPr.Append(new W.TableLayout { Type = W.TableLayoutValues.Fixed });
			short padL = (short)Math.Clamp(first.PadL ?? first.Gap, 0, 2000), padR = (short)Math.Clamp(first.PadR ?? first.Gap, 0, 2000);
			tblPr.Append(new W.TableCellMarginDefault(new W.TableCellLeftMargin { Width = padL, Type = W.TableWidthValues.Dxa }, new W.TableCellRightMargin { Width = padR, Type = W.TableWidthValues.Dxa }));
			tblPr.Append(new W.TableLook { Val = "0000", FirstRow = false, LastRow = false, FirstColumn = false, LastColumn = false, NoHorizontalBand = true, NoVerticalBand = true });
			table.Append(tblPr);
			var grid = new W.TableGrid();
			for (int k = 1; k < edges.Count; k++) grid.Append(new W.GridColumn { Width = (edges[k] - edges[k - 1]).ToString(CultureInfo.InvariantCulture) });
			table.Append(grid);
			foreach (var (def, cells) in laid) {
				var tr = new W.TableRow();
				var trPr = new W.TableRowProperties();
				int before = edges.IndexOf(cells[0].Left);
				if (before > 0) trPr.Append(new W.GridBefore { Val = before });
				if (def.Keep) trPr.Append(new W.CantSplit());
				if (def.Height != 0) trPr.Append(new W.TableRowHeight { Val = (uint)Math.Abs(def.Height), HeightType = def.Height > 0 ? W.HeightRuleValues.AtLeast : W.HeightRuleValues.Exact });
				if (def.Header) trPr.Append(new W.TableHeader());
				if (trPr.HasChildren) tr.Append(trPr);
				W.TableCell? previous = null;
				foreach (var (left, right, cd, blocks) in cells) {
					if (cd.HMerge && previous != null) {
						// (the old horizontal merge: this cell's columns join the cell before)
						var span = previous.TableCellProperties!.GetFirstChild<W.GridSpan>();
						int had = span?.Val?.Value ?? 1;
						int more = Math.Max(1, edges.IndexOf(right) - edges.IndexOf(left));
						if (span == null) previous.TableCellProperties.InsertAfter(new W.GridSpan { Val = had + more }, previous.TableCellProperties.TableCellWidth);
						else span.Val = had + more;
						var width = previous.TableCellProperties.TableCellWidth!;
						width.Width = (int.Parse(width.Width!.Value!, CultureInfo.InvariantCulture) + right - left).ToString(CultureInfo.InvariantCulture);
						continue;
					}
					var tc = new W.TableCell();
					var tcPr = new W.TableCellProperties();
					tcPr.Append(new W.TableCellWidth { Width = (right - left).ToString(CultureInfo.InvariantCulture), Type = W.TableWidthUnitValues.Dxa });
					int spanCount = edges.IndexOf(right) - edges.IndexOf(left);
					if (spanCount > 1) tcPr.Append(new W.GridSpan { Val = spanCount });
					if (cd.VMerge != null) tcPr.Append(new W.VerticalMerge { Val = cd.VMerge == "restart" ? W.MergedCellValues.Restart : W.MergedCellValues.Continue });
					if (cd.Borders.Any(b => b != null)) {
						var borders = new W.TableCellBorders();
						if (cd.Borders[0] is { } t) borders.Append(BorderOf<W.TopBorder>(t));
						if (cd.Borders[1] is { } l) borders.Append(BorderOf<W.LeftBorder>(l));
						if (cd.Borders[2] is { } b) borders.Append(BorderOf<W.BottomBorder>(b));
						if (cd.Borders[3] is { } r) borders.Append(BorderOf<W.RightBorder>(r));
						tcPr.Append(borders);
					}
					string? fill = cd.CfPat >= 0 && cd.Shading >= 10000 ? ColorHex(cd.CfPat) : cd.CbPat >= 0 ? ColorHex(cd.CbPat) : null;
					if (cd.CfPat >= 0 && cd.CbPat >= 0 && cd.Shading is > 0 and < 10000) fill = Blend(ColorHex(cd.CbPat) ?? "FFFFFF", ColorHex(cd.CfPat) ?? "000000", cd.Shading / 10000.0);
					if (fill != null) tcPr.Append(new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = fill });
					if (cd.VAlign != 't') tcPr.Append(new W.TableCellVerticalAlignment { Val = cd.VAlign == 'c' ? W.TableVerticalAlignmentValues.Center : W.TableVerticalAlignmentValues.Bottom });
					tc.Append(tcPr);
					foreach (var block in blocks) tc.Append(block);
					if (!blocks.Any() || blocks[^1] is W.Table) tc.Append(new W.Paragraph());
					tr.Append(tc);
					previous = tc;
				}
				table.Append(tr);
			}
			return table;
		}

		static string Blend(string back, string fore, double amount) {
			int Channel(string hex, int k) => Convert.ToInt32(hex.Substring(k * 2, 2), 16);
			return string.Concat(Enumerable.Range(0, 3).Select(k => ((int)Math.Round(Channel(back, k) * (1 - amount) + Channel(fore, k) * amount)).ToString("X2")));
		}

		T BorderOf<T>(Border b) where T : W.BorderType, new() => new T {
			Val = b.Style switch {
				"none" => W.BorderValues.None, "double" => W.BorderValues.Double, "thick" => W.BorderValues.Thick, "dotted" => W.BorderValues.Dotted,
				"dashed" => W.BorderValues.Dashed, "dashSmallGap" => W.BorderValues.DashSmallGap, "dotDash" => W.BorderValues.DotDash, "dotDotDash" => W.BorderValues.DotDotDash,
				"triple" => W.BorderValues.Triple, "wave" => W.BorderValues.Wave, "inset" => W.BorderValues.Inset, "outset" => W.BorderValues.Outset, _ => W.BorderValues.Single },
			Size = (uint)Math.Clamp((int)Math.Round(b.Width * 0.4), 2, 96),
			Space = (uint)Math.Clamp(b.Space / 20, 0, 31),
			Color = ColorHex(b.Color) ?? "auto",
		};

		// ---- one paragraph

		W.Paragraph BuildParagraph(Story st, ParaFmt f, CharFmt mark, W.SectionProperties? sectPr) {
			var style = StyleOf(f.Style) ?? StyleOf(0);
			var basisP = style?.P ?? new ParaFmt();
			var basisC = style?.C ?? new CharFmt();
			var numbering = Numbering(f, st.PnText);
			if (f.Pn == null && f.Ls == 0 && st.Part == main) { /* (a paragraph between list items: the list goes on after it) */ }
			var pPr = new W.ParagraphProperties();
			if (style != null && style.Number != 0) pPr.Append(new W.ParagraphStyleId { Val = style.Id });
			var clear = numbering is { } nb && abstractOfNum.TryGetValue(nb.Id, out int ab) && levelTabs.TryGetValue((ab, nb.Level), out var lt) ? lt : null;
			ParaDiff(pPr, f, basisP, numbering, forceIndent: numbering != null, clearTabs: clear);
			var markProps = RunDiff(mark, basisC);
			if (markProps.Count > 0) pPr.Append(new W.ParagraphMarkRunProperties(markProps));
			if (sectPr != null) pPr.Append(sectPr);
			var para = new W.Paragraph();
			if (pPr.HasChildren) para.Append(pPr);
			OpenXmlCompositeElement container = para;
			foreach (var item in st.Line) {
				switch (item.Kind) {
					case K.Run: {
						var run = NewRun(item.Fmt, basisC);
						foreach (var (type, text) in item.Pieces) {
							run.Append(type switch {
								't' => new W.Text(text) { Space = SpaceProcessingModeValues.Preserve },
								'T' => new W.TabChar(),
								'b' => new W.Break(),
								'p' => new W.Break { Type = W.BreakValues.Page },
								'c' => new W.Break { Type = W.BreakValues.Column },
								'h' => new W.NoBreakHyphen(),
								_ => new W.SoftHyphen(),
							});
						}
						container.Append(run);
						break;
					}
					case K.FldBegin: {
						var begin = new W.FieldChar { FieldCharType = W.FieldCharValues.Begin };
						if (item.Lock) begin.FieldLock = true;
						container.Append(NewRun(item.Fmt, basisC, begin));
						container.Append(NewRun(item.Fmt, basisC, new W.FieldCode(item.Text) { Space = SpaceProcessingModeValues.Preserve }));
						break;
					}
					case K.FldSep: container.Append(NewRun(item.Fmt, basisC, new W.FieldChar { FieldCharType = W.FieldCharValues.Separate })); break;
					case K.FldEnd: container.Append(NewRun(item.Fmt, basisC, new W.FieldChar { FieldCharType = W.FieldCharValues.End })); break;
					case K.BmStart: container.Append(new W.BookmarkStart { Id = item.Id.ToString(CultureInfo.InvariantCulture), Name = item.Text }); break;
					case K.BmEnd: container.Append(new W.BookmarkEnd { Id = item.Id.ToString(CultureInfo.InvariantCulture) }); break;
					case K.HlStart: {
						var link = new W.Hyperlink { History = true };
						if (item.Url != null) {
							try { link.Id = st.Part.AddHyperlinkRelationship(new Uri(item.Url, UriKind.RelativeOrAbsolute), true).Id; }
							catch { link.Anchor = item.Anchor ?? ""; }
						}
						if (item.Anchor != null) link.Anchor = item.Anchor;
						para.Append(link);
						container = link;
						break;
					}
					case K.HlEnd: container = para; break;
					case K.NoteRef:
						if (item.Note is { Id: > 0 } n)
							container.Append(NewRun(item.Fmt, basisC, n.End ? new W.EndnoteReference { Id = n.Id } : new W.FootnoteReference { Id = n.Id }));
						break;
					case K.NoteMark:
						container.Append(NewRun(item.Fmt, basisC, st.Part is EndnotesPart ? new W.EndnoteReferenceMark() : new W.FootnoteReferenceMark()));
						break;
					case K.Pic:
						if (item.Pic is { } pic && Drawing(st.Part, pic) is { } drawing) container.Append(NewRun(item.Fmt, basisC, drawing));
						break;
				}
			}
			return para;
		}

		W.Run NewRun(CharFmt f, CharFmt basis, params OpenXmlElement[] content) {
			var run = new W.Run();
			var props = RunDiff(f, basis);
			if (props.Count > 0) run.Append(new W.RunProperties(props));
			foreach (var e in content) run.Append(e);
			return run;
		}

		W.Drawing? Drawing(OpenXmlPart part, PicData pic) {
			var type = pic.Type switch { "jpeg" => ImagePartType.Jpeg, "emf" => ImagePartType.Emf, "wmf" => ImagePartType.Wmf, _ => ImagePartType.Png };
			ImagePart image = part switch {
				MainDocumentPart m => m.AddImagePart(type), HeaderPart h => h.AddImagePart(type), FooterPart f => f.AddImagePart(type),
				FootnotesPart fn => fn.AddImagePart(type), EndnotesPart en => en.AddImagePart(type), _ => throw new InvalidOperationException() };
			using (var stream = new MemoryStream(pic.Data)) image.FeedData(stream);
			string rel = part.GetIdOfPart(image);
			uint id = nextPicture++;
			string name = "Picture " + id;
			return new W.Drawing(new DW.Inline(
				new DW.Extent { Cx = pic.Cx, Cy = pic.Cy },
				new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
				new DW.DocProperties { Id = id, Name = name, Description = pic.Alt },
				new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
				new A.Graphic(new A.GraphicData(
					new PIC.Picture(
						new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties { Id = 0U, Name = name, Description = pic.Alt }, new PIC.NonVisualPictureDrawingProperties()),
						new PIC.BlipFill(new A.Blip { Embed = rel }, new A.Stretch(new A.FillRectangle())),
						new PIC.ShapeProperties(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = pic.Cx, Cy = pic.Cy }),
							new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
				) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
			) { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });
		}

		// ---- properties: what differs from the basis (the style, or the defaults)

		static OnOffValue? Off(bool on) => on ? null : new OnOffValue(false);

		static string Tw(int v) => v.ToString(CultureInfo.InvariantCulture);

		static W.Indentation Indent(int left, int right, int first) {
			var ind = new W.Indentation { Left = Tw(left), Right = Tw(right) };
			if (first >= 0) ind.FirstLine = Tw(first); else ind.Hanging = Tw(-first);
			return ind;
		}

		static (string Line, W.LineSpacingRuleValues Rule) LineOf(ParaFmt f) =>
			f.Sl == 0 ? ("240", W.LineSpacingRuleValues.Auto)
			: f.SlMult ? (Tw(Math.Abs(f.Sl)), W.LineSpacingRuleValues.Auto)
			: f.Sl > 0 ? (Tw(f.Sl), W.LineSpacingRuleValues.AtLeast) : (Tw(-f.Sl), W.LineSpacingRuleValues.Exact);

		void ParaDiff(OpenXmlCompositeElement pPr, ParaFmt f, ParaFmt b, (int Id, int Level)? numbering, bool forceIndent, List<int>? clearTabs = null) {
			if (f.KeepNext != b.KeepNext) pPr.Append(new W.KeepNext { Val = f.KeepNext });
			if (f.Keep != b.Keep) pPr.Append(new W.KeepLines { Val = f.Keep });
			if (f.PageBefore != b.PageBefore) pPr.Append(new W.PageBreakBefore { Val = f.PageBefore });
			if (f.Widow != b.Widow) pPr.Append(new W.WidowControl { Val = f.Widow });
			if (numbering is { } num) pPr.Append(new W.NumberingProperties(new W.NumberingLevelReference { Val = num.Level }, new W.NumberingId { Val = num.Id }));
			string BordersKey(ParaFmt x) => string.Join(";", x.Borders.Select(y => y?.Key ?? "-"));
			if (BordersKey(f) != BordersKey(b) && f.Borders.Any(x => x != null)) {
				var borders = new W.ParagraphBorders();
				if (f.Borders[0] is { } t) borders.Append(BorderOf<W.TopBorder>(t));
				if (f.Borders[1] is { } l) borders.Append(BorderOf<W.LeftBorder>(l));
				if (f.Borders[2] is { } bo) borders.Append(BorderOf<W.BottomBorder>(bo));
				if (f.Borders[3] is { } r) borders.Append(BorderOf<W.RightBorder>(r));
				pPr.Append(borders);
			}
			if (f.Shading != b.Shading && ColorHex(f.Shading) is { } shade) pPr.Append(new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = shade });
			var mine = f.Tabs.Select(t => t.Key).ToHashSet();
			var theirs = b.Tabs.Select(t => t.Key).ToHashSet();
			var levelOnly = (clearTabs ?? new List<int>()).Where(pos => !f.Tabs.Any(x => x.Pos == pos) && !b.Tabs.Any(x => x.Pos == pos)).Distinct().ToList();
			if (!mine.SetEquals(theirs) || levelOnly.Count > 0) {
				var tabs = new List<(int Pos, W.TabStop Stop)>();
				foreach (var t in f.Tabs.Where(t => !theirs.Contains(t.Key))) tabs.Add((t.Pos, TabOf(t)));
				foreach (var t in b.Tabs.Where(t => !mine.Contains(t.Key) && !f.Tabs.Any(x => x.Pos == t.Pos))) tabs.Add((t.Pos, new W.TabStop { Val = W.TabStopValues.Clear, Position = t.Pos }));
				foreach (int pos in levelOnly) tabs.Add((pos, new W.TabStop { Val = W.TabStopValues.Clear, Position = pos }));
				if (tabs.Count > 0) pPr.Append(new W.Tabs(tabs.OrderBy(t => t.Pos).Select(t => t.Stop)));
			}
			var spacing = new W.SpacingBetweenLines();
			if (f.Before != b.Before) spacing.Before = Tw(f.Before);
			if (f.After != b.After) spacing.After = Tw(f.After);
			var (line, rule) = LineOf(f);
			var (bline, brule) = LineOf(b);
			if (line != bline || rule != brule) { spacing.Line = line; spacing.LineRule = rule; }
			if (spacing.HasAttributes) pPr.Append(spacing);
			if (forceIndent || f.Left != b.Left || f.Right != b.Right || f.First != b.First) pPr.Append(Indent(f.Left, f.Right, f.First));
			if (f.Contextual != b.Contextual) pPr.Append(new W.ContextualSpacing { Val = f.Contextual });
			if (f.Align != b.Align) pPr.Append(new W.Justification { Val = f.Align switch {
				'c' => W.JustificationValues.Center, 'r' => W.JustificationValues.Right, 'j' => W.JustificationValues.Both, 'd' => W.JustificationValues.Distribute, _ => W.JustificationValues.Left } });
			if (f.Outline != b.Outline && f.Outline is >= 0 and <= 9) pPr.Append(new W.OutlineLevel { Val = f.Outline });
		}

		static W.TabStop TabOf(TabStop t) => new W.TabStop {
			Val = t.Kind switch { 'c' => W.TabStopValues.Center, 'r' => W.TabStopValues.Right, 'd' => W.TabStopValues.Decimal, 'b' => W.TabStopValues.Bar, _ => W.TabStopValues.Left },
			Leader = t.Leader switch { 'd' => W.TabStopLeaderCharValues.Dot, 'h' => W.TabStopLeaderCharValues.Hyphen, 'u' => W.TabStopLeaderCharValues.Underscore, 't' => W.TabStopLeaderCharValues.Heavy, 'm' => W.TabStopLeaderCharValues.MiddleDot, _ => W.TabStopLeaderCharValues.None },
			Position = t.Pos,
		};

		static readonly (string Hex, string Name)[] Highlights = {
			("FFFF00", "yellow"), ("00FF00", "green"), ("00FFFF", "cyan"), ("FF00FF", "magenta"), ("0000FF", "blue"), ("FF0000", "red"), ("000080", "darkBlue"), ("008080", "darkCyan"),
			("008000", "darkGreen"), ("800080", "darkMagenta"), ("800000", "darkRed"), ("808000", "darkYellow"), ("808080", "darkGray"), ("C0C0C0", "lightGray"), ("000000", "black"), ("FFFFFF", "white") };

		List<OpenXmlElement> RunDiff(CharFmt f, CharFmt b) {
			var list = new List<OpenXmlElement>();
			string name = FontName(f.Font);
			if (!string.Equals(name, FontName(b.Font), StringComparison.OrdinalIgnoreCase) && name.Length > 0)
				list.Add(new W.RunFonts { Ascii = name, HighAnsi = name, ComplexScript = name, EastAsia = name });
			if (f.B != b.B) { list.Add(new W.Bold { Val = Off(f.B) }); list.Add(new W.BoldComplexScript { Val = Off(f.B) }); }
			if (f.I != b.I) { list.Add(new W.Italic { Val = Off(f.I) }); list.Add(new W.ItalicComplexScript { Val = Off(f.I) }); }
			if (f.Caps != b.Caps) list.Add(new W.Caps { Val = Off(f.Caps) });
			if (f.SCaps != b.SCaps) list.Add(new W.SmallCaps { Val = Off(f.SCaps) });
			if (f.Strike != b.Strike) list.Add(new W.Strike { Val = Off(f.Strike) });
			if (f.DStrike != b.DStrike) list.Add(new W.DoubleStrike { Val = Off(f.DStrike) });
			if (f.Outline != b.Outline) list.Add(new W.Outline { Val = Off(f.Outline) });
			if (f.Shadow != b.Shadow) list.Add(new W.Shadow { Val = Off(f.Shadow) });
			if (f.Emboss != b.Emboss) list.Add(new W.Emboss { Val = Off(f.Emboss) });
			if (f.Imprint != b.Imprint) list.Add(new W.Imprint { Val = Off(f.Imprint) });
			if (f.Hidden != b.Hidden) list.Add(new W.Vanish { Val = Off(f.Hidden) });
			if (ColorHex(f.Color) != ColorHex(b.Color)) list.Add(new W.Color { Val = ColorHex(f.Color) ?? "auto" });
			if (f.Spacing != b.Spacing) list.Add(new W.Spacing { Val = f.Spacing });
			if (f.Scale != b.Scale) list.Add(new W.CharacterScale { Val = f.Scale });
			if (f.Position != b.Position) list.Add(new W.Position { Val = Tw(f.Position) });
			if (f.Size != b.Size) { list.Add(new W.FontSize { Val = Tw(f.Size) }); list.Add(new W.FontSizeComplexScript { Val = Tw(f.Size) }); }
			string? light = ColorHex(f.Highlight), blight = ColorHex(b.Highlight);
			string? namedLight = light == null ? null : Highlights.FirstOrDefault(h => h.Hex == light).Name;
			if (light != blight && light != null && namedLight != null) list.Add(new W.Highlight { Val = new W.HighlightColorValues(namedLight) });
			else if (light != blight && light == null) list.Add(new W.Highlight { Val = W.HighlightColorValues.None });
			if (f.Ul != b.Ul) list.Add(new W.Underline { Val = f.Ul == "" ? W.UnderlineValues.None : new W.UnderlineValues(f.Ul) });
			string? back = ColorHex(f.Back) ?? (light != null && namedLight == null ? light : null);
			string? bback = ColorHex(b.Back);
			if (back != bback) list.Add(new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = back ?? "auto" });
			if (f.Vert != b.Vert) list.Add(new W.VerticalTextAlignment { Val = f.Vert == 1 ? W.VerticalPositionValues.Superscript : f.Vert == 2 ? W.VerticalPositionValues.Subscript : W.VerticalPositionValues.Baseline });
			return list;
		}

		// ---- sections

		W.SectionProperties SectPr(SectFmt s) {
			var sp = new W.SectionProperties();
			foreach (var (header, type, id) in s.Refs) {
				var t = type switch { "first" => W.HeaderFooterValues.First, "even" => W.HeaderFooterValues.Even, _ => W.HeaderFooterValues.Default };
				sp.Append(header ? new W.HeaderReference { Type = t, Id = id } : new W.FooterReference { Type = t, Id = id });
			}
			sp.Append(new W.SectionType { Val = s.Break switch {
				"continuous" => W.SectionMarkValues.Continuous, "column" => W.SectionMarkValues.NextColumn, "even" => W.SectionMarkValues.EvenPage, "odd" => W.SectionMarkValues.OddPage, _ => W.SectionMarkValues.NextPage } });
			var size = new W.PageSize { Width = (uint)s.PgW, Height = (uint)s.PgH };
			if (s.Landscape) size.Orient = W.PageOrientationValues.Landscape;
			sp.Append(size);
			sp.Append(new W.PageMargin { Top = s.MT, Right = (uint)Math.Max(0, s.MR), Bottom = s.MB, Left = (uint)Math.Max(0, s.ML), Header = (uint)Math.Max(0, s.HeaderY), Footer = (uint)Math.Max(0, s.FooterY), Gutter = (uint)Math.Max(0, s.Gutter) });
			if (s.LineMod > 0) {
				var ln = new W.LineNumberType { CountBy = (short)s.LineMod };
				if (s.LineStart != 1) ln.Start = (short)(s.LineStart - 1);
				if (s.LineX > 0) ln.Distance = Tw(s.LineX);
				ln.Restart = s.LineRestart switch { "continuous" => W.LineNumberRestartValues.Continuous, "newSection" => W.LineNumberRestartValues.NewSection, _ => W.LineNumberRestartValues.NewPage };
				sp.Append(ln);
			}
			if (s.PgnRestart || s.PgnFmt != null) {
				var pg = new W.PageNumberType();
				if (s.PgnFmt != null) pg.Format = new W.NumberFormatValues(s.PgnFmt);
				if (s.PgnRestart) pg.Start = s.PgnStart;
				sp.Append(pg);
			}
			var cols = new W.Columns { Space = Tw(s.ColSpace) };
			if (s.Cols > 1) cols.ColumnCount = (short)s.Cols;
			if (s.ColLine) cols.Separator = true;
			sp.Append(cols);
			if (s.VAlign != null) sp.Append(new W.VerticalTextAlignmentOnPage { Val = new W.VerticalJustificationValues(s.VAlign) });
			if (s.TitlePg) sp.Append(new W.TitlePage());
			return sp;
		}

		// ---- the other parts

		void WriteStyles() {
			var part = main.AddNewPart<StyleDefinitionsPart>();
			var root = new W.Styles();
			var defC = new CharFmt();
			var defP = new ParaFmt();
			string font = FontName(-1);
			var runDefaults = new W.RunPropertiesBaseStyle();
			if (font.Length > 0) runDefaults.Append(new W.RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font });
			runDefaults.Append(new W.FontSize { Val = "24" }, new W.FontSizeComplexScript { Val = "24" }, new W.Languages { Val = "en-US" });
			root.Append(new W.DocDefaults(
				new W.RunPropertiesDefault(runDefaults),
				new W.ParagraphPropertiesDefault(new W.ParagraphPropertiesBaseStyle(new W.WidowControl { Val = false }, new W.SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }))));
			if (!styles.ContainsKey(0)) styles[0] = new StyleDef { Number = 0, Name = "Normal", Id = "Normal" };
			foreach (var def in styles.Values.OrderBy(d => d.Number)) {
				var basedOn = def.BasedOn >= 0 && def.BasedOn != def.Number ? StyleOf(def.BasedOn) : null;
				var style = new W.Style { Type = W.StyleValues.Paragraph, StyleId = def.Id };
				if (def.Number == 0) style.Default = true;
				style.Append(new W.StyleName { Val = def.Name });
				if (basedOn != null) style.Append(new W.BasedOn { Val = basedOn.Id });
				if (def.Next >= 0 && StyleOf(def.Next) is { } next) style.Append(new W.NextParagraphStyle { Val = next.Id });
				style.Append(new W.PrimaryStyle());
				var pPr = new W.StyleParagraphProperties();
				var basisP = basedOn?.P ?? defP;
				ParaDiff(pPr, def.P, basisP, null, forceIndent: false);
				var heading = Regex.Match(def.Name, @"^heading ([1-9])$", RegexOptions.IgnoreCase);
				if (heading.Success && def.P.Outline < 0 && !pPr.Elements<W.OutlineLevel>().Any()) pPr.Append(new W.OutlineLevel { Val = int.Parse(heading.Groups[1].Value) - 1 });
				if (pPr.HasChildren) style.Append(pPr);
				var rPr = RunDiff(def.C, basedOn?.C ?? defC);
				if (rPr.Count > 0) style.Append(new W.StyleRunProperties(rPr));
				root.Append(style);
			}
			// (the notes' reference and text styles Word expects, when there are notes)
			if (footnotes != null || endnotes != null) {
				foreach (var (id, name) in new[] { ("FootnoteReference", "footnote reference"), ("EndnoteReference", "endnote reference") }) {
					if (root.Elements<W.Style>().Any(s => s.StyleId == id)) continue;
					root.Append(new W.Style(new W.StyleName { Val = name }, new W.UIPriority { Val = 99 }, new W.UnhideWhenUsed(),
						new W.StyleRunProperties(new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript })) { Type = W.StyleValues.Character, StyleId = id });
				}
			}
			// (the default character, table and list styles every Word document has: Word adds what is missing, and the
			// document then counts as changed)
			root.Append(new W.Style(new W.StyleName { Val = "Default Paragraph Font" }, new W.UIPriority { Val = 1 }, new W.SemiHidden(), new W.UnhideWhenUsed())
				{ Type = W.StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true });
			root.Append(new W.Style(new W.StyleName { Val = "Normal Table" }, new W.UIPriority { Val = 99 }, new W.SemiHidden(), new W.UnhideWhenUsed(),
				new W.StyleTableProperties(new W.TableIndentation { Width = 0, Type = W.TableWidthUnitValues.Dxa },
					new W.TableCellMarginDefault(new W.TopMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa }, new W.TableCellLeftMargin { Width = 108, Type = W.TableWidthValues.Dxa },
						new W.BottomMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa }, new W.TableCellRightMargin { Width = 108, Type = W.TableWidthValues.Dxa })))
				{ Type = W.StyleValues.Table, StyleId = "TableNormal", Default = true });
			root.Append(new W.Style(new W.StyleName { Val = "No List" }, new W.UIPriority { Val = 99 }, new W.SemiHidden(), new W.UnhideWhenUsed())
				{ Type = W.StyleValues.Numbering, StyleId = "NoList", Default = true });
			root.Append(new W.Style(new W.StyleName { Val = "Hyperlink" }, new W.UIPriority { Val = 99 }, new W.UnhideWhenUsed(),
				new W.StyleRunProperties(new W.Color { Val = "0563C1" }, new W.Underline { Val = W.UnderlineValues.Single })) { Type = W.StyleValues.Character, StyleId = "Hyperlink" });
			part.Styles = root;
			part.Styles.Save();
		}

		void WriteNumbering() {
			if (nums.Count == 0) return;
			var part = main.AddNewPart<NumberingDefinitionsPart>();
			var numbering = new W.Numbering();
			foreach (var a in abstracts) numbering.Append(a);
			foreach (var n in nums) numbering.Append(n);
			part.Numbering = numbering;
			part.Numbering.Save();
		}

		void WriteSettings() {
			var part = main.AddNewPart<DocumentSettingsPart>();
			var settings = new W.Settings();
			settings.Append(new W.Zoom { Percent = "100" });
			settings.Append(new W.DefaultTabStop { Val = (short)Math.Clamp(defTab, 1, short.MaxValue) });
			if (facing) settings.Append(new W.EvenAndOddHeaders());
			settings.Append(new W.CharacterSpacingControl { Val = W.CharacterSpacingValues.DoNotCompress });
			settings.Append(new W.Compatibility(new W.CompatibilitySetting { Name = W.CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" }));
			part.Settings = settings;
			part.Settings.Save();
		}
	}
}
