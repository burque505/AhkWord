#:package DocumentFormat.OpenXml@3.5.1
#:package DocumentFormat.OpenXml.Framework@3.5.1
#:package HtmlToOpenXml.dll@3.5.0
#:package AngleSharp@1.5.0
#:package Microsoft.Extensions.Logging.Abstractions@6.0.0
#:package Codeuctivity.OpenXmlPowerTools@10.0.621
#:package SkiaSharp@4.152.0
// Word and OpenDocument conversions for RTE Writer (rte_writer34.cs, through #load).
// Version 7: the built-in reader's paragraphs get the .docx's own line spacing, space before and after, indents and tab
// stops, as Word works them out (DocxLineSpacing.IntoHtml, in rte_linespacing2.cs): the PowerTools read a paragraph
// without spacing of its own as 1.08 lines, and "at least" as exactly. All of a paragraph's tabs come back as tabs (the
// PowerTools write those after the first as a margin on a no-break space), and its empty paragraphs stay.
// Version 6: pictures keep their alt text in RTF, as the picture property Word and LibreOffice use for it
// ({\*\picprop{\sp{\sn wzDescription}{\sv ...}}}): WithAltTexts writes it, Pictures reads it, PngPictures keeps it.
// Version 5: Export takes a last change to the RTF given to Word and LibreOffice (RTE Writer adds the header, footer
// and page layout there, rte_headerfooter.cs, after the other changes).
// Version 4: page breaks. Word and LibreOffice get them in the RTF (\page); the openxml converter writes the editor's
// HTML page breaks (rte_pagebreaks.cs) as Word page breaks (HtmlToOpenXml does that itself), and reads Word's
// (<w:br w:type="page"/>) as <br style="page-break-before: always">, which the PowerTools would write as a plain <br>.
// (LibreOffice's and Word's RTF give a page break as \pagebb on the paragraph after it, which RichEdit ignores:
// PageBreaks.FromPageBreakBefore makes that a \page when the editor loads RTF.) Word repaginates before saving: it
// had left list items after the first page unnumbered.
// Version 3: Word and LibreOffice get legal lists as real list numbering too (LegalRtf): the RTF they are given has
// the legal items' number text replaced by an RTF list (\ls, \ilvl) with legal level texts (%1.%2.%3), so the documents
// they write renumber themselves when edited.
// Version 2: legal numbering (rte_legal.cs) with the openxml converter: a legal list (<ol class="rte-legal">) becomes
// one Word list numbered %1.%2.%3, and Word's compound numbers ("1.2") come back as a legal list. (Word and
// LibreOffice get the numbers as text in the RTF, which they show as they are.)
//
// RTE Writer's own format for these is RTF ((RichTextBox)editor).Rtf, which keeps everything RichEdit has: fonts,
// colours, alignment, lists, tables and pictures. A converter turns that RTF into .docx or .odt, or a .docx/.odt into
// RTF, using whichever of these is installed:
//   Word         - an invisible Word.Application through COM, and only when Word is NOT already running (an automation
//                  client can attach to the user's Word, and quitting it would close the user's documents)
//   LibreOffice  - soffice --headless --convert-to, with its own profile in %LOCALAPPDATA%\RTE Writer, so it never
//                  touches the user's LibreOffice settings or a LibreOffice that is running
//   openxml      - .docx only, from and to the editor's HTML, with NuGet packages and nothing installed: HtmlToOpenXml
//                  writes .docx (its list numbering is repaired here), Codeuctivity.OpenXmlPowerTools
//                  (WmlToHtmlConverter) reads it. Keeps colours, alignment, lists, tables and images.
//   pandoc       - from and to the editor's HTML; the last resort, since it drops colours and alignment
// Csharp-in-AHK takes packages from the local NuGet cache and doesn't resolve dependencies, so each one is listed above;
// they are restored once with dotnet restore (any project referencing HtmlToOpenXml.dll and Codeuctivity.OpenXmlPowerTools).
// For .docx the order is Word, openxml, LibreOffice, pandoc, writing and reading. For .odt it is LibreOffice, Word, pandoc
// writing, and Word, LibreOffice, pandoc reading: RichEdit makes real lists of Word's RTF, but only text ("1.<tab>") of
// LibreOffice's. A caller can name the converters to use ("word", "openxml", "libreoffice", "pandoc", comma-separated);
// the RTE_WRITER_CONVERTERS environment variable does the same for all calls.

using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Packaging;
using Codeuctivity.OpenXmlPowerTools;
using Codeuctivity.OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter;
using W = DocumentFormat.OpenXml.Wordprocessing;

public static class OfficeConverters {
	public const string Word = "word", LibreOffice = "libreoffice", OpenXml = "openxml", Pandoc = "pandoc";
	static readonly string[] All = { Word, LibreOffice, OpenXml, Pandoc };

	// ---- what is installed

	public static bool WordInstalled => Type.GetTypeFromProgID("Word.Application") != null;
	public static bool WordRunning => Process.GetProcessesByName("WINWORD").Length > 0;

	public static string? LibreOfficePath {
		get {
			var candidates = new List<string?>();
			foreach (var root in new[] { Registry.LocalMachine, Registry.CurrentUser }) {
				try {
					using var key = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe");
					candidates.Add(key?.GetValue(null) as string);
				}
				catch { }
			}
			foreach (var folder in new[] { Environment.SpecialFolder.ProgramFiles, Environment.SpecialFolder.ProgramFilesX86 })
				candidates.Add(Path.Combine(Environment.GetFolderPath(folder), "LibreOffice", "program", "soffice.exe"));
			return candidates.FirstOrDefault(p => !string.IsNullOrEmpty(p) && File.Exists(p.Trim('"')))?.Trim('"');
		}
	}

	public static string? PandocPath {
		get {
			var folders = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries).ToList();
			folders.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Pandoc"));
			folders.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Pandoc"));
			foreach (var folder in folders) {
				try {
					string path = Path.Combine(folder.Trim().Trim('"'), "pandoc.exe");
					if (File.Exists(path)) return path;
				}
				catch { }
			}
			return null;
		}
	}

	// One line per converter, for the About box and Rte.Converters().
	public static string Describe() {
		string word = !WordInstalled ? "not installed" : WordRunning ? "installed, but running (not used while it is open)" : "installed";
		return $"Word: {word}\nLibreOffice: {LibreOfficePath ?? "not found"}\nOpen XML (built in, .docx only): {(OpenXmlAvailable ? "available" : "packages missing")}\npandoc: {PandocPath ?? "not found"}";
	}

	// ---- conversions

	// Writes the document as .docx or .odt (by the extension of output). rtf is the document's RTF; html (only made if
	// pandoc is needed) its HTML. Returns the name of the converter that did it.
	public static string Export(string output, string rtf, Func<string> html, string? via = null, Func<string, string>? finish = null) {
		finish ??= r => r;
		string format = Path.GetExtension(output).ToLowerInvariant().TrimStart('.');
		if (format != "docx" && format != "odt") throw new ArgumentException("Export writes .docx or .odt, not " + Path.GetFileName(output));
		var order = Order(format == "docx" ? new[] { Word, OpenXml, LibreOffice, Pandoc } : new[] { LibreOffice, Word, Pandoc }, via);
		rtf = PngPictures(rtf);    // (RichEdit keeps pictures as WMF, which LibreOffice kept as it was)
		string temp = TempFolder();
		var errors = new List<string>();
		try {
			string? legalRtf = null;    // the RTF with legal lists as RTF lists, made when Word or LibreOffice needs it
			foreach (var converter in order) {
				try {
					string result = converter switch {
						Word => WordConvert(WriteRtf(temp, "document-word.rtf", finish(legalRtf ??= WithNormalStyle(LegalRtf(rtf)))), Path.Combine(temp, "word." + format), format == "docx" ? 16 : 23),    // wdFormatDocumentDefault, wdFormatOpenDocumentText
						LibreOffice => LibreOfficeConvert(WriteRtf(temp, "document-lo.rtf", finish(ModernBullets(legalRtf ??= WithNormalStyle(LegalRtf(rtf))))), temp, format == "docx" ? "docx:MS Word 2007 XML" : "odt", format),
						OpenXml => format == "docx" ? OpenXmlWrite(html(), Path.Combine(temp, "openxml.docx")) : throw new InvalidOperationException("writes only .docx"),
						_ => PandocConvert(WriteHtml(temp, html()), Path.Combine(temp, "pandoc." + format), "html", format, temp),
					};
					File.Copy(result, output, overwrite: true);
					return converter;
				}
				catch (Exception ex) { errors.Add(Name(converter) + ": " + Message(ex)); }
			}
		}
		finally { TryDelete(temp); }
		throw new InvalidOperationException("Could not write " + Path.GetFileName(output) + ".\n" + string.Join("\n", errors));
	}

	// Reads a .docx or .odt (or .doc). Returns ("rtf", text) from Word or LibreOffice, or ("html", text) from pandoc,
	// with the converter's name.
	public static (string Kind, string Content, string Converter) Import(string input, string? via = null) {
		string format = Path.GetExtension(input).ToLowerInvariant().TrimStart('.');
		// (.odt: LibreOffice first, as for writing: Word's RTF of an .odt leaves its text's font to whichever paragraph set one last)
		var order = Order(format == "docx" ? new[] { Word, OpenXml, LibreOffice, Pandoc } : format == "odt" ? new[] { LibreOffice, Word, Pandoc } : new[] { Word, LibreOffice, Pandoc }, via);
		if (format == "doc") order.Remove(Pandoc);    // pandoc can't read the old binary format
		string temp = TempFolder();
		var errors = new List<string>();
		try {
			// a copy, so the converter never locks or changes the user's file (and Word doesn't add it to its recent files)
			string copy = Path.Combine(temp, "input." + format);
			File.Copy(input, copy);
			foreach (var converter in order) {
				try {
					return converter switch {
						Word => ("rtf", File.ReadAllText(WordConvert(copy, Path.Combine(temp, "word.rtf"), 6), System.Text.Encoding.Latin1), converter),    // wdFormatRTF
						LibreOffice => LibreOfficeImport(copy, temp, format),
						OpenXml => format == "docx" ? ("html", OpenXmlRead(copy), converter) : throw new InvalidOperationException("reads only .docx"),
						_ => ("html", File.ReadAllText(PandocConvert(copy, Path.Combine(temp, "pandoc.html"), format, "html", temp, "--embed-resources", "--standalone")), converter),
					};
				}
				catch (Exception ex) { errors.Add(Name(converter) + ": " + Message(ex)); }
			}
		}
		finally { TryDelete(temp); }
		throw new InvalidOperationException("Could not read " + Path.GetFileName(input) + ".\n" + string.Join("\n", errors));
	}

	// Reading with LibreOffice: it writes .docx, which the built-in reader turns into HTML with real lists (LibreOffice's
	// RTF has its lists as typed numbers, which RichEdit keeps as text). RTF is the fallback (and for .docx input, which
	// the built-in reader has had its turn at already).
	static (string Kind, string Content, string Converter) LibreOfficeImport(string copy, string temp, string format) {
		if (format != "docx" && OpenXmlAvailable) {
			try {
				string docx = LibreOfficeConvert(copy, temp, "docx:MS Word 2007 XML", "docx");
				return ("html", OpenXmlRead(docx), LibreOffice);
			} catch { }
		}
		return ("rtf", File.ReadAllText(LibreOfficeConvert(copy, temp, "rtf", "rtf"), System.Text.Encoding.Latin1), LibreOffice);
	}

	// ---- Word

	// Converts with Word on a thread of its own, so that a Word that stops answering (a message it shows while hidden)
	// can be given up on: after 90 seconds that Word is ended and the conversion fails, instead of RTE Writer hanging.
	static string WordConvert(string input, string output, int fileFormat) {
		if (!WordInstalled) throw new InvalidOperationException("not installed");
		EndLeftoverWord();
		if (WordRunning) throw new InvalidOperationException("Word is open; close it to use it for conversions");
		var since = DateTime.Now.AddSeconds(-2);    // Word processes started from here on are ours (Word wasn't running)
		Exception? error = null;
		var thread = new Thread(() => {
			try { WordConvertOnThread(input, output, fileFormat, since); }
			catch (Exception ex) { error = ex; }
		}) { IsBackground = true, Name = "RTE Writer: Word conversion" };
		thread.SetApartmentState(ApartmentState.STA);
		thread.Start();
		if (!thread.Join(TimeSpan.FromSeconds(90))) {
			int ended = EndOurWord(since, wait: false);
			Log($"Word conversion of {Path.GetFileName(input)} took more than 90 seconds: ended {ended} Word process(es)");
			throw new TimeoutException("Word did not finish within 90 seconds (it may have been showing a message), so it was closed");
		}
		if (error != null) throw error is TargetInvocationException { InnerException: { } inner } ? inner : error;
		return output;
	}

	static void WordConvertOnThread(string input, string output, int fileFormat, DateTime since) {
		object? word = null, documents = null, document = null;
		try {
			word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!)!;
			Remember(since);
			Set(word, "Visible", false);
			Set(word, "DisplayAlerts", 0);    // wdAlertsNone
			try { Set(word, "AutomationSecurity", 3); } catch { }    // msoAutomationSecurityForceDisable: no macros run
			documents = Get(word, "Documents");
			// Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
			document = Call(documents, "Open", input, false, true, false);
			// an invisible Word makes real lists of RTF's old \pn numbering only where it has laid the pages out: without
			// this, list items after the first page break (or the first full page) lose their numbering
			Call(document, "Repaginate");
			Call(document, "SaveAs2", output, fileFormat);
			Call(document, "Close", false);
			Release(ref document);
		}
		finally {
			if (document != null) { try { Call(document, "Close", false); } catch { } Release(ref document); }
			Release(ref documents);
			if (word != null) {
				try { Call(word, "Quit", false); } catch { }
				Release(ref word);
			}
			// always, also when Word failed while starting (it may be running all the same)
			int ended = EndOurWord(since);
			Log($"Word converted {Path.GetFileName(input)} -> {Path.GetExtension(output)}: {(File.Exists(output) ? "done" : "failed")}{(ended > 0 ? $", ended {ended} Word process(es) that had not quit" : "")}");
		}
		if (!File.Exists(output)) throw new InvalidOperationException("Word did not write the file");
	}

	// ---- Words that RTE Writer starts: a record of them, so that one left running (RTE Writer closed or killed during
	// a conversion) is known to be RTE Writer's, and is ended by the next conversion instead of being taken for the
	// user's Word. The user's own Word is never in the record.

	static string StateFolder => AppFolders.Local;    // (%LOCALAPPDATA%\RTE Writer, or beside the program in portable mode)
	static string WordRecord => Path.Combine(StateFolder, "word-processes.txt");

	/// The converters' log (Word conversions and the Words they end).
	public static string LogFile => Path.Combine(StateFolder, "converters.log");

	static void Log(string text) {
		try {
			Directory.CreateDirectory(StateFolder);
			if (File.Exists(LogFile) && new FileInfo(LogFile).Length > 200_000) File.Delete(LogFile);
			File.AppendAllText(LogFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{Environment.ProcessId}] {text}\r\n");
		} catch { }
	}

	static string Describe(IEnumerable<Process> processes) => string.Join(", ", processes.Select(p => {
		try { return $"{p.Id} (started {p.StartTime:HH:mm:ss.fff}, window {p.MainWindowHandle})"; } catch (Exception ex) { return $"{p.Id} ({ex.Message})"; }
	})) is { Length: > 0 } list ? list : "none";

	static IEnumerable<Process> OursSince(DateTime since) => Process.GetProcessesByName("WINWORD").Where(p => {
		try { return p.StartTime >= since && p.MainWindowHandle == IntPtr.Zero; } catch { return false; }
	});

	// Records the Words started for this conversion (process id and start time).
	static void Remember(DateTime since) {
		try {
			Directory.CreateDirectory(StateFolder);
			var lines = OursSince(since).Select(p => $"{p.Id} {p.StartTime.Ticks}").ToList();
			if (lines.Count == 0) return;
			File.AppendAllLines(WordRecord, lines);
			Log("Word started: process " + string.Join(", ", lines.Select(l => l.Split(' ')[0])));
		} catch { }
	}

	// Ends Words from the record that are still running without a window (left over from an earlier conversion), and
	// forgets the ones that have gone.
	static void EndLeftoverWord() {
		try {
			if (!File.Exists(WordRecord)) return;
			var keep = new List<string>();
			foreach (var line in File.ReadAllLines(WordRecord)) {
				var parts = line.Split(' ');
				if (parts.Length != 2 || !int.TryParse(parts[0], out int id) || !long.TryParse(parts[1], out long ticks)) continue;
				Process? process = null;
				try { process = Process.GetProcessById(id); } catch (ArgumentException) { continue; }    // gone
				using (process) {
					bool same = false;
					try { same = process.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase) && process.StartTime.Ticks == ticks; } catch { }
					if (!same) continue;    // the id now belongs to another process
					if (process.MainWindowHandle != IntPtr.Zero) { keep.Add(line); continue; }    // (it has a window: leave it)
					try { process.Kill(); process.WaitForExit(5_000); Log($"ended Word process {id}, left running by an earlier conversion"); } catch { keep.Add(line); }
				}
			}
			File.WriteAllLines(WordRecord, keep);
		} catch { }
	}

	// After Quit: waits up to 15 seconds for the Word processes started since the conversion began to exit, then ends
	// the ones still there that have no window. They are ours: Word wasn't running when the conversion began, and a Word
	// the user opens meanwhile has a window. (Word can start a second process, so this looks for all of them, not just
	// the first one.) Otherwise a hidden Word would stay, and RTE Writer would take it for the user's.
	static int EndOurWord(DateTime since, bool wait = true) {
		var clock = Stopwatch.StartNew();
		Log("Word processes after the conversion: " + Describe(Process.GetProcessesByName("WINWORD")));
		var left = OursSince(since).ToList();
		while (wait && left.Count > 0 && clock.ElapsedMilliseconds < 15_000) {
			Thread.Sleep(250);
			foreach (var process in left) process.Dispose();
			left = OursSince(since).ToList();
		}
		int ended = 0;
		foreach (var process in left) {
			try { process.Kill(); process.WaitForExit(5_000); ended++; } catch { }
			process.Dispose();
		}
		return ended;
	}


	static object? Get(object target, string name) => target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, null);
	static void Set(object target, string name, object value) => target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, new[] { value });
	static object? Call(object? target, string name, params object[] args) => target!.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args);
	static void Release(ref object? com) {
		if (com != null && Marshal.IsComObject(com)) Marshal.FinalReleaseComObject(com);
		com = null;
	}

	// ---- LibreOffice

	public static string LibreOfficeProfile => Path.Combine(AppFolders.Local, "libreoffice-profile");

	static string LibreOfficeConvert(string input, string folder, string filter, string extension) {
		string soffice = LibreOfficePath ?? throw new InvalidOperationException("not found");
		string outFolder = Path.Combine(folder, "libreoffice");
		Directory.CreateDirectory(outFolder);
		Directory.CreateDirectory(LibreOfficeProfile);
		string output = Path.Combine(outFolder, Path.GetFileNameWithoutExtension(input) + "." + extension);
		var start = new ProcessStartInfo(soffice) { UseShellExecute = false, CreateNoWindow = true };
		foreach (var arg in new[] { "-env:UserInstallation=" + new Uri(LibreOfficeProfile).AbsoluteUri, "--headless", "--invisible", "--nologo",
			"--norestore", "--nolockcheck", "--convert-to", filter, "--outdir", outFolder, input })
			start.ArgumentList.Add(arg);
		using (var process = Process.Start(start)!) {
			if (!process.WaitForExit(180_000)) { try { process.Kill(true); } catch { } throw new TimeoutException("took more than 3 minutes"); }
		}
		// soffice.exe is a launcher that can return before the real program (soffice.bin) has written the file
		var clock = Stopwatch.StartNew();
		long size = -1;
		while (clock.ElapsedMilliseconds < 120_000) {
			if (File.Exists(output)) {
				long now = new FileInfo(output).Length;
				if (now > 0 && now == size && CanOpen(output)) return output;
				size = now;
			}
			else if (clock.ElapsedMilliseconds > 20_000 && !Process.GetProcessesByName("soffice.bin").Any()) break;
			Thread.Sleep(250);
		}
		throw new InvalidOperationException("LibreOffice did not write the file");
	}

	static bool CanOpen(string path) {
		try { using var _ = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None); return true; }
		catch (IOException) { return false; }
	}

	// ---- Open XML (NuGet packages; nothing to install)

	static bool OpenXmlAvailable {
		get {
			try { return typeof(HtmlToOpenXml.HtmlConverter).Assembly != null && typeof(WmlToHtmlConverter).Assembly != null; }
			catch { return false; }
		}
	}

	/// The font of text without one of its own, for the openxml converter (RTE Writer sets it to the editor's font).
	public static (string Name, float Points)? DefaultFont;

	static string OpenXmlWrite(string html, string output) {
		using (var package = WordprocessingDocument.Create(output, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)) {
			var main = package.AddMainDocumentPart();
			main.Document = new W.Document(new W.Body());
			var converter = new HtmlToOpenXml.HtmlConverter(main);
			// on a thread of its own: waiting for it here, on the window's thread, could otherwise deadlock
			html = KeepBlanksAndTabs(HeadingSizes(UnquotedFonts(html)));
			Task.Run(() => converter.ParseBody(ListStyles(html), CancellationToken.None)).GetAwaiter().GetResult();
			BlanksAndTabs(main);
			HeadingFonts(main);
			RepairNumbering(main, html);
			MarksFromRuns(main);
			// (a styles part, which the document's styles go into: DocxStyles.IntoDocx)
			if (main.StyleDefinitionsPart == null) {
				var part = main.AddNewPart<StyleDefinitionsPart>();
				part.Styles = new W.Styles();
				part.Styles.Save();
			}
			// the editor's HTML has no font for ordinary text: the document's default font is the editor's
			if (DefaultFont is { } font && main.StyleDefinitionsPart?.Styles is { } styles) {
				var defaults = styles.DocDefaults ??= new W.DocDefaults();
				var runDefaults = defaults.RunPropertiesDefault ??= new W.RunPropertiesDefault();
				var run = runDefaults.RunPropertiesBaseStyle ??= new W.RunPropertiesBaseStyle();
				run.RunFonts = new W.RunFonts { Ascii = font.Name, HighAnsi = font.Name, ComplexScript = font.Name, EastAsia = font.Name };
				run.FontSize = new W.FontSize { Val = ((int)Math.Round(font.Points * 2)).ToString() };
				styles.Save();
			}
			main.Document.Save();
		}
		return output;
	}

	// HtmlToOpenXml keeps the quotes of font-family: 'Calibri Light' and cuts the name short ('Calibri Lig): one font
	// name, without quotes.
	static string UnquotedFonts(string html) =>
		Regex.Replace(html, @"font-family:\s*(?:'([^']+)'|&quot;([^&]+)&quot;|&#39;([^&]+)&#39;|""([^""]+)"")[^;""]*", m => "font-family: " + string.Concat(m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value, m.Groups[4].Value));

	// Headings: the editor's sizes (24, 18, 14 ... points); HtmlToOpenXml's Heading styles have Word's (14 points for
	// Heading 1), and the editor's HTML gives headings no size of their own.
	static string HeadingSizes(string html) => Regex.Replace(html, @"<h([1-6])(\s[^>]*)?>", m => {
		string attributes = m.Groups[2].Value;
		if (Regex.IsMatch(attributes, @"font-size", RegexOptions.IgnoreCase)) return m.Value;
		string size = $"font-size: {MarkdownConverter.HeadingSize(int.Parse(m.Groups[1].Value)).ToString(System.Globalization.CultureInfo.InvariantCulture)}pt";
		var style = Regex.Match(attributes, @"\bstyle\s*=\s*""([^""]*)""");
		return style.Success
			? $"<h{m.Groups[1].Value}{attributes.Remove(style.Index, style.Length).Insert(style.Index, $"style=\"{size}; {style.Groups[1].Value}\"")}>"
			: $"<h{m.Groups[1].Value}{attributes} style=\"{size}\">";
	}, RegexOptions.IgnoreCase);

	// HtmlToOpenXml's Heading styles use the theme's heading font, and its documents have no theme (Word and LibreOffice
	// then fell back to a serif font): the editor's font instead.
	// Nor are the editor's headings italic (Heading 4's is), and they are all bold (Heading 5's isn't): italics are in
	// the text where it has them.
	static void HeadingFonts(MainDocumentPart main) {
		if (main.StyleDefinitionsPart?.Styles is not { } styles) return;
		foreach (var style in styles.Elements<W.Style>().Where(s => s.StyleId?.Value?.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) == true)) {
			var run = style.StyleRunProperties;
			if (run == null) continue;
			if (DefaultFont is { } font) run.RunFonts = new W.RunFonts { Ascii = font.Name, HighAnsi = font.Name, ComplexScript = font.Name, EastAsia = font.Name };
			run.Color = null;    // (the styles' own blue: the editor's headings have no colour but the one given them, which is in the text)
			run.Italic = null;
			run.ItalicComplexScript = null;
			run.Bold ??= new W.Bold();
			run.BoldComplexScript ??= new W.BoldComplexScript();
		}
		styles.Save();
	}

	// HtmlToOpenXml leaves out empty paragraphs (Word's blank lines) and makes tabs spaces: they go through it as text
	// (KeepBlanksAndTabs), and come back as empty paragraphs and tabs (BlanksAndTabs).
	const string EmptyMark = "@@RTE-EMPTY@@", TabMark = "@@RTE-TAB@@";

	static string KeepBlanksAndTabs(string html) {
		int body = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
		if (body < 0) body = 0;
		string text = Regex.Replace(html[body..], @"(<(p|h[1-6])\b[^>]*>)\s*(</\2\s*>)", m => m.Groups[1].Value + EmptyMark + m.Groups[3].Value, RegexOptions.IgnoreCase);
		text = Regex.Replace(text, @"(?<=>)[^<]+", m => m.Value.Replace("&#9;", TabMark).Replace("&#x9;", TabMark).Replace("&#x09;", TabMark).Replace("\t", TabMark));
		return html[..body] + text;
	}

	static void BlanksAndTabs(MainDocumentPart main) {
		if (main.Document?.Body is not { } body) return;
		foreach (var text in body.Descendants<W.Text>().Where(t => t.Text.Contains(EmptyMark) || t.Text.Contains(TabMark)).ToList()) {
			var pieces = text.Text.Replace(EmptyMark, "").Split(TabMark);
			DocumentFormat.OpenXml.OpenXmlElement at = text;
			for (int i = 0; i < pieces.Length; i++) {
				if (i > 0) at = at.InsertAfterSelf(new W.TabChar());
				if (pieces[i] != "") at = at.InsertAfterSelf(new W.Text(pieces[i]) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
			}
			var run = text.Parent;
			text.Remove();
			if (run is W.Run r && !r.ChildElements.Any(e => e is not W.RunProperties)) r.Remove();
		}
		main.Document.Save();
	}

	// HtmlToOpenXml takes a list's numbering from CSS, not from <ol type="a">.
	static string ListStyles(string html) => Regex.Replace(html, @"<(ol|ul)\b([^>]*)>", m => {
		string attributes = m.Groups[2].Value;
		if (Regex.IsMatch(attributes, @"\bstyle\s*=", RegexOptions.IgnoreCase)) return m.Value;
		string type = Regex.Match(attributes, @"\btype\s*=\s*""?([aAiI1])").Groups[1].Value;
		string style = m.Groups[1].Value.ToLowerInvariant() == "ul" ? "disc" : type switch {
			"a" => "lower-alpha", "A" => "upper-alpha", "i" => "lower-roman", "I" => "upper-roman", _ => "decimal" };
		return $"<{m.Groups[1].Value}{attributes} style=\"list-style-type: {style}\">";
	}, RegexOptions.IgnoreCase);

	// HtmlToOpenXml (3.5) gives its built-in list templates abstract numbering ids that its own lists also use, so lists
	// get the wrong numbering (lower-roman shows upper-roman, upper-roman shows bullets), and it leaves out punctuation
	// such as "a)". Each list paragraph it makes is one <li> of the HTML, in order, so every list (w:num) gets a
	// numbering definition of its own here, made from its <ol>: style, punctuation (rte_htmllists.cs classes) and start.
	sealed record HtmlList(string Format, string Punctuation, int Start, int Root = -1, int Depth = 0);    // Root: which legal list (-1: none)

	static void RepairNumbering(MainDocumentPart main, string html) {
		var numbering = main.NumberingDefinitionsPart?.Numbering;
		if (numbering == null) return;
		// the list of each <li>, in order
		var items = new List<HtmlList>();
		var sizes = new List<float>();    // each <li>'s font size (its number's), from its style
		var lists = new Stack<HtmlList>();
		int roots = 0;
		foreach (Match tag in Regex.Matches(html, @"<(/?)(ol|ul|li)\b([^>]*)>", RegexOptions.IgnoreCase)) {
			string name = tag.Groups[2].Value.ToLowerInvariant(), attributes = tag.Groups[3].Value;
			if (tag.Groups[1].Value == "/") { if (name != "li" && lists.Count > 0) lists.Pop(); continue; }
			if (name == "li") {
				if (lists.Count == 0) continue;
				items.Add(lists.Peek());
				var size = Regex.Match(attributes, @"font-size:\s*([\d.]+)pt");
				sizes.Add(size.Success ? float.Parse(size.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) : DefaultFont?.Points ?? 11);
				continue;
			}
			var start = Regex.Match(attributes, @"\bstart\s*=\s*""?(-?\d+)");
			int first = start.Success ? int.Parse(start.Groups[1].Value) : 1;
			if (name == "ol" && Regex.IsMatch(attributes, @"\bclass\s*=\s*""[^""]*\brte-legal\b")) {
				var parent = lists.Count > 0 ? lists.Peek() : null;
				lists.Push(parent is { Root: >= 0 } ? new HtmlList("legal", "", parent.Start, parent.Root, parent.Depth + 1) : new HtmlList("legal", "", first, roots++, 0));
				continue;
			}
			string type = Regex.Match(attributes, @"\btype\s*=\s*""?([aAiI1])").Groups[1].Value;
			string format = name == "ul" ? "bullet" : type switch { "a" => "lowerLetter", "A" => "upperLetter", "i" => "lowerRoman", "I" => "upperRoman", _ => "decimal" };
			string punctuation = Regex.Match(attributes, @"\brte-[a-z]+(?:-[a-z]+)?-(paren|parens|plain)\b").Groups[1].Value;
			lists.Push(new HtmlList(format, punctuation, first));
		}
		var paragraphs = main.Document.Body!.Descendants<W.Paragraph>().Select(p => p.ParagraphProperties?.NumberingProperties)
			.Where(n => n?.NumberingId?.Val != null).ToList();
		if (paragraphs.Count != items.Count) return;    // not one paragraph per <li>: leave it as it is

		// a legal list is one Word list: all its paragraphs take the numId of its first one, at the level of their depth
		var rootNum = new Dictionary<int, int>();
		for (int i = 0; i < paragraphs.Count; i++) {
			if (items[i].Root < 0) continue;
			var properties = paragraphs[i]!;
			if (!rootNum.TryGetValue(items[i].Root, out int numId)) rootNum[items[i].Root] = numId = properties.NumberingId!.Val!.Value;
			properties.NumberingId = new W.NumberingId { Val = numId };
			properties.NumberingLevelReference = new W.NumberingLevelReference { Val = Math.Min(8, items[i].Depth) };
		}

		var listOf = new Dictionary<int, (HtmlList List, int Level)>();
		var sizeOf = new Dictionary<int, float>();    // each list's largest item size
		for (int i = 0; i < paragraphs.Count; i++) {
			int numId = paragraphs[i]!.NumberingId!.Val!.Value;
			listOf.TryAdd(numId, (items[i], items[i].Root >= 0 ? 0 : paragraphs[i]!.NumberingLevelReference?.Val?.Value ?? 0));
			sizeOf[numId] = Math.Max(sizeOf.GetValueOrDefault(numId), sizes[i]);
		}
		// lists no paragraph uses any more (the nested parts of legal lists)
		foreach (var unused in numbering.Elements<W.NumberingInstance>().Where(n => n.NumberID?.Value is not int id || !listOf.ContainsKey(id)).ToList()) unused.Remove();
		var instances = numbering.Elements<W.NumberingInstance>().ToList();
		int nextId = numbering.Elements<W.AbstractNum>().Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
		var firstInstance = instances.FirstOrDefault();
		foreach (var instance in instances) {
			if (instance.NumberID?.Value is not int numId || !listOf.TryGetValue(numId, out var use)) continue;
			int id = nextId++;
			var definition = new W.AbstractNum { AbstractNumberId = id };
			definition.Append(new W.MultiLevelType { Val = W.MultiLevelValues.HybridMultilevel });
			bool legal = use.List.Format == "legal";
			for (int level = 0; level < 9; level++) {
				bool bullet = use.List.Format == "bullet";
				string number = "%" + (level + 1);
				string text = legal ? LegalText(level)
					: bullet ? "\u2022" : use.List.Punctuation switch { "paren" => number + ")", "parens" => "(" + number + ")", "plain" => number, _ => number + "." };
				var lvl = new W.Level { LevelIndex = level };
				var (first, left) = legal ? LegalLists.Indents(level) : OrdinaryIndents(use.List, level, sizeOf.GetValueOrDefault(numId, 11));
				lvl.Append(new W.StartNumberingValue { Val = level == use.Level ? use.List.Start : 1 },
					new W.NumberingFormat { Val = FormatValue(legal ? "decimal" : use.List.Format) },
					new W.LevelText { Val = text },
					new W.LevelJustification { Val = W.LevelJustificationValues.Left },
					new W.PreviousParagraphProperties(new W.Indentation { Left = ((int)(left * 20)).ToString(), Hanging = ((int)(-first * 20)).ToString() }));
				definition.Append(lvl);
			}
			numbering.InsertBefore(definition, firstInstance);
			instance.AbstractNumId = new W.AbstractNumId { Val = id };
			foreach (var start in instance.Descendants<W.StartOverrideNumberingValue>()) start.Val = use.List.Start;
		}
		// the definitions no list uses any more (HtmlToOpenXml's, with their clashing ids)
		var used = numbering.Elements<W.NumberingInstance>().Select(n => n.AbstractNumId?.Val?.Value).ToHashSet();
		foreach (var unused in numbering.Elements<W.AbstractNum>().Where(a => !used.Contains(a.AbstractNumberId?.Value)).ToList()) unused.Remove();
	}

	// An ordinary list level's indents (points): the number where the editor draws it (18 points a level, from 18), the
	// text after the widest number the level is likely to show, in the list's font size, and a gap.
	static (float First, float Left) OrdinaryIndents(HtmlList list, int level, float size) {
		string sample = list.Format switch { "bullet" => "\u2022", "lowerLetter" => "m", "upperLetter" => "M", "lowerRoman" => "viii", "upperRoman" => "VIII", _ => "88" };
		sample = list.Format == "bullet" ? sample : list.Punctuation switch { "paren" => sample + ")", "parens" => "(" + sample + ")", "plain" => sample, _ => sample + "." };
		float hang = MathF.Ceiling(Math.Max(18, LegalLists.TextWidth(sample, DefaultFont?.Name ?? "Segoe UI", size) + size * 0.5f));
		return (-hang, 18 * (level + 1) + hang);
	}

	// Word's level text for legal numbering at a level: "%1.", "%1.%2", "%1.%2.%3" ... (periods as LegalLists.Format)
	static string LegalText(int level) {
		string text = string.Join(".", Enumerable.Range(1, level + 1).Select(n => "%" + n));
		return LegalLists.Format switch { LegalLists.Periods.All => text + ".", LegalLists.Periods.Top => level == 0 ? text + "." : text, _ => text };
	}

	// Word draws a list number in its paragraph mark's format, and HtmlToOpenXml gives marks none (the document
	// default, 11 pt, next to 18-point text): each list paragraph's mark takes its first run's formatting.
	static void MarksFromRuns(MainDocumentPart main) {
		foreach (var paragraph in main.Document.Body!.Descendants<W.Paragraph>()) {
			var properties = paragraph.ParagraphProperties;
			if (properties?.NumberingProperties == null) continue;
			var run = paragraph.Descendants<W.Run>().FirstOrDefault(r => r.RunProperties != null && r.Descendants<W.Text>().Any(t => t.Text.Trim() != ""));
			if (run?.RunProperties == null) continue;
			var mark = new W.ParagraphMarkRunProperties();
			foreach (var child in run.RunProperties.ChildElements.Where(c => c is not W.RunStyle && c is not W.Underline)) mark.Append(child.CloneNode(true));
			properties.ParagraphMarkRunProperties = mark;
		}
	}

	static W.NumberFormatValues FormatValue(string format) => format switch {
		"lowerLetter" => W.NumberFormatValues.LowerLetter, "upperLetter" => W.NumberFormatValues.UpperLetter,
		"lowerRoman" => W.NumberFormatValues.LowerRoman, "upperRoman" => W.NumberFormatValues.UpperRoman,
		"bullet" => W.NumberFormatValues.Bullet, _ => W.NumberFormatValues.Decimal
	};

	static string OpenXmlRead(string input) {
		TypedIndents = null;
		LoadSkia();
		// inline styles: the editor's HTML import doesn't read CSS classes
		string title = Path.GetFileNameWithoutExtension(input);
		var d = new WmlToHtmlConverterSettings(title);
		var settings = new WmlToHtmlConverterSettings(title, d.ImageHandler, d.TextHandler, d.SymbolHandler, new PageBreakHandler(d.BreakHandler), d.FontHandler,
			false, "span { white-space: pre-wrap; }", "", d.CssClassPrefix);
		XElement html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument(input), settings);
		DocxLineSpacing.IntoHtml(input, html);    // (before the lists are rebuilt: one block per paragraph, in order)
		NoNamesAsAltTexts(html);
		ListsFromMarkers(html);
		TabsFromMargins(html);
		TypedIndents = TypedNumberIndents(html);
		return html.ToString(SaveOptions.DisableFormatting);
	}

	/// Set by the last read with the built-in reader: where each paragraph that starts with a typed list number has it
	/// (left margin plus first-line indent, in points), in order. The editor's HTML import drops paragraph margins, and
	/// LegalLists.ConvertTypedLists takes the list levels from these.
	public static List<float>? TypedIndents;

	static List<float> TypedNumberIndents(XElement html) {
		XNamespace x = html.Name.Namespace;
		var indents = new List<float>();
		foreach (var p in html.Descendants(x + "p")) {
			if (!LegalLists.TypedMarker.IsMatch(p.Value)) continue;
			string style = (string?)p.Attribute("style") ?? "";
			indents.Add(Points(style, "margin-left") + Points(style, "text-indent"));
		}
		return indents;
	}

	static float Points(string style, string property) {
		var m = Regex.Match(style, @"(?:^|;)\s*" + property + @":\s*(-?[\d.]+)(in|pt|cm|mm|px)?");
		if (!m.Success || !float.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float value)) return 0;
		return value * (m.Groups[2].Value switch { "in" => 72f, "cm" => 72 / 2.54f, "mm" => 72 / 25.4f, "px" => 0.75f, _ => 1f });
	}

	// The PowerTools write a tab as a right margin on the text before it (<span style="margin-right: 0.25in">1.</span>),
	// which the editor would drop: the tab character goes back in.
	static void TabsFromMargins(XElement html) {
		XNamespace x = html.Name.Namespace;
		foreach (var span in html.Descendants(x + "span").ToList()) {
			string style = (string?)span.Attribute("style") ?? "";
			// (and the tabs after the first as a left margin on a space: <span style="margin: 0 0 0 0.50in">&nbsp;</span>)
			if (span.Value == " " && !span.HasElements && Regex.IsMatch(style, @"^\s*margin:\s*0 0 0 [\d.]+(in|pt|cm|mm|px)?;")) {
				span.Value = "\t";
				span.SetAttributeValue("style", null);
				continue;
			}
			var margin = Regex.Match(style, @"margin-right:\s*([\d.]+)(in|pt|cm|mm|px)?");
			if (!margin.Success || !double.TryParse(margin.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double width) || width <= 0) continue;
			if (span.ElementsAfterSelf().FirstOrDefault() == null && span.NodesAfterSelf().FirstOrDefault() == null) continue;    // nothing after it
			span.Add("\t");
			span.SetAttributeValue("style", Regex.Replace(style, @"margin-right:\s*[^;]*;?", ""));
		}
	}

	// The PowerTools write each list number as text (<span data-pt-list-item-run="1.1">a)</span>) in an <ol> without
	// a list style. The first number of each list gives the list its style, as the editor's HTML has it (type="a", and
	// class="rte-lower-alpha-paren" for "a)", read by rte_htmllists.cs); bullets make it a <ul>. The numbers go.
	static void ListsFromMarkers(XElement html) {
		XNamespace x = html.Name.Namespace;
		static bool IsMarker(XElement span) => span.Attribute("data-pt-list-item-run") != null;
		// legal lists: those with compound numbers ("1.2"), and the lists they are in
		XElement? MarkerOf(XElement list) => list.Elements(x + "li").Descendants(x + "span").FirstOrDefault(IsMarker);
		var legal = new HashSet<XElement>();
		foreach (var list in html.Descendants(x + "ol").Where(l => MarkerOf(l) is { } m && Regex.IsMatch(m.Value.Trim(), @"^\d+(\.\d+)+\.?$")).ToList())
			for (var e = list; e != null && e.Name == x + "ol"; e = e.Parent)
				legal.Add(e);
		// a top-level numbered list that starts at the number after a legal list's last one carries that numbering on
		// (a clause after body text, with no sub-clauses of its own to show compound numbers)
		int? lastTop = null;
		foreach (var top in html.Descendants(x + "ol").Where(l => !l.Ancestors(x + "ol").Any())) {
			var tops = top.Elements(x + "li").Select(li => li.Descendants(x + "span").FirstOrDefault(IsMarker)?.Value.Trim() ?? "")
				.Select(t => Regex.Match(t, @"^(\d+)\.?$")).ToList();
			if (!legal.Contains(top) && lastTop is int last && tops.Count > 0 && tops[0].Success && int.Parse(tops[0].Groups[1].Value) == last + 1)
				foreach (var e in top.DescendantsAndSelf(x + "ol")) legal.Add(e);
			if (legal.Contains(top)) lastTop = tops.LastOrDefault(t => t.Success) is { } m ? int.Parse(m.Groups[1].Value) : lastTop;
		}
		foreach (var list in html.Descendants(x + "ol").ToList()) {
			var marker = MarkerOf(list);
			if (marker == null) continue;
			list.Attribute("class")?.Remove();
			if (legal.Contains(list)) {
				list.Attribute("type")?.Remove();
				list.SetAttributeValue("class", "rte-legal");
				continue;
			}
			var m = Regex.Match(marker.Value.Trim(), @"^(\()?([0-9]+|[a-zA-Z]+)([.)])?$");
			if (!m.Success) {    // a bullet
				list.Name = x + "ul";
				list.Attribute("start")?.Remove();
				continue;
			}
			string number = m.Groups[2].Value;
			bool roman = Regex.IsMatch(number, "^[ivxlcdm]+$", RegexOptions.IgnoreCase) && (number.Length > 1 || number is "i" or "I");
			string? type = char.IsDigit(number[0]) ? null : roman ? (char.IsLower(number[0]) ? "i" : "I") : (char.IsLower(number[0]) ? "a" : "A");
			if (type != null) list.SetAttributeValue("type", type);
			string style = type switch { null => "decimal", "a" => "lower-alpha", "A" => "upper-alpha", "i" => "lower-roman", _ => "upper-roman" };
			string? punctuation = m.Groups[1].Success ? "parens" : m.Groups[3].Value == ")" ? "paren" : m.Groups[3].Value == "" ? "plain" : null;
			if (punctuation != null) list.SetAttributeValue("class", $"rte-{style}-{punctuation}");
		}
		foreach (var span in html.Descendants(x + "span").Where(IsMarker).ToList()) span.Remove();
		// image sizes are in a style ("width: 0.61in"); the editor reads width and height attributes (pixels)
		foreach (var image in html.Descendants(x + "img")) {
			string style = (string?)image.Attribute("style") ?? "";
			foreach (var side in new[] { "width", "height" }) {
				var m = Regex.Match(style, side + @"\s*:\s*([\d.]+)\s*(in|pt|cm|mm|px)");
				if (image.Attribute(side) != null || !m.Success || !double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value)) continue;
				double pixels = value * m.Groups[2].Value switch { "in" => 96, "pt" => 96 / 72.0, "cm" => 96 / 2.54, "mm" => 96 / 25.4, _ => 1 };
				image.SetAttributeValue(side, (int)Math.Round(pixels));
			}
		}
	}

	// The PowerTools draw images with SkiaSharp, whose native libSkiaSharp.dll is in the NuGet cache, where Windows
	// doesn't look for it; once it is loaded, SkiaSharp finds it by name.
	static bool skiaLoaded;
	static void LoadSkia() {
		if (skiaLoaded) return;
		skiaLoaded = true;
		string root = Environment.GetEnvironmentVariable("NUGET_PACKAGES") ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nuget", "packages");
		string version = typeof(SkiaSharp.SKImage).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0] ?? "4.152.0";
		foreach (var v in new[] { version, "4.152.0" }) {
			string dll = Path.Combine(root, "skiasharp.nativeassets.win32", v, "runtimes", Environment.Is64BitProcess ? "win-x64" : "win-x86", "native", "libSkiaSharp.dll");
			if (File.Exists(dll)) { try { NativeLibrary.Load(dll); return; } catch { } }
		}
	}

	// ---- pandoc

	static string PandocConvert(string input, string output, string from, string to, string resources, params string[] extra) {
		string pandoc = PandocPath ?? throw new InvalidOperationException("not found");
		var start = new ProcessStartInfo(pandoc) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true };
		foreach (var arg in new[] { "-f", from, "-t", to, "--resource-path", resources, "-o", output }.Concat(extra).Append(input))
			start.ArgumentList.Add(arg);
		using var process = Process.Start(start)!;
		var error = process.StandardError.ReadToEndAsync();
		process.StandardOutput.ReadToEnd();
		if (!process.WaitForExit(120_000)) { try { process.Kill(true); } catch { } throw new TimeoutException("took more than 2 minutes"); }
		if (process.ExitCode != 0 || !File.Exists(output)) throw new InvalidOperationException(error.Result.Trim() is { Length: > 0 } e ? e : "exit code " + process.ExitCode);
		return output;
	}

	static string WriteRtf(string folder, string name, string rtf) {
		string path = Path.Combine(folder, name);
		File.WriteAllText(path, rtf, System.Text.Encoding.Latin1);
		return path;
	}

	// RichEdit writes numbered lists with a list table (\ls), but bullets the old way ({\*\pn\pnlvlblt ...}), and
	// LibreOffice makes those bullets part of the numbered list before them ("2. bullet"). This gives the bullets a list of
	// their own in the list table.
	public static string ModernBullets(string rtf) {
		var bullet = new Regex(@"\{\\pntext((?:[^{}])*)\}\s*\{\\\*\\pn\\pnlvlblt(?:[^{}]|\{[^{}]*\})*\}");
		if (!bullet.IsMatch(rtf)) return rtf;
		int listId = Regex.Matches(rtf, @"\\listid(\d+)").Select(m => int.Parse(m.Groups[1].Value)).DefaultIfEmpty(0).Max() + 1;
		int ls = Regex.Matches(rtf, @"\\ls(\d+)").Select(m => int.Parse(m.Groups[1].Value)).DefaultIfEmpty(0).Max() + 1;
		string list = $@"{{\list\listtemplateid{listId}\listsimple{{\listlevel\levelnfc23\leveljc0\levelstartat1{{\leveltext\'01\u8226 ?;}}{{\levelnumbers;}}\fi-360\li720}}\listid{listId}}}";
		string over = $@"{{\listoverride\listid{listId}\listoverridecount0\ls{ls}}}";
		string body = bullet.Replace(rtf, m => @"{\listtext" + m.Groups[1].Value + @"}\ls" + ls + " ");
		int tableAt = body.IndexOf(@"{\*\listtable", StringComparison.Ordinal), overAt = body.IndexOf(@"{\*\listoverridetable", StringComparison.Ordinal);
		if (tableAt >= 0 && overAt >= 0) {
			// after the control word and the space that ends it (else LibreOffice takes the space for text)
			int After(int at, string word) => at + word.Length + (at + word.Length < body.Length && body[at + word.Length] == ' ' ? 1 : 0);
			body = body.Insert(After(overAt, @"{\*\listoverridetable"), over);
			return body.Insert(After(tableAt, @"{\*\listtable"), list);
		}
		// no list table yet: both go before the document text (the first \pard)
		int text = body.IndexOf(@"\pard", StringComparison.Ordinal);
		return text < 0 ? rtf : body.Insert(text, @"{\*\listtable" + list + @"}{\*\listoverridetable" + over + "}\r\n");
	}

	// ---- legal lists for Word and LibreOffice

	// A marker in place of a legal number: plain ASCII, so that RichEdit writes it as it is (characters the font lacks
	// would get font changes around them).
	const string Mark = "@@RTE-LEGAL:";

	/// The RTF with legal lists (rte_legal.cs) as RTF list numbering. RichEdit writes a legal item as an old-style
	/// continuation paragraph ({\*\pn\pnlvlcont}) with its number as text, which Word and LibreOffice keep as text. Here,
	/// in a copy of the document, each item's number is replaced by a marker with its level and series; in that copy's
	/// RTF the markers become {\listtext number\tab}\lsN\ilvlK (the text for readers without lists), and the list table
	/// gets a legal list with one override per series (restarts start at their number).
	public static string LegalRtf(string rtf) {
		if (!rtf.Contains(@"\pnlvlcont", StringComparison.Ordinal)) return ListIndentsRtf(rtf);
		using var scratch = new EllipticBit.RichEditorNET.RichEditBox();
		_ = scratch.Handle;
		((System.Windows.Forms.RichTextBox)scratch).Rtf = rtf;
		var doc = scratch.TextDocument;
		LegalLists.LevelsFromNumbers(doc);
		var paragraphs = LegalLists.Paragraphs(doc);
		var numbers = LegalLists.Numbers(paragraphs);
		if (numbers.Count == 0) return rtf;
		var series = numbers.Values.Select(n => n.Series).Distinct().ToList();
		var texts = new List<string>();    // each marker's number text, by marker index
		foreach (var p in Enumerable.Reverse(paragraphs)) {
			if (!numbers.TryGetValue(p.Start, out var number)) continue;
			texts.Insert(0, number.Number);
			doc.Range2(p.Start, p.Start + p.PrefixLength).Text = $"{Mark}{p.Level},{series.IndexOf(number.Series)}@@";
		}
		FitListIndents(doc);    // (the ordinary lists' indents too)
		string marked = ((System.Windows.Forms.RichTextBox)scratch).Rtf ?? rtf;

		int listId = Regex.Matches(marked, @"\\listid(\d+)").Select(m => int.Parse(m.Groups[1].Value)).DefaultIfEmpty(0).Max() + 1;
		int firstLs = Regex.Matches(marked, @"\\ls(\d+)").Select(m => int.Parse(m.Groups[1].Value)).DefaultIfEmpty(0).Max() + 1;
		int index = 0;
		string body = Regex.Replace(marked, Regex.Escape(Mark) + @"(\d),(\d+)@@", m => {
			string text = index < texts.Count ? texts[index] : "";
			index++;
			return $@"{{\listtext {text}\tab}}\ls{firstLs + int.Parse(m.Groups[2].Value)}\ilvl{m.Groups[1].Value} ";
		});
		if (index != texts.Count) return rtf;    // the markers did not all come through: leave the RTF as it was
		body = Regex.Replace(body, @"\{\\\*\\pn\\pnlvlcont(?:[^{}]|\{[^{}]*\})*\}", "");    // RichEdit's continuation paragraphs

		// the legal list: level k is "%1.%2...%k" with LegalLists.Format's periods, indented as in the editor
		// one list per series, its top level starting at the series' number (LibreOffice ignores start overrides)
		var lists = new System.Text.StringBuilder();
		var overrides = new System.Text.StringBuilder();
		for (int k = 0; k < series.Count; k++) {
			lists.Append(LegalListRtf(listId + k, LegalLists.StartOf(series[k])));
			overrides.Append($@"{{\listoverride\listid{listId + k}\listoverridecount0\ls{firstLs + k}}}");
		}
		int tableAt = body.IndexOf(@"{\*\listtable", StringComparison.Ordinal), overAt = body.IndexOf(@"{\*\listoverridetable", StringComparison.Ordinal);
		if (tableAt >= 0 && overAt >= 0) {
			int After(int at, string word) => at + word.Length + (at + word.Length < body.Length && body[at + word.Length] == ' ' ? 1 : 0);
			body = body.Insert(After(overAt, @"{\*\listoverridetable"), overrides.ToString());
			return body.Insert(After(tableAt, @"{\*\listtable"), lists.ToString());
		}
		int header = body.IndexOf(@"\viewkind", StringComparison.Ordinal);
		if (header < 0) header = body.IndexOf(@"\pard", StringComparison.Ordinal);
		return header < 0 ? rtf : body.Insert(header, $@"{{\*\listtable{lists}}}{{\*\listoverridetable{overrides}}}" + "\r\n");
	}

	// An RTF list definition for legal numbering: level k is "%1.%2...%k" with LegalLists.Format's periods, indented as
	// in the editor; the top level starts at start.
	static string LegalListRtf(int listId, int start) {
		var list = new System.Text.StringBuilder($@"{{\list\listtemplateid{listId}");
		for (int level = 0; level < 9; level++) {
			var text = new System.Text.StringBuilder();
			var positions = new System.Text.StringBuilder();
			int length = 0;
			for (int k = 0; k <= level; k++) {
				if (k > 0) { text.Append('.'); length++; }
				positions.Append($@"\'{length + 1:x2}");
				text.Append($@"\'{k:x2}");
				length++;
			}
			bool period = LegalLists.Format == LegalLists.Periods.All || LegalLists.Format == LegalLists.Periods.Top && level == 0;
			if (period) { text.Append('.'); length++; }
			var (first, left) = LegalLists.Indents(level);
			list.Append($@"{{\listlevel\levelnfc0\leveljc0\levelfollow0\levelstartat{(level == 0 ? start : 1)}{{\leveltext\'{length:x2}{text};}}{{\levelnumbers{positions};}}\fi{(int)(first * 20)}\li{(int)(left * 20)}\jclisttab\tx{(int)(left * 20)}}}");
		}
		return list.Append($@"\listid{listId}}}").ToString();
	}

	// ---- pictures in RTF

	/// The RTF with Word's picture wrappers taken off: {\*\shppict {\pict ...}} becomes the {\pict ...} in it, and the
	/// {\nonshppict ...} copy for older readers goes. RichEdit drops a {\*\shppict} inside a link (Word's RTF of a linked
	/// picture), and would show both copies of a picture where it reads neither wrapper.
	public static string SimplePictures(string rtf) {
		if (!rtf.Contains(@"\shppict", StringComparison.Ordinal)) return rtf;
		var sb = new System.Text.StringBuilder(rtf.Length);
		int at = 0;
		foreach (Match m in Regex.Matches(rtf, @"\{\\\*\\shppict\b|\{\\nonshppict\b")) {
			if (m.Index < at) continue;    // inside a wrapper already handled
			int end = GroupEnd(rtf, m.Index);
			if (end < 0) continue;
			sb.Append(rtf, at, m.Index - at);
			if (m.Value.Contains("nonshppict")) { at = end; continue; }
			int pict = rtf.IndexOf(@"{\pict", m.Index, end - m.Index, StringComparison.Ordinal);
			if (pict >= 0 && GroupEnd(rtf, pict) is int pictEnd && pictEnd > 0 && pictEnd <= end) sb.Append(rtf, pict, pictEnd - pict);
			at = end;
		}
		sb.Append(rtf, at, rtf.Length - at);
		return sb.ToString();
	}

	/// A picture in RTF: its image (PNG, JPEG or GIF), and the size it is shown at, in pixels.
	public sealed record RtfPicture(byte[] Image, int Width, int Height, string Alt = "");

	static readonly Regex Description = new(@"\{\\sp\s*\{\\sn\s+wzDescription\s*\}\s*\{\\sv\s?((?:[^{}\\]|\\.|\\[a-z]+-?\d*\s?)*)\}\s*\}");

	// A picture group's alt text (Word's and RichEdit-with-RTE-Writer's description property), or "".
	static string AltOf(string group) {
		var m = Description.Match(group);
		if (!m.Success) return "";
		var sb = new System.Text.StringBuilder();
		string v = m.Groups[1].Value;
		for (int i = 0; i < v.Length; i++) {
			if (v[i] != '\\') { if (v[i] is not ('\r' or '\n')) sb.Append(v[i]); continue; }
			if (i + 1 >= v.Length) break;
			char next = v[i + 1];
			if (next == '\'' && i + 3 < v.Length) {
				sb.Append(System.Text.Encoding.Latin1.GetString(new[] { Convert.ToByte(v.Substring(i + 2, 2), 16) }));
				i += 3;
			} else if (next == 'u' && Regex.Match(v[(i + 2)..], @"^(-?\d+) ?(\?|\\'[0-9a-fA-F]{2}|[^\\{}])?") is { Success: true } u) {
				sb.Append((char)(short)int.Parse(u.Groups[1].Value));    // (and the character for readers without Unicode goes)
				i += 1 + u.Length;
			} else if (char.IsLetter(next)) {
				var word = Regex.Match(v[(i + 1)..], @"^[a-z]+-?\d* ?");
				i += word.Length;
			} else {
				sb.Append(next);
				i++;
			}
		}
		return sb.ToString().Trim();
	}

	static string RtfEscape(string text) {
		var sb = new System.Text.StringBuilder();
		foreach (char c in text) {
			if (c is '\\' or '{' or '}') sb.Append('\\').Append(c);
			else if (c > 126 || c < 32) sb.Append(@"\u").Append((short)c).Append('?');
			else sb.Append(c);
		}
		return sb.ToString();
	}

	static string PictureProperty(string alt) => alt == "" ? "" : @"{\*\picprop{\sp{\sn wzDescription}{\sv " + RtfEscape(alt) + "}}}";

	// The PowerTools give a picture without a description its name as alt text ("Image1", "Picture 3"): that is not one.
	static void NoNamesAsAltTexts(XElement html) {
		foreach (var img in html.Descendants().Where(e => e.Name.LocalName == "img"))
			if ((string?)img.Attribute("alt") is { } alt && Regex.IsMatch(alt.Trim(), @"^(Image|Picture|Graphic|Graphics|Grafik|Bild|Imagen|Immagine)\s*\d*$", RegexOptions.IgnoreCase))
				img.SetAttributeValue("alt", null);
	}

	static readonly System.Xml.Linq.XNamespace WordDrawing = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		OdtDraw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", OdtSvg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

	/// Gives the pictures of a .docx or .odt their alt texts (in document order) where the converter left them out
	/// (LibreOffice ignores RTF's, the built-in converter wrote empty ones). Left as it is if the number of pictures differs.
	public static void AltTextsToFile(string path, IList<string> alts) {
		if (alts.All(a => a == "")) return;
		bool docx = path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase);
		string entryName = docx ? "word/document.xml" : "content.xml";
		using var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update);
		var entry = zip.GetEntry(entryName);
		if (entry == null) return;
		XDocument xml;
		using (var stream = entry.Open()) xml = XDocument.Load(stream);
		// the pictures: Word's drawing properties, or LibreOffice's frames that hold an image
		var pictures = docx
			? xml.Descendants(WordDrawing + "docPr").ToList()
			: xml.Descendants(OdtDraw + "frame").Where(f => f.Element(OdtDraw + "image") != null).ToList();
		if (pictures.Count != alts.Count) return;
		bool changed = false;
		for (int k = 0; k < pictures.Count; k++) {
			if (alts[k] == "") continue;
			if (docx) {
				if (!string.IsNullOrEmpty((string?)pictures[k].Attribute("descr"))) continue;
				pictures[k].SetAttributeValue("descr", alts[k]);
			} else {
				if (pictures[k].Element(OdtSvg + "desc") != null) continue;
				pictures[k].Add(new XElement(OdtSvg + "desc", alts[k]));    // (after draw:image, as LibreOffice writes it)
			}
			changed = true;
		}
		if (!changed) return;
		entry.Delete();
		var added = zip.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Optimal);
		using (var stream = added.Open()) xml.Save(stream, SaveOptions.DisableFormatting);
	}

	/// The RTF without its pictures' descriptions (alt texts): RichEdit puts their non-ASCII characters into the text
	/// ("café" left an "é" before the picture). Read them with Pictures first.
	public static string WithoutAltTexts(string rtf) => rtf.Contains("wzDescription", StringComparison.Ordinal) ? Description.Replace(rtf, "") : rtf;

	/// The RTF with alt texts for its pictures, in document order ("" or a missing one: none), written as each picture's
	/// description property. Left as it is if the number of pictures differs.
	public static string WithAltTexts(string rtf, IList<string> alts) {
		var groups = PictureGroups(rtf);
		if (groups.Count == 0 || groups.Count != alts.Count || alts.All(a => a == "")) return rtf;
		var sb = new System.Text.StringBuilder(rtf.Length + 256);
		int at = 0;
		for (int k = 0; k < groups.Count; k++) {
			var (start, end) = groups[k];
			if (alts[k] == "" || Description.IsMatch(rtf[start..end])) continue;
			sb.Append(rtf, at, start + @"{\pict".Length - at);
			sb.Append(PictureProperty(alts[k]));
			at = start + @"{\pict".Length;
		}
		sb.Append(rtf, at, rtf.Length - at);
		return sb.ToString();
	}

	// Each picture group in RTF, in document order (not the {\nonshppict} copies, which repeat the one before them).
	static List<(int Start, int End)> PictureGroups(string rtf) {
		var groups = new List<(int, int)>();
		var skip = new List<(int, int)>();
		foreach (Match m in Regex.Matches(rtf, @"\{\\nonshppict\b")) { int end = GroupEnd(rtf, m.Index); if (end > 0) skip.Add((m.Index, end)); }
		foreach (Match m in Regex.Matches(rtf, @"\{\\pict\b")) {
			if (skip.Any(s => m.Index > s.Item1 && m.Index < s.Item2)) continue;
			int end = GroupEnd(rtf, m.Index);
			if (end > 0) groups.Add((m.Index, end));
		}
		return groups;
	}

	/// The pictures in RTF, in document order (null for one that can't be read).
	public static List<RtfPicture?> Pictures(string rtf) => PictureGroups(rtf).Select(g => ReadPicture(rtf[g.Start..g.End])).ToList();

	/// The RTF with its metafile pictures (WMF, EMF: RichEdit keeps every picture as a WMF) as PNG, which Word and
	/// LibreOffice keep as it is; LibreOffice kept a WMF, which neither it nor Word could read back.
	public static string PngPictures(string rtf) {
		var groups = PictureGroups(rtf);
		if (groups.Count == 0) return rtf;
		var sb = new System.Text.StringBuilder(rtf.Length);
		int at = 0;
		foreach (var (start, end) in groups) {
			string group = rtf[start..end];
			sb.Append(rtf, at, start - at);
			at = end;
			if (Regex.IsMatch(group, @"\\(pngblip|jpegblip)\b") || ReadPicture(group) is not { } picture) { sb.Append(group); continue; }
			string png = Convert.ToHexString(picture.Image);
			string blip = picture.Image.Length > 2 && picture.Image[0] == 0xFF && picture.Image[1] == 0xD8 ? "jpegblip" : "pngblip";
			var goal = (W: Regex.Match(group, @"\\picwgoal(\d+)"), H: Regex.Match(group, @"\\pichgoal(\d+)"));
			int pixelsW, pixelsH;
			using (var ms = new MemoryStream(picture.Image)) using (var image = System.Drawing.Image.FromStream(ms)) { pixelsW = image.Width; pixelsH = image.Height; }
			sb.Append("{\\pict").Append(PictureProperty(picture.Alt)).Append($@"\{blip}\picw{pixelsW}\pich{pixelsH}\picwgoal{(goal.W.Success ? goal.W.Groups[1].Value : (picture.Width * 15).ToString())}\pichgoal{(goal.H.Success ? goal.H.Groups[1].Value : (picture.Height * 15).ToString())} ");
			for (int i = 0; i < png.Length; i += 128) sb.Append(png, i, Math.Min(128, png.Length - i)).Append("\r\n");
			sb.Append('}');
		}
		sb.Append(rtf, at, rtf.Length - at);
		return sb.ToString();
	}

	// One {\pict ...} group: its image as PNG (or the JPEG/PNG it holds) and its shown size.
	static RtfPicture? ReadPicture(string group) {
		try {
			// without the groups in it ({\*\picprop {\sp {\sn ...}{\sv ...}}}, {\*lipuid ...}), innermost first
			string words = group[1..^1], previous;
			do { previous = words; words = Regex.Replace(words, @"\{[^{}]*\}", " "); } while (words != previous);    // (a space: "liptag-1{...}89504e" isn't one number)
			int Number(string word, int fallback) => Regex.Match(words, @"\\" + word + @"(-?\d+)") is { Success: true } m ? int.Parse(m.Groups[1].Value) : fallback;
			string hex = Regex.Replace(Regex.Replace(words, @"\\[a-z]+-?\d* ?", ""), @"[^0-9a-fA-F]", "");
			if (hex.Length < 8) return null;
			byte[] data = Convert.FromHexString(hex.Length % 2 == 0 ? hex : hex[..^1]);
			int goalW = Number("picwgoal", 0), goalH = Number("pichgoal", 0);
			int scaleX = Number("picscalex", 100), scaleY = Number("picscaley", 100);
			int width = Math.Max(1, (int)Math.Round(goalW * scaleX / 100.0 / 15)), height = Math.Max(1, (int)Math.Round(goalH * scaleY / 100.0 / 15));
			byte[]? image = null;
			if (Regex.IsMatch(words, @"\\(pngblip|jpegblip)\b")) image = data;
			else if (Regex.IsMatch(words, @"\\wmetafile\d*")) image = BitmapOfWmf(data) ?? RenderMetafile(data, wmf: true, width, height);
			else if (Regex.IsMatch(words, @"\\emfblip\b")) image = RenderMetafile(data, wmf: false, width, height);
			else if (Regex.IsMatch(words, @"\\dibitmap\d*")) image = PngOfDib(data, 0);
			if (image == null) return null;
			using (var check = new MemoryStream(image)) using (System.Drawing.Image.FromStream(check)) { }    // a picture that can be read, or none
			if (goalW == 0 || goalH == 0) {
				using var ms = new MemoryStream(image);
				using var picture = System.Drawing.Image.FromStream(ms);
				(width, height) = (picture.Width, picture.Height);
			}
			return new RtfPicture(image, width, height, AltOf(group));
		} catch { return null; }
	}

	// RichEdit's WMF of a picture is a single StretchDIB (or DIBStretchBlt) record holding the original bitmap: the
	// bitmap itself, as PNG (null if the WMF is anything else).
	static byte[]? BitmapOfWmf(byte[] wmf) {
		for (int at = 18; at + 6 <= wmf.Length;) {    // records after the 18-byte header: size (16-bit words), function
			int size = BitConverter.ToInt32(wmf, at) * 2, function = BitConverter.ToUInt16(wmf, at + 4);
			if (size < 6 || at + size > wmf.Length) break;
			if (function == 0x0F43) return PngOfDib(wmf[(at + 28)..(at + size)], 0);    // META_STRETCHDIB
			if (function == 0x0B41) return PngOfDib(wmf[(at + 26)..(at + size)], 0);    // META_DIBSTRETCHBLT
			at += size;
		}
		return null;
	}

	// A device-independent bitmap (BITMAPINFOHEADER, colours, bits) as PNG.
	static byte[]? PngOfDib(byte[] dib, int offset) {
		int headerSize = BitConverter.ToInt32(dib, offset);
		int bitCount = BitConverter.ToUInt16(dib, offset + 14), compression = BitConverter.ToInt32(dib, offset + 16);
		int colorsUsed = headerSize >= 40 ? BitConverter.ToInt32(dib, offset + 32) : 0;
		int palette = (colorsUsed > 0 ? colorsUsed : bitCount <= 8 ? 1 << bitCount : 0) * 4 + (compression == 3 && headerSize == 40 ? 12 : 0);
		var bmp = new byte[14 + dib.Length - offset];
		bmp[0] = (byte)'B'; bmp[1] = (byte)'M';
		BitConverter.GetBytes(bmp.Length).CopyTo(bmp, 2);
		BitConverter.GetBytes(14 + headerSize + palette).CopyTo(bmp, 10);
		Array.Copy(dib, offset, bmp, 14, dib.Length - offset);
		using var input = new MemoryStream(bmp);
		using var bitmap = new System.Drawing.Bitmap(input);
		using var output = new MemoryStream();
		bitmap.Save(output, System.Drawing.Imaging.ImageFormat.Png);
		byte[] png = output.ToArray();
		// a photo (many colours, no transparency) is much smaller as a high-quality JPEG
		if (bitCount < 24 || png.Length < 100_000) return png;
		var codec = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
		using var parameters = new System.Drawing.Imaging.EncoderParameters(1);
		parameters.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 92L);
		using var jpeg = new MemoryStream();
		bitmap.Save(jpeg, codec, parameters);
		return jpeg.Length * 2 < png.Length ? jpeg.ToArray() : png;
	}

	// Draws a metafile (WMF or EMF) at a size, as PNG.
	static byte[]? RenderMetafile(byte[] data, bool wmf, int width, int height) {
		System.Drawing.Imaging.Metafile? metafile = null;
		try {
			if (wmf) {
				var pict = new MetafilePict { mm = 8, xExt = width * 2540 / 96, yExt = height * 2540 / 96 };    // MM_ANISOTROPIC, in 0.01 mm
				IntPtr emf = SetWinMetaFileBits((uint)data.Length, data, IntPtr.Zero, ref pict);
				if (emf == IntPtr.Zero) return null;
				metafile = new System.Drawing.Imaging.Metafile(emf, true);
			} else {
				metafile = new System.Drawing.Imaging.Metafile(new MemoryStream(data));
			}
			using var bitmap = new System.Drawing.Bitmap(Math.Max(1, width), Math.Max(1, height));
			using (var g = System.Drawing.Graphics.FromImage(bitmap)) {
				g.Clear(System.Drawing.Color.White);
				g.DrawImage(metafile, 0, 0, bitmap.Width, bitmap.Height);
			}
			using var output = new MemoryStream();
			bitmap.Save(output, System.Drawing.Imaging.ImageFormat.Png);
			return output.ToArray();
		} catch { return null; } finally { metafile?.Dispose(); }
	}

	[StructLayout(LayoutKind.Sequential)]
	struct MetafilePict { public int mm, xExt, yExt; public IntPtr hMF; }

	[DllImport("gdi32.dll")]
	static extern IntPtr SetWinMetaFileBits(uint size, byte[] data, IntPtr hdcRef, ref MetafilePict pict);

	// ---- list indents for Word and LibreOffice

	/// The RTF with ordinary lists' indents made hanging indents wide enough for their numbers. RichEdit draws list
	/// numbers itself, and its RTF gives every level an 18-point indent, too narrow for "(a)" at 18 points: in Word and
	/// LibreOffice the tab after the number then went to the next tab stop, leaving gaps of different widths.
	public static string ListIndentsRtf(string rtf) {
		if (!rtf.Contains(@"\ls", StringComparison.Ordinal) && !rtf.Contains(@"{\*\pn", StringComparison.Ordinal)) return rtf;
		using var scratch = new EllipticBit.RichEditorNET.RichEditBox();
		_ = scratch.Handle;
		((System.Windows.Forms.RichTextBox)scratch).Rtf = rtf;
		return FitListIndents(scratch.TextDocument) ? ((System.Windows.Forms.RichTextBox)scratch).Rtf ?? rtf : rtf;
	}

	// Ordinary (not legal) list paragraphs: the number where the editor draws it (18 points a level, from 18), the text
	// after the level's widest number and a gap. Returns true if any changed.
	static bool FitListIndents(EllipticBit.RichEditorNET.TextObjectModel2.ITextDocument2 doc) {
		var items = new List<(int Start, int Level, string Number, string Font, float Size)>();
		foreach (var p in LegalLists.Paragraphs(doc)) {
			if ((p.ListType & 0xFFFF) == 0 || LegalLists.IsLegal(p.ListType)) continue;
			var range = doc.Range2(p.Start, p.End);
			string shown = range.GetText2(EllipticBit.RichEditorNET.TextObjectModel2.tomConstants.tomIncludeNumbering) ?? "";
			int tab = shown.IndexOf('\t');
			string number = tab > 0 ? shown[..tab] : "8.";
			var mark = doc.Range2(Math.Max(p.Start, p.End - 1), Math.Max(p.Start, p.End - 1) + 1).Font;    // numbers are drawn in the mark's format
			items.Add((p.Start, Math.Min(p.Level, 8), number, mark.Name ?? "Segoe UI", mark.Size > 0 ? mark.Size : 11));
		}
		if (items.Count == 0) return false;
		var hang = new float[9];
		foreach (var item in items) hang[item.Level] = Math.Max(hang[item.Level], LegalLists.TextWidth(item.Number, item.Font, item.Size) + item.Size * 0.5f);
		bool changed = false;
		foreach (var item in items) {
			float h = MathF.Ceiling(Math.Max(18, hang[item.Level]));
			var para = doc.Range2(item.Start, item.Start).Para;
			float first = -h, left = 18 * (item.Level + 1) + h;
			if (Math.Abs(para.FirstLineIndent - first) < 0.5f && Math.Abs(para.LeftIndent - left) < 0.5f) continue;
			para.SetIndents(first, left, para.RightIndent);
			changed = true;
		}
		return changed;
	}

	// ---- the document's default font for Word and LibreOffice

	/// The RTF with a Normal style in the editor's font and size (DefaultFont). RichEdit's RTF has no styles, so what has
	/// no formatting of its own got the converter's default: in LibreOffice, list numbers came out in Liberation Serif
	/// 12 pt. Leaves the RTF as it is without a DefaultFont, or when it already has a style sheet.
	public static string WithNormalStyle(string rtf) {
		if (DefaultFont is not { } font || rtf.Contains(@"{\stylesheet", StringComparison.Ordinal)) return rtf;
		int table = rtf.IndexOf(@"{\fonttbl", StringComparison.Ordinal);
		if (table < 0) return rtf;
		int tableEnd = GroupEnd(rtf, table);
		if (tableEnd < 0) return rtf;
		// the editor's font in the font table
		var entry = Regex.Matches(rtf[table..tableEnd], @"\{\\f(\d+)(?:\\[a-z]+-?\d*\s?)*(?:\{[^{}]*\})?\s*([^{};\\]+);\}")
			.FirstOrDefault(m => m.Groups[2].Value.Trim().Equals(font.Name, StringComparison.OrdinalIgnoreCase));
		string fontWord = entry != null ? @"\f" + entry.Groups[1].Value : "";
		string style = $@"{{\stylesheet{{\s0\snext0{fontWord}\fs{(int)Math.Round(font.Points * 2)} Normal;}}}}";
		// after the colour table if there is one (RTF's order: fonts, colours, styles, lists)
		int colors = rtf.IndexOf(@"{\colortbl", tableEnd, StringComparison.Ordinal);
		int at = colors >= 0 && colors - tableEnd < 8 ? GroupEnd(rtf, colors) : tableEnd;
		return at < 0 ? rtf : rtf.Insert(at, style);
	}

	// The position after the group that starts at start ({ ... }), or -1.
	static int GroupEnd(string rtf, int start) {
		int depth = 0;
		for (int i = start; i < rtf.Length; i++) {
			if (rtf[i] == '\\') { i++; continue; }
			if (rtf[i] == '{') depth++;
			else if (rtf[i] == '}' && --depth == 0) return i + 1;
		}
		return -1;
	}

	static string WriteHtml(string folder, string html) {
		string path = Path.Combine(folder, "document.html");
		File.WriteAllText(path, html, new System.Text.UTF8Encoding(false));
		return path;
	}

	// ---- helpers

	static List<string> Order(string[] preferred, string? via) {
		via = string.IsNullOrWhiteSpace(via) ? Environment.GetEnvironmentVariable("RTE_WRITER_CONVERTERS") : via;
		if (string.IsNullOrWhiteSpace(via)) return preferred.ToList();
		var named = via.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(v => v.ToLowerInvariant()).ToList();
		var unknown = named.Where(n => !All.Contains(n)).ToList();
		if (unknown.Count > 0) throw new ArgumentException("Unknown converter " + string.Join(", ", unknown) + " (use word, libreoffice, openxml or pandoc)");
		return named.Distinct().ToList();
	}

	static string Name(string converter) => converter switch { Word => "Word", LibreOffice => "LibreOffice", OpenXml => "Open XML", _ => "pandoc" };
	static string Message(Exception ex) => (ex is TargetInvocationException { InnerException: { } inner } ? inner : ex).Message.Trim();

	static string TempFolder() {
		string path = Path.Combine(Path.GetTempPath(), "RTE Writer", Guid.NewGuid().ToString("N"));
		Directory.CreateDirectory(path);
		return path;
	}

	static void TryDelete(string folder) {
		try { Directory.Delete(folder, true); } catch { }
	}
}

// The PowerTools' breaks, with page breaks marked as such (their own handler writes every break as <br />).
sealed class PageBreakHandler : IBreakHandler
{
	static readonly System.Xml.Linq.XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	readonly IBreakHandler inner;

	public PageBreakHandler(IBreakHandler inner) => this.inner = inner;

	public IEnumerable<System.Xml.Linq.XNode> TransformBreak(System.Xml.Linq.XElement element) {
		var nodes = inner.TransformBreak(element).ToList();
		if ((string?)element.Attribute(W + "type") != "page") return nodes;
		foreach (var br in nodes.OfType<System.Xml.Linq.XElement>().SelectMany(n => n.DescendantsAndSelf()).Where(n => n.Name.LocalName == "br"))
			br.SetAttributeValue("style", "page-break-before: always");
		return nodes;
	}
}
