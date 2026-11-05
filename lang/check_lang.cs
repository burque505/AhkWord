// Checks the languages' files against the program's strings (multilingual, AhkWord-es). Run by check_lang.ahk.
// The program's strings are the T("key", "English") and TN("key", "one", "other", count) calls in its *.cs files (Tr in
// rte_strings.cs), and the ribbon's (ribbon\RteRibbon.strings.json, "cmdPaste.LabelTitle"). For each lang\<code>.json:
// what it lacks (shown in English), what it has that nothing uses any more, and translations whose {0} placeholders
// differ from the English's. In the code: a key used with two Englishes, and a T call this can't read (its key and
// English must be literals). Writes lang\en.json, every key with its English: the table a new language starts from, and,
// committed, a record whose changes (git diff) show which translations to look at again; what changed since the last
// run is reported too.
// Iteration 80: an extension's strings (a build with more features: rte_extensions.cs), in a folder of the program's that
// has a lang folder of its own (as Lang.ExtensionFolders finds them: pro\lang). Its code is that folder's *.cs (and its
// subfolders'), its translations <folder>\lang\<code>.json (read over AhkWord's), its English <folder>\lang\en.json. A
// key the extension's code uses goes in its own file, AhkWord's in lang\<code>.json; the ribbon's in either.
// args: [0] the program's folder, [1] a language code (optional: every lang\*.json)
// Returns the report; exit code 1 when there are errors (not for strings a language lacks).
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;

string root = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
string langDir = Path.Combine(root, "lang");
var report = new StringBuilder();
int errors = 0;
void Error(string text) { report.AppendLine("  ERROR " + text); errors++; }

// --- the program's strings
var code = new Dictionary<string, Entry>(StringComparer.Ordinal);    // key -> its English (TN: key.one, key.other)
var extensions = Directory.GetDirectories(root).Where(d => !Path.GetFileName(d).Equals("lang", StringComparison.OrdinalIgnoreCase) && Directory.Exists(Path.Combine(d, "lang")))
	.Select(d => Path.GetFileName(d)).OrderBy(d => d, StringComparer.OrdinalIgnoreCase).ToList();    // (as Lang.ExtensionFolders)
var files = Directory.GetFiles(root, "*.cs").OrderBy(f => Path.GetFileName(f) == "AhkWord.cs" ? "" : Path.GetFileName(f), StringComparer.OrdinalIgnoreCase)
	.Select(f => (File: f, Origin: ""))
	.Concat(extensions.SelectMany(x => Directory.GetFiles(Path.Combine(root, x), "*.cs", SearchOption.AllDirectories).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).Select(f => (File: f, Origin: x))));
int calls = 0;
foreach (var (file, origin) in files) {
	string text = File.ReadAllText(file);
	string name = Path.GetRelativePath(root, file);
	var toks = new List<Tok>();
	int i = 0;
	try { Lex(text, ref i, false, toks); }
	catch (Exception ex) { Error($"{name}: could not be read ({ex.Message}, line {Line(text, i)})."); continue; }
	for (int t = 0; t < toks.Count; t++) {
		if (toks[t].Kind != K.Ident || toks[t].Text is not ("T" or "TN")) continue;
		if (t + 1 >= toks.Count || !toks[t + 1].Is('(')) continue;
		if (t > 0 && toks[t - 1].Is('.') && !(t > 1 && toks[t - 2].Text == "Tr")) continue;    // (x.T(...): another T)
		if (t > 0 && toks[t - 1].Kind == K.Ident && toks[t - 1].Text is not ("return" or "await" or "throw" or "case" or "in" or "yield" or "else" or "when")) continue;    // (a declaration)
		bool plural = toks[t].Text == "TN";
		string where = $"{name}:{Line(text, toks[t].Pos)}";
		int j = t + 2;
		var parts = new List<string>();
		for (int n = 0; n < (plural ? 3 : 2); n++) {
			if (n > 0) { if (j < toks.Count && toks[j].Is(',')) j++; else { parts = null!; break; } }
			if (ReadLiteral(toks, ref j) is not string s) { parts = null!; break; }
			parts.Add(s);
		}
		if (parts == null) { Error($"{where}: a {toks[t].Text} call whose key and English aren't literals: \"{Excerpt(text, toks[t].Pos)}\""); continue; }
		bool formatted = plural || (j < toks.Count && toks[j].Is(','));
		calls++;
		if (plural) { Add(parts[0] + ".one", parts[1], true, where, origin); Add(parts[0] + ".other", parts[2], true, where, origin); }
		else Add(parts[0], parts[1], formatted, where, origin);
	}
}
void Add(string key, string english, bool formatted, string where, string origin) {
	if (!Regex.IsMatch(key, @"^[A-Za-z][A-Za-z0-9_]*(\.[A-Za-z0-9_]+)+$")) { Error($"{where}: \"{key}\" is not a key (words with dots: msg.saveChanges)."); return; }
	if (key.StartsWith("cmd")) { Error($"{where}: \"{key}\": keys starting with cmd are the ribbon's."); return; }
	if (code.TryGetValue(key, out var had)) {
		if (had.English != english) Error($"{where}: \"{key}\" is used with another English at {had.Where}: \"{english}\" / \"{had.English}\".");
		else if (had.Formatted != formatted) Error($"{where}: \"{key}\" is formatted (arguments) here and not at {had.Where}, or the other way round.");
		else if (had.Origin != origin) Error($"{where}: \"{key}\" is used by {(had.Origin == "" ? "AhkWord" : "the extension " + had.Origin)} ({had.Where}) too: an extension's keys are its own.");
		return;
	}
	if (formatted && Placeholders(english) == null) Error($"{where}: \"{key}\": its English isn't a format string (a {{ or }} not doubled?): \"{english}\".");
	code[key] = new Entry(english, formatted, where, origin);
}

// --- the ribbon's strings
var ribbon = new Dictionary<string, string>(StringComparer.Ordinal);
string stringsFile = Path.Combine(root, "ribbon", "RteRibbon.strings.json");
using (var doc = JsonDocument.Parse(File.ReadAllText(stringsFile)))
	foreach (var command in doc.RootElement.EnumerateObject())
		foreach (var p in command.Value.EnumerateObject())
			if (p.Value.ValueKind == JsonValueKind.String) ribbon[command.Name + "." + p.Name] = p.Value.GetString()!;

report.Insert(0, $"Strings: {code.Count} of the program's ({calls} T and TN calls), {ribbon.Count} of the ribbon's.\n");

// --- en.json: every key and its English (an extension's in its own); what changed since it was written
var english = ribbon.Concat(code.Select(c => KeyValuePair.Create(c.Key, c.Value.English))).ToDictionary(e => e.Key, e => e.Value);
var changed = new Dictionary<string, string>();    // key -> the English before
foreach (var origin in extensions.Prepend("")) {
	string enFile = Path.Combine(LangDir(origin), "en.json");
	var mine = (origin == "" ? ribbon.ToList() : new List<KeyValuePair<string, string>>())
		.Concat(code.Where(c => c.Value.Origin == origin).OrderBy(c => c.Key, StringComparer.Ordinal).Select(c => KeyValuePair.Create(c.Key, c.Value.English))).ToList();
	if (File.Exists(enFile))
		try {
			foreach (var (key, was) in ReadTable(enFile))
				if (english.TryGetValue(key, out var now) && now != was) changed[key] = was;
		} catch (Exception ex) { report.AppendLine($"  ({Shown(enFile)} could not be read, so what changed isn't known: {ex.Message})"); }
	var json = new MemoryStream();
	using (var w = new Utf8JsonWriter(json, new JsonWriterOptions { Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping })) {
		w.WriteStartObject();
		foreach (var (key, value) in mine) w.WriteString(key, value);
		w.WriteEndObject();
	}
	string header = origin == ""
		? "// English: every string of AhkWord's interface by its key, the table a language's file translates (lang\\<code>.json).\n"
			+ "// Written by check_lang.ahk; don't edit: the English is in the code (T and TN calls) and in ribbon\\RteRibbon.xml.\n"
		: $"// English: every string of the extension {origin} by its key, the table a language's file translates ({origin}\\lang\\<code>.json).\n"
			+ "// Written by check_lang.ahk; don't edit: the English is in the code (T and TN calls).\n";
	File.WriteAllText(enFile, header + Encoding.UTF8.GetString(json.ToArray()) + "\n", new UTF8Encoding(false));
}

// --- each language: AhkWord's file and the extensions'
var codes = args.Length > 1 && args[1] != ""
	? new[] { args[1] }
	: extensions.Prepend("").SelectMany(o => Directory.Exists(LangDir(o)) ? Directory.GetFiles(LangDir(o), "*.json") : Array.Empty<string>())
		.Select(f => Path.GetFileNameWithoutExtension(f)).Where(c => c != "en").Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(c => c).ToArray();
foreach (var lang in codes) {
	report.AppendLine().AppendLine($"{lang}.json");
	var tables = new Dictionary<string, Dictionary<string, string>>();    // origin -> its file's table (an extension's only when it has one)
	bool unreadable = false;
	foreach (var origin in extensions.Prepend("")) {
		string file = Path.Combine(LangDir(origin), lang + ".json");
		if (!File.Exists(file)) { if (origin == "") tables[""] = new(); continue; }    // (AhkWord's: none, every string lacking)
		try { tables[origin] = ReadTable(file); }
		catch (Exception ex) { Error($"{Shown(file)} could not be read: {ex.Message}"); unreadable = true; }
	}
	if (unreadable) continue;
	foreach (var origin in tables.Keys.Where(o => o != "")) report.AppendLine($"  and {Shown(Path.Combine(LangDir(origin), lang + ".json"))}");
	bool Has(string key) => tables.Values.Any(t => t.ContainsKey(key));
	var missingCode = code.Where(c => !(tables.TryGetValue(c.Value.Origin, out var t) && t.ContainsKey(c.Key))).ToList();
	var missingRibbon = ribbon.Keys.Where(k => !k.EndsWith(".Keytip") && !Has(k)).ToList();    // (a keytip: may stay English)
	var unused = new List<string>();    // (each as the report shows it)
	int done = code.Count - missingCode.Count;
	report.AppendLine($"  the program's: {done} of {code.Count} translated; the ribbon's: {ribbon.Keys.Count(k => !k.EndsWith(".Keytip")) - missingRibbon.Count} of {ribbon.Keys.Count(k => !k.EndsWith(".Keytip"))} (and {ribbon.Keys.Count(k => k.EndsWith(".Keytip") && Has(k))} keytips)");
	foreach (var (origin, table) in tables) {
		string FileOf(string o) => Shown(Path.Combine(LangDir(o), lang + ".json"));
		foreach (var (key, value) in table) {
			if (value.Trim() == "") Error($"\"{key}\" is empty.");
			if (code.TryGetValue(key, out var e)) {
				if (e.Origin != origin) Error($"\"{key}\" is in {FileOf(origin)}, but it is {(e.Origin == "" ? "AhkWord's" : "the extension " + e.Origin + "'s")}: it goes in {FileOf(e.Origin)}.");
				if (e.Formatted) {
					var want = Placeholders(e.English);
					var have = Placeholders(value);
					if (have == null) Error($"\"{key}\" isn't a format string (a {{ or }} not doubled?): \"{value}\".");
					else if (want != null && !have.SetEquals(want)) Error($"\"{key}\": placeholders {Show(have)}, the English's {Show(want)}: \"{value}\".");
				}
			} else if (ribbon.ContainsKey(key)) {
				if (origin != "" && tables[""].ContainsKey(key)) Error($"\"{key}\" is in {FileOf("")} and in {FileOf(origin)}: a string goes in one file.");
			} else unused.Add($"{key}  {Q(value)}" + (origin == "" ? "" : $"  ({FileOf(origin)})"));
		}
	}
	if (missingRibbon.Count > 0) {
		report.AppendLine($"  ribbon strings it lacks ({missingRibbon.Count}):");
		foreach (var k in missingRibbon) report.AppendLine($"    {k}  {Q(ribbon[k])}");
	}
	if (missingCode.Count > 0) {
		report.AppendLine($"  the program's strings it lacks ({missingCode.Count}):");
		foreach (var (k, e) in missingCode) report.AppendLine($"    {k}  {Q(e.English)}  ({e.Where})");
	}
	if (unused.Count > 0) {
		report.AppendLine($"  keys nothing uses any more ({unused.Count}):");
		foreach (var line in unused) report.AppendLine("    " + line);
	}
	var review = changed.Keys.Where(Has).ToList();
	if (review.Count > 0) {
		report.AppendLine($"  English changed since the last check, translation to look at again ({review.Count}):");
		foreach (var k in review) report.AppendLine($"    {k}  {Q(changed[k])} -> {Q(english[k])}; now {Q(tables.Values.First(t => t.ContainsKey(k))[k])}");
	}
}
report.AppendLine().Append(errors == 0 ? "No errors." : $"{errors} error{(errors == 1 ? "" : "s")}.");
AHK.Return(report.ToString());
return errors == 0 ? 0 : 1;

// The language files' folder: AhkWord's (origin "") or an extension's; a file's name as the report shows it.
string LangDir(string origin) => origin == "" ? langDir : Path.Combine(root, origin, "lang");
string Shown(string path) => Path.GetRelativePath(root, path);

// A language's file (as Lang.Read: one object of key -> string, comments allowed).
static Dictionary<string, string> ReadTable(string path) {
	var table = new Dictionary<string, string>(StringComparer.Ordinal);
	using var doc = JsonDocument.Parse(File.ReadAllText(path), new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true });
	if (doc.RootElement.ValueKind != JsonValueKind.Object) throw new InvalidDataException("not a JSON object of key -> string");
	foreach (var p in doc.RootElement.EnumerateObject()) {
		if (p.Value.ValueKind != JsonValueKind.String) throw new InvalidDataException($"\"{p.Name}\" is not a string");
		if (table.ContainsKey(p.Name)) throw new InvalidDataException($"\"{p.Name}\" is there twice");
		table[p.Name] = p.Value.GetString()!;
	}
	return table;
}

// The indices of a format string's placeholders ({0}, {1:N0}, {2,5}); null if it isn't one.
static HashSet<int>? Placeholders(string s) {
	var set = new HashSet<int>();
	for (int i = 0; i < s.Length; i++) {
		if (s[i] == '{') {
			if (i + 1 < s.Length && s[i + 1] == '{') { i++; continue; }
			int end = s.IndexOf('}', i);
			if (end < 0) return null;
			var m = Regex.Match(s[(i + 1)..end], @"^\s*(\d+)\s*(,\s*-?\d+\s*)?(:[^{}]*)?$");
			if (!m.Success) return null;
			set.Add(int.Parse(m.Groups[1].Value));
			i = end;
		} else if (s[i] == '}') {
			if (i + 1 < s.Length && s[i + 1] == '}') { i++; continue; }
			return null;
		}
	}
	return set;
}
// A string as the report shows it: quoted, its line breaks as \n.
static string Q(string s) => "\"" + s.Replace("\r", "").Replace("\n", "\\n") + "\"";
static string Show(HashSet<int> set) => set.Count == 0 ? "none" : string.Join(" ", set.OrderBy(n => n).Select(n => "{" + n + "}"));

static int Line(string text, int pos) => text.AsSpan(0, Math.Min(pos, text.Length)).Count('\n') + 1;
static string Excerpt(string text, int pos) {
	int end = text.IndexOfAny(new[] { '\r', '\n' }, pos);
	string s = text[pos..(end < 0 ? text.Length : end)];
	return s.Length > 80 ? s[..80] + "..." : s;
}

// A string literal, or literals joined with +; null when it's something else.
static string? ReadLiteral(List<Tok> toks, ref int j) {
	if (j >= toks.Count || toks[j].Kind != K.Str) return null;
	var sb = new StringBuilder(toks[j++].Text);
	while (j + 1 < toks.Count && toks[j].Is('+') && toks[j + 1].Kind == K.Str) { sb.Append(toks[j + 1].Text); j += 2; }
	if (j < toks.Count && toks[j].Is('+')) return null;    // (a literal + something else)
	return sb.ToString();
}

// C# source to tokens, enough to find calls: identifiers, punctuation, and string literals (Str: a regular or verbatim
// one, its value; IStr: an interpolated or raw one, whose holes' code is lexed in its place). Comments are skipped.
// inHole: the code of an interpolation hole, which ends at its unmatched }.
static void Lex(string s, ref int i, bool inHole, List<Tok> toks) {
	int depth = 0;
	while (i < s.Length) {
		char c = s[i];
		if (char.IsWhiteSpace(c)) { i++; continue; }
		if (c == '/' && i + 1 < s.Length && s[i + 1] == '/') { while (i < s.Length && s[i] != '\n') i++; continue; }
		if (c == '/' && i + 1 < s.Length && s[i + 1] == '*') { int end = s.IndexOf("*/", i + 2); i = end < 0 ? s.Length : end + 2; continue; }
		if (c == '\'') {    // a char literal
			i++;
			if (s[i] == '\\') i += 2; else i++;
			while (s[i] != '\'') i++;
			i++;
			toks.Add(new Tok(K.Other, "'", i));
			continue;
		}
		// string literals: "...", @"...", $"...", $@"...", @$"...", """...""" (raw), $"""...""" (raw, interpolated)
		int start = i;
		bool verbatim = false, interpolated = false;
		int j = i;
		while (j < s.Length && (s[j] == '@' || s[j] == '$') && j - i < 3) { if (s[j] == '@') verbatim = true; else interpolated = true; j++; }
		if (j < s.Length && s[j] == '"' && (j > i || c == '"')) {
			i = j;
			if (!verbatim && i + 2 < s.Length && s[i + 1] == '"' && s[i + 2] == '"') {    // raw
				int quotes = 0;
				while (i < s.Length && s[i] == '"') { quotes++; i++; }
				string close = new('"', quotes);
				int end = s.IndexOf(close, i);
				if (end < 0) throw new FormatException("a raw string that doesn't end");
				i = end + quotes;
				toks.Add(new Tok(K.IStr, "", start));
				continue;
			}
			i++;
			var value = new StringBuilder();
			while (true) {
				if (i >= s.Length) throw new FormatException("a string that doesn't end");
				char d = s[i];
				if (verbatim) {
					if (d == '"') { if (i + 1 < s.Length && s[i + 1] == '"') { value.Append('"'); i += 2; continue; } i++; break; }
				} else {
					if (d == '"') { i++; break; }
					if (d == '\\') { value.Append(Escape(s, ref i)); continue; }
					if (d == '\n') throw new FormatException("a line break in a string");
				}
				if (interpolated && d == '{') {
					if (i + 1 < s.Length && s[i + 1] == '{') { value.Append('{'); i += 2; continue; }
					i++;
					Lex(s, ref i, true, toks);    // (the hole's code, up to its })
					continue;
				}
				if (interpolated && d == '}' && i + 1 < s.Length && s[i + 1] == '}') { value.Append('}'); i += 2; continue; }
				value.Append(d);
				i++;
			}
			toks.Add(new Tok(interpolated ? K.IStr : K.Str, value.ToString(), start));
			continue;
		}
		if (char.IsLetter(c) || c == '_' || (c == '@' && i + 1 < s.Length && (char.IsLetter(s[i + 1]) || s[i + 1] == '_'))) {
			if (c == '@') i++;
			int from = i;
			while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '_')) i++;
			toks.Add(new Tok(K.Ident, s[from..i], from));
			continue;
		}
		if (char.IsDigit(c)) { while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '.' || s[i] == '_')) i++; toks.Add(new Tok(K.Other, "0", start)); continue; }
		if (inHole) {
			if (c == '{') depth++;
			else if (c == '}') { if (depth == 0) { i++; return; } depth--; }
		}
		toks.Add(new Tok(K.Punct, c.ToString(), i));
		i++;
	}
	if (inHole) throw new FormatException("an interpolation hole that doesn't end");
}

static string Escape(string s, ref int i) {
	char e = s[i + 1];
	i += 2;
	switch (e) {
		case 'n': return "\n";
		case 't': return "\t";
		case 'r': return "\r";
		case '0': return "\0";
		case 'u': { string hex = s.Substring(i, 4); i += 4; return ((char)Convert.ToInt32(hex, 16)).ToString(); }
		case 'x': { int from = i; while (i < s.Length && i - from < 4 && Uri.IsHexDigit(s[i])) i++; return ((char)Convert.ToInt32(s[from..i], 16)).ToString(); }
		default: return e.ToString();    // \" \\ \'
	}
}

enum K { Ident, Punct, Str, IStr, Other }
record Tok(K Kind, string Text, int Pos) { public bool Is(char c) => Kind == K.Punct && Text.Length == 1 && Text[0] == c; }
record Entry(string English, bool Formatted, string Where, string Origin);
