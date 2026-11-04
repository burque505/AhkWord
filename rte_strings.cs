// Multilingual (AhkWord-es), step 1: the ribbon's strings. build_ribbon.ahk compiles the ribbon without its labels,
// keytips and tooltips, and writes them, as ribbon\RteRibbon.xml has them (English), into ribbon\RteRibbon.strings.json.
// The Windows Ribbon asks for whatever its compiled markup lacks when it makes a control, on a tab's first showing too:
// so each command is given a control object holding its strings before the ribbon loads, and that is all. One ribbon
// then serves every language. (A string compiled in is never asked for: it could be changed only by telling the ribbon,
// command by command, and again for each tab when it first shows; telling it for all commands at once blanks every
// string that has no value. Seen in tests 231 and 232.)
//
// WindowsRibbon answers the ribbon only for commands that have a control object, and a control only for the properties
// its class has (no label for the File menu, the font control, a split button): a command without a control gets a
// group's control as a holder, which any command can have, and a control without a property gets that property's
// provider. Both by reflection, as SetRibbonLabel: WindowsRibbon keeps them private.

//
// Step 2: the interface's language (Lang). English is the strings as written; another language is a table,
// lang\<code>.json beside the program, whose strings are shown instead (the ribbon's keyed "cmdPaste.LabelTitle"), and
// whatever it lacks stays English.
//
// Step 4: the program's own strings (Tr): T("msg.saveChanges", "Save changes to {0}?", name) is the language's string
// for that key, or the English written beside it. The English stays in the code, where it's read and edited; the key
// keeps the translation when the English is reworded. lang\check_lang.ahk lists what a language lacks.

global using static Tr;    // (T and TN in every file)
using System.Globalization;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using RibbonLib;
using RibbonLib.Controls;
using RibbonLib.Controls.Properties;
using RibbonLib.Interop;

/// The interface's language: the setting Language, or the AHKWORD_LANG environment variable over it (for tests, and a
/// try-out without changing the settings). "en", or a language without a file, is English. Read once per process, when
/// the first window builds its ribbon: a change shows from the next start.
static class Lang
{
	public const string English = "en";
	/// The language shown (its file's name: "es"), English when none.
	public static string Code { get; private set; } = English;
	static Dictionary<string, string> table = new();
	static string? error;
	static bool loaded;

	/// The language's string for key; null: English (none in its table).
	public static string? Get(string key) => table.TryGetValue(key, out var s) ? s : null;
	public static int Count => table.Count;

	/// The language file that couldn't be read (once: then null), for a message when the window shows.
	public static string? TakeError() { var e = error; error = null; return e; }

	/// The language's table from lang\<code>.json in appDir ("es-MX": es-MX.json, else es.json). A file that can't be read
	/// leaves English, and its error for TakeError.
	public static void Load(string appDir, string? setting) {
		if (loaded) return;
		loaded = true;
		string wanted = (Environment.GetEnvironmentVariable("AHKWORD_LANG") is { Length: > 0 } env ? env : setting ?? "").Trim();
		if (wanted == "" || wanted.Equals(English, StringComparison.OrdinalIgnoreCase) || !Regex.IsMatch(wanted, "^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$")) return;
		string folder = Path.Combine(appDir, "lang");
		string? path = new[] { wanted, wanted.Split('-')[0] }.Select(c => Path.Combine(folder, c + ".json")).FirstOrDefault(File.Exists);
		if (path == null) return;
		try {
			table = Read(path);
			Code = Path.GetFileNameWithoutExtension(path);
		} catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidDataException) {
			error = $"The language file {path} could not be read, so AhkWord is in English:\n\n{ex.Message}";
		}
	}

	/// A language file: one object of key -> string (comments allowed).
	public static Dictionary<string, string> Read(string path) {
		var strings = new Dictionary<string, string>();
		using var doc = JsonDocument.Parse(File.ReadAllText(path), new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true });
		if (doc.RootElement.ValueKind != JsonValueKind.Object) throw new InvalidDataException($"{path}: not a JSON object of key -> string.");
		foreach (var p in doc.RootElement.EnumerateObject()) {
			if (p.Value.ValueKind != JsonValueKind.String) throw new InvalidDataException($"{path}: \"{p.Name}\" is not a string.");
			strings[p.Name] = p.Value.GetString()!;
		}
		return strings;
	}
}

/// The program's strings in the interface's language (Lang), English when it has none.
/// Keys by where the string shows: "msg." a message box, "status." the status bar, "menu." a context menu, "ui." the
/// window's own texts, "dlg.<Dialog>." a dialog's. Key and English are written as literals, so that check_lang.cs can find
/// them; a key is used with one English only. Formatting, with arguments: {0}, {1} ... as string.Format, in the user's
/// culture; a translation whose placeholders don't fit its arguments shows the English.
static class Tr
{
	/// key's string in the language, or english: as it is (braces too).
	public static string T(string key, string english) => Lang.Get(key) ?? english;

	/// key's string in the language, or english, formatted with args.
	public static string T(string key, string english, params object?[] args) => Format(Lang.Get(key), english, args);

	/// A count's string: key + ".one" for 1, key + ".other" for any other count (English's and Spanish's rule), formatted
	/// with the count as {0} and args after it: TN("status.restyled", "{0} paragraph restyled.", "{0} paragraphs restyled.", n).
	public static string TN(string key, string one, string other, int count, params object?[] args) {
		bool single = count == 1;
		return Format(Lang.Get(key + (single ? ".one" : ".other")), single ? one : other, args.Prepend(count).ToArray());
	}

	/// A shortcut as shown, its key names in the language: "Ctrl+Shift+Del" -> "Ctrl+Mayús+Supr". Only for showing: a
	/// shortcut kept (settings, the macro API) stays as Keys or KeysConverter write it.
	public static string TKeys(string shortcut) => string.Join("+", shortcut.Split('+').Select(part => part.Trim() switch {
		"Shift" => T("key.Shift", "Shift"),
		"Del" => T("key.Del", "Del"),
		"Delete" => T("key.Delete", "Delete"),
		"Ins" => T("key.Ins", "Ins"),
		"Insert" => T("key.Insert", "Insert"),
		"Enter" => T("key.Enter", "Enter"),
		"Return" => T("key.Return", "Return"),
		"Space" => T("key.Space", "Space"),
		"Home" => T("key.Home", "Home"),
		"End" => T("key.End", "End"),
		"PgUp" => T("key.PgUp", "PgUp"),
		"PageUp" => T("key.PageUp", "PageUp"),
		"PgDn" => T("key.PgDn", "PgDn"),
		"PageDown" => T("key.PageDown", "PageDown"),
		"Next" => T("key.Next", "Next"),
		"Back" => T("key.Back", "Back"),
		"Backspace" => T("key.Backspace", "Backspace"),
		"Up" => T("key.Up", "Up"),
		"Down" => T("key.Down", "Down"),
		"Left" => T("key.Left", "Left"),
		"Right" => T("key.Right", "Right"),
		_ => part,
	}));

	static string Format(string? translated, string english, object?[] args) {
		if (translated != null)
			try { return string.Format(CultureInfo.CurrentCulture, translated, args); }
			catch (FormatException) { }    // (a placeholder it shouldn't have: check_lang.cs reports it)
		return string.Format(CultureInfo.CurrentCulture, english, args);
	}
}

static class RibbonStrings
{
	public const string FileName = "RteRibbon.strings.json";
	static readonly string[] Names = { "LabelTitle", "LabelDescription", "TooltipTitle", "TooltipDescription", "Keytip" };

	/// The strings file: command name -> its Id and its strings (markup attribute -> text).
	public static Dictionary<string, (uint Id, Dictionary<string, string> Strings)> Load(string path) {
		var commands = new Dictionary<string, (uint, Dictionary<string, string>)>();
		using var doc = JsonDocument.Parse(File.ReadAllText(path), new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true });
		foreach (var command in doc.RootElement.EnumerateObject()) {
			uint id = 0;
			var strings = new Dictionary<string, string>();
			foreach (var p in command.Value.EnumerateObject()) {
				if (p.Name == "Id") id = p.Value.GetUInt32();
				else if (Names.Contains(p.Name)) strings[p.Name] = p.Value.GetString() ?? "";
				else throw new InvalidDataException($"{path}: {command.Name} has \"{p.Name}\", which is none of Id, {string.Join(", ", Names)}.");
			}
			if (id == 0) throw new InvalidDataException($"{path}: {command.Name} has no Id.");
			commands[command.Name] = (id, strings);
		}
		return commands;
	}

	static readonly FieldInfo ControlsField = typeof(Ribbon).GetField("_mapRibbonControls", BindingFlags.NonPublic | BindingFlags.Instance)
		?? throw new MissingFieldException("WindowsRibbon's Ribbon._mapRibbonControls (another version of WindowsRibbon?)");
	static readonly FieldInfo PropertiesField = typeof(BaseRibbonControl).GetField("_mapProperties", BindingFlags.NonPublic | BindingFlags.Instance)
		?? throw new MissingFieldException("WindowsRibbon's BaseRibbonControl._mapProperties (another version of WindowsRibbon?)");

	/// A command's strings in the language (Lang), over the English ones: a property the language has replaces English's.
	public static Dictionary<string, string> InLanguage(string command, Dictionary<string, string> english) {
		var strings = new Dictionary<string, string>(english);
		foreach (var property in Names)
			if (Lang.Get(command + "." + property) is string text) strings[property] = text;
		return strings;
	}

	/// Gives every command of the strings file in folder its strings, in the language (Lang): to the control AhkWord made
	/// for it, or to a holder.
	/// After every control is made (a holder would take a command's place) and before the ribbon loads (until then the
	/// setters only keep the values: later they would tell the ribbon with a key that isn't pinned; see SetRibbonLabel).
	public static int Give(Ribbon ribbon, string folder) {
		if (ribbon.Initialized) throw new InvalidOperationException("The ribbon's strings go in before it loads.");
		string path = Path.Combine(folder, FileName);
		if (!File.Exists(path))
			throw new FileNotFoundException($"The ribbon's strings are missing: {path}. Build them with ribbon\\build_ribbon.ahk, which makes them with the ribbon.");
		var controls = (Dictionary<uint, BaseRibbonControl>)ControlsField.GetValue(ribbon)!;
		int given = 0;
		foreach (var (name, (id, english)) in Load(path)) {
			var strings = InLanguage(name, english);
			if (strings.Count == 0) continue;
			if (!controls.TryGetValue(id, out var control)) control = new RibbonGroup(ribbon, id);
			var properties = (Dictionary<PropertyKey, IPropertiesProvider>)PropertiesField.GetValue(control)!;
			// the control's provider of a property, or one added to it (for the keys it has none for)
			T Provider<T>(PropertyKey key, Func<IPropertiesProvider> make) where T : class {
				if (properties.TryGetValue(key, out var found) && found is T mine) return mine;
				var made = make();
				foreach (var k in made.SupportedProperties) properties.TryAdd(k, made);
				properties[key] = made;
				return (T)made;
			}
			foreach (var (property, text) in strings) {
				switch (property) {
					case "LabelTitle": Provider<ILabelPropertiesProvider>(RibbonProperties.Label, () => new LabelPropertiesProvider(ribbon, id)).Label = text; break;
					case "LabelDescription": Provider<ILabelDescriptionPropertiesProvider>(RibbonProperties.LabelDescription, () => new LabelDescriptionPropertiesProvider(ribbon, id)).LabelDescription = text; break;
					case "TooltipTitle": Provider<ITooltipPropertiesProvider>(RibbonProperties.TooltipTitle, () => new TooltipPropertiesProvider(ribbon, id)).TooltipTitle = text; break;
					case "TooltipDescription": Provider<ITooltipPropertiesProvider>(RibbonProperties.TooltipDescription, () => new TooltipPropertiesProvider(ribbon, id)).TooltipDescription = text; break;
					case "Keytip": Provider<IKeytipPropertiesProvider>(RibbonProperties.Keytip, () => new KeytipPropertiesProvider(ribbon, id)).Keytip = text; break;
				}
				given++;
			}
		}
		return given;
	}
}
