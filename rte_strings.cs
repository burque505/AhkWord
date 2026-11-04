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

using System.Reflection;
using System.Text.Json;
using RibbonLib;
using RibbonLib.Controls;
using RibbonLib.Controls.Properties;
using RibbonLib.Interop;

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

	/// Gives every command of the strings file in folder its strings: to the control AhkWord made for it, or to a holder.
	/// After every control is made (a holder would take a command's place) and before the ribbon loads (until then the
	/// setters only keep the values: later they would tell the ribbon with a key that isn't pinned; see SetRibbonLabel).
	public static int Give(Ribbon ribbon, string folder) {
		if (ribbon.Initialized) throw new InvalidOperationException("The ribbon's strings go in before it loads.");
		string path = Path.Combine(folder, FileName);
		if (!File.Exists(path))
			throw new FileNotFoundException($"The ribbon's strings are missing: {path}. Build them with ribbon\\build_ribbon.ahk, which makes them with the ribbon.");
		var controls = (Dictionary<uint, BaseRibbonControl>)ControlsField.GetValue(ribbon)!;
		int given = 0;
		foreach (var (name, (id, strings)) in Load(path)) {
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
