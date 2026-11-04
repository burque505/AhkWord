// Takes the ribbon's strings out of its markup (multilingual, AhkWord-es): each command's LabelTitle, LabelDescription,
// TooltipTitle, TooltipDescription and Keytip go into a JSON file, by command name with its Id, and the markup is written
// again without them, to be compiled. The Windows Ribbon asks for whatever its compiled markup lacks when it makes a
// control, and AhkWord gives each command these strings (rte_strings.cs), or a language's; a string compiled in would
// never be asked for. Run by build_ribbon.ahk. RteRibbon.xml itself keeps its strings: it stays where English is edited.
// args: [0] the markup, [1] the markup to write without its strings, [2] the strings to write (JSON)
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Xml.Linq;

XNamespace ns = "http://schemas.microsoft.com/windows/2009/Ribbon";
string[] names = { "LabelTitle", "LabelDescription", "TooltipTitle", "TooltipDescription", "Keytip" };
var doc = XDocument.Load(args[0], LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
var commands = doc.Root?.Element(ns + "Application.Commands") ?? throw new Exception("No <Application.Commands> in " + args[0]);
// (a string written as an element, <Command.LabelTitle>, is not taken out: refused, so that none stays compiled in)
foreach (var e in doc.Descendants())
	if (e.Name.Namespace == ns && names.Any(n => e.Name.LocalName == "Command." + n))
		throw new Exception($"<{e.Name.LocalName}> (line {((System.Xml.IXmlLineInfo)e).LineNumber}): write it as an attribute of <Command>.");

using var json = new MemoryStream();
int count = 0, strings = 0;
var ids = new HashSet<uint>();
using (var w = new Utf8JsonWriter(json, new JsonWriterOptions { Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping })) {
	w.WriteStartObject();
	foreach (var c in commands.Elements(ns + "Command")) {
		string name = (string?)c.Attribute("Name") ?? throw new Exception("A <Command> without a Name.");
		if (!uint.TryParse((string?)c.Attribute("Id"), out uint id)) throw new Exception($"{name} has no Id: AhkWord finds a command's strings by its Id.");
		if (!ids.Add(id)) throw new Exception($"{name}: Id {id} is used twice.");
		w.WriteStartObject(name);
		w.WriteNumber("Id", id);
		foreach (var n in names)
			if (c.Attribute(n) is { } a) { w.WriteString(n, a.Value); a.Remove(); strings++; }
		w.WriteEndObject();
		count++;
	}
	w.WriteEndObject();
}
doc.Save(args[1], SaveOptions.DisableFormatting);
File.WriteAllBytes(args[2], json.ToArray().Concat("\n"u8.ToArray()).ToArray());
AHK.Return($"{strings} strings of {count} commands taken out of the markup: {Path.GetFileName(args[2])}");
