// Where RTE Writer keeps its files (AppFolders: portable mode), and the recovery copies it autosaves (Recovery), for
// RTE Writer (rte_writer29.cs, through #load).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;

// RTE Writer's own files: its settings in %APPDATA%\RTE Writer, the converters' state (LibreOffice's profile, the log)
// in %LOCALAPPDATA%\RTE Writer; in portable mode (a file named "portable" or "portable.txt" beside the program) all of
// them in a "data" folder beside it instead, so that the folder can be carried to another computer with them. (The
// macros are beside the program already, in rte_macros. The NuGet packages and the C# compiler stay in the user's
// caches, and are fetched again on a computer that hasn't got them.)
public static class AppFolders
{
	/// The program's folder (rte_writer*.ahk's): set before anything else is read.
	public static string ProgramFolder { get; set; } = AppContext.BaseDirectory;

	public static bool Portable => File.Exists(Path.Combine(ProgramFolder, "portable")) || File.Exists(Path.Combine(ProgramFolder, "portable.txt"));

	/// The settings' folder (roaming).
	public static string Roaming => Portable ? Path.Combine(ProgramFolder, "data")
		: Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RTE Writer");

	/// The converters' state (this computer's).
	public static string Local => Portable ? Path.Combine(ProgramFolder, "data")
		: Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RTE Writer");
}

// The recovery copy of the document being edited: written (as RTF, with its header, footer and line numbering) while
// it has unsaved changes, and removed when it is saved, closed or discarded, and when RTE Writer closes. A copy left
// behind is a document whose RTE Writer ended without closing (a crash, a power cut, Windows restarting): the next
// start offers it. Each RTE Writer has copies of its own (named by its process), so two running at once don't take each
// other's for left behind.
public sealed class Recovery
{
	public sealed record Info(string? FileName, string Name, DateTime Saved, int ProcessId, long ProcessStart);

	/// A copy left behind: what it was, and its RTF file.
	public sealed record Orphan(Info Info, string RtfPath, string InfoPath);

	readonly string folder, id;
	readonly int processId;
	readonly long processStart;
	string? written;    // the RTF written last (null: no copy now)

	public Recovery(string folder) {
		this.folder = folder;
		using var process = Process.GetCurrentProcess();
		processId = process.Id;
		processStart = process.StartTime.ToUniversalTime().Ticks;
		id = $"{processId}-{processStart}";
	}

	public string Folder => folder;
	string RtfPath => Path.Combine(folder, id + ".rtf");
	string InfoPath => Path.Combine(folder, id + ".json");

	/// Whether there is a copy now.
	public bool HasCopy => written != null;

	/// Writes the copy (unless it is as written last); fileName: the document's file (null: not saved yet).
	public void Write(string rtf, string? fileName, string name) {
		if (rtf == written) return;
		Directory.CreateDirectory(folder);
		// (written beside, then moved over: a crash while writing leaves the copy before)
		string temporary = RtfPath + ".tmp";
		File.WriteAllText(temporary, rtf, Encoding.Latin1);
		File.Move(temporary, RtfPath, true);
		File.WriteAllText(InfoPath, JsonSerializer.Serialize(new Info(fileName, name, DateTime.Now, processId, processStart)));
		written = rtf;
	}

	/// Removes the copy (saved, closed or discarded).
	public void Clear() {
		if (written == null && !File.Exists(InfoPath)) return;
		Delete(RtfPath, InfoPath);
		written = null;
	}

	static void Delete(params string[] paths) {
		foreach (var path in paths) try { File.Delete(path); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
	}

	/// The copies in the folder whose RTE Writer isn't running any more, the newest first.
	public static List<Orphan> Orphans(string folder) {
		var orphans = new List<Orphan>();
		if (!Directory.Exists(folder)) return orphans;
		foreach (var infoPath in Directory.GetFiles(folder, "*.json")) {
			string rtfPath = Path.ChangeExtension(infoPath, ".rtf");
			Info? info;
			try { info = JsonSerializer.Deserialize<Info>(File.ReadAllText(infoPath)); } catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { continue; }
			if (info == null || !File.Exists(rtfPath) || Running(info)) continue;
			orphans.Add(new Orphan(info, rtfPath, infoPath));
		}
		return orphans.OrderByDescending(o => o.Info.Saved).ToList();
	}

	// (a process with the copy's id and start time: its RTE Writer is still running)
	static bool Running(Info info) {
		try {
			using var process = Process.GetProcessById(info.ProcessId);
			return process.StartTime.ToUniversalTime().Ticks == info.ProcessStart;
		} catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) {
			return false;
		}
	}

	/// The RTF of a copy left behind.
	public static string Read(Orphan orphan) => File.ReadAllText(orphan.RtfPath, Encoding.Latin1);

	/// Removes a copy left behind (recovered or discarded).
	public static void Discard(Orphan orphan) => Delete(orphan.RtfPath, orphan.InfoPath);
}
