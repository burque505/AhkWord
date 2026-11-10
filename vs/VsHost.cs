// VS build trial: what Csharp-in-AHK's host and AhkWord.ahk give AhkWord.cs, and the exe has to give itself.
// AhkWord.exe's arguments are AhkWord.ahk's RunCSharpFile ones: "AutoHotkey64.exe" "program folder" ["document"].
// Started without them (double-clicked), or with only a document (Open with), it finds both itself.
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

/// <summary>Csharp-in-AHK's AHK.Return: hands a value back to AutoHotkey's RunCSharp. The exe has nobody to hand it to.</summary>
public static class AHK
{
	public static object? Result { get; private set; }

	public static void Return(object? value)
	{
		Result = value;
		// (the test runner, regress42x.ps1, reads what a test build returns from the exe's output)
		if (VsHost.TestRun) { Console.Out.Write(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)); Console.Out.Flush(); }
	}
}

static class VsHost
{
	/// Run by the test runner (regress42x.ps1, AHKWORD_TEST_RUN=1): no companion and no message box (both would be on
	/// the user's desktop, the test's window is on a hidden one); the result and errors go to the exe's output.
	internal static readonly bool TestRun = Environment.GetEnvironmentVariable("AHKWORD_TEST_RUN") == "1";

	/// Before the program starts: it is AhkWord.exe (New Window starts it again), its arguments are completed, an error
	/// that stops it is shown (as AhkWord.ahk does), and AhkWordCompanion.ahk is started beside it for the hotkeys and
	/// the snippets' hotstrings that AhkWord.ahk would have run.
	[ModuleInitializer]
	internal static void Start()
	{
		AppFolders.ProgramExe = Environment.ProcessPath;
		AppFolders.CompleteArgs = Complete;
		AppDomain.CurrentDomain.UnhandledException += (_, e) => {
			if (TestRun) { Console.Error.WriteLine("ERROR C# " + e.ExceptionObject); Console.Error.Flush(); return; }
			try { MessageBox.Show("AhkWord stopped because of an error:\n\n" + e.ExceptionObject, "AhkWord", MessageBoxButtons.OK, MessageBoxIcon.Error); }
			catch { }
		};
		if (TestRun) return;
		string[] args = Complete(Environment.GetCommandLineArgs()[1..]);
		StartCompanion(args[0], args[1]);
	}

	/// "AutoHotkey64.exe" "program folder" ["document"], whatever was left out filled in.
	static string[] Complete(string[] args)
	{
		if (args.Length >= 2 && Directory.Exists(args[1])) return args;
		string? document = args.Length == 1 && !args[0].EndsWith(".exe", StringComparison.OrdinalIgnoreCase) ? args[0] : null;
		string ahk = args.Length == 1 && document == null ? args[0] : FindAutoHotkey();
		return document == null ? [ahk, ProgramFolder()] : [ahk, ProgramFolder(), Path.GetFullPath(document)];
	}

	/// The folder with AhkWord's files (ribbon\RteRibbon.ribbon): the exe's own, or the nearest one above it (the
	/// trial's exe is in vs\bin\...), else the current folder.
	static string ProgramFolder()
	{
		for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir != null; dir = dir.Parent)
			if (File.Exists(Path.Combine(dir.FullName, "ribbon", "RteRibbon.ribbon"))) return dir.FullName;
		return Environment.CurrentDirectory;
	}

	/// AutoHotkey v2 (64-bit): where its installer registers it, else its usual place; "" if not found.
	static string FindAutoHotkey()
	{
		foreach (var hive in new[] { Microsoft.Win32.Registry.CurrentUser, Microsoft.Win32.Registry.LocalMachine }) {
			using var key = hive.OpenSubKey(@"SOFTWARE\AutoHotkey");
			if (key?.GetValue("InstallDir") is string dir && File.Exists(Path.Combine(dir, "v2", "AutoHotkey64.exe")))
				return Path.Combine(dir, "v2", "AutoHotkey64.exe");
		}
		string usual = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "AutoHotkey", "v2", "AutoHotkey64.exe");
		return File.Exists(usual) ? usual : "";
	}

	static void StartCompanion(string ahk, string folder)
	{
		try {
			string script = Path.Combine(folder, "AhkWordCompanion.ahk");
			if (ahk == "" || !File.Exists(ahk) || !File.Exists(script)) return;    // (AhkWord works without them: no hotkeys or snippets' hotstrings)
			var start = new ProcessStartInfo(ahk) { UseShellExecute = false, WorkingDirectory = folder };
			start.ArgumentList.Add(script);
			start.ArgumentList.Add(Environment.ProcessId.ToString());
			start.Environment.Remove("RTE_WRITER_HWND");
			using var _ = Process.Start(start);
		}
		catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or IOException or InvalidOperationException) {
			Debug.WriteLine("AhkWordCompanion.ahk could not be started: " + ex.Message);
		}
	}
}
