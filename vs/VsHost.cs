// VS build trial: what Csharp-in-AHK's host gives AhkWord.cs and the exe has to give itself.

/// <summary>Csharp-in-AHK's AHK.Return: hands a value back to AutoHotkey's RunCSharp. The exe has nobody to hand it to.</summary>
public static class AHK
{
	public static object? Result { get; private set; }

	public static void Return(object? value) => Result = value;
}
