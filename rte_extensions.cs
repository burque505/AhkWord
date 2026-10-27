// AhkWord's extension point (AhkWord.cs, through #load). In AhkWord it adds nothing.
//
// A build of AhkWord with more features (its own repository) puts its own version of this file here: it #loads its
// files, and Started adds their commands to the window. Everything else stays as in AhkWord, so the two stay one code
// base, and AhkWord's changes merge into it as they are. (Iteration 60)
static class AhkWordExtensions
{
	/// The window is open, with its first document: an extension sets itself up here.
	public static void Started(WriterForm form) { }
}
