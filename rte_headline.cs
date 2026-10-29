// Headline Case, for AhkWord (AhkWord.cs, through #load). Iteration 64.
//
// Change Case > Headline Case: a heading's or a title's words capitalized as Chicago does it, with the words a legal
// paper needs left alone (Title Case, TOM's, capitalizes every word). Written for AhkWord; its rules:
//   - every word capitalized, except the small words (a, the, and, of, to ...) inside a line: the first and the last word
//     of a line, and the first after : . ; ? ! or a dash, are capitalized whatever they are;
//   - words that are always lower case (v., vs., et al., a/k/a, d/b/a, f/k/a, e.g.), wherever they are but first;
//   - words that are always as written (LLC, L.L.C., Inc., P.C., N.A., ISAOA, the state codes that are no English word);
//   - words already in capitals (USA, NOTE) and in mixed case (McDonald, iPhone) kept, unless the whole line is in
//     capitals ("DEED OF TRUST" becomes "Deed of Trust"); initials with periods (u.s.) and Roman numerals (iii) in
//     capitals; ordinals (1st, 2ND) in lower case after the number; each part of a hyphenated or slashed word on its own
//     (Self-Employed, Made-Up, Run-of-the-Mill, and/or); O' and D' (O'Brien).
// The word lists are a text file beside the settings (headline-case.txt, made the first time with the lists here), to be
// edited: sections [small], [lower] and [exact], words separated by spaces or lines, # starts a comment.
// Only letters change case, one for one, so the text keeps its length and each character its formatting; hidden text
// (the cross-references', tables of contents' and footnotes' markers) is passed over as if it were a space.
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

public static class HeadlineCase
{
	public static string FilePath => Path.Combine(Path.GetDirectoryName(Settings.FilePath) is { Length: > 0 } dir ? dir : AppFolders.Roaming, "headline-case.txt");

	const string DefaultLists = """
		# AhkWord: Headline Case's word lists (Change Case > Headline Case). Edit them as you like; AhkWord reads this
		# file again when it has changed. Words are separated by spaces or lines; # starts a comment.

		[small]
		# Lower case inside a line; capitalized as the first or last word of a line, or after : . ; ? ! or a dash.
		a an the
		and but for nor or
		as at by en from in into of off on onto out over per re than to up upon via with

		[lower]
		# Always lower case, wherever they are (except as a line's first word).
		v. vs. et al. etc. e.g. i.e. a/k/a d/b/a f/k/a n/k/a c/o a.m. p.m.

		[exact]
		# Always as written here, wherever they are.
		LLC L.L.C. LLP L.L.P. PLLC Inc. Inc Corp. Co. Ltd. L.P. LP P.C. P.A. N.A. FSB
		ISAOA ATIMA MERS HUD FHA VA USDA HOA PUD IRS EIN SSN UCC DBA AKA FKA
		U.S. USA U.S.A. D.C. DC
		AK AZ AR CA CT FL GA IA KS KY MI MN MO MT NC ND NE NH NJ NM NV NY RI SC SD TN TX UT VT WA WI WV WY PR
		""";

	sealed class Lists
	{
		public readonly HashSet<string> Small = new(StringComparer.OrdinalIgnoreCase), Lower = new(StringComparer.OrdinalIgnoreCase);
		public readonly Dictionary<string, string> Exact = new(StringComparer.OrdinalIgnoreCase);
	}

	static Lists? lists;
	static DateTime listsTime;
	static string? listsPath;

	// The lists, from the file (made with the lists above the first time; if it can't be read or made, the lists above).
	static Lists Current() {
		string path = FilePath;
		try {
			if (!File.Exists(path)) {
				Directory.CreateDirectory(Path.GetDirectoryName(path)!);
				File.WriteAllText(path, DefaultLists.Replace("\n", "\r\n").Replace("\r\r\n", "\r\n") + "\r\n");
			}
			var time = File.GetLastWriteTimeUtc(path);
			if (lists == null || time != listsTime || path != listsPath) {
				lists = Parse(File.ReadAllText(path));
				listsTime = time;
				listsPath = path;
			}
			return lists;
		} catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
			return lists ??= Parse(DefaultLists);
		}
	}

	static Lists Parse(string text) {
		var result = new Lists();
		string section = "";
		foreach (string raw in text.Split('\n')) {
			string line = raw.Split('#')[0].Trim();
			if (line.StartsWith('[') && line.EndsWith(']')) { section = line[1..^1].Trim().ToLowerInvariant(); continue; }
			foreach (string word in line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
				if (section == "small") result.Small.Add(word);
				else if (section == "lower") result.Lower.Add(word);
				else if (section == "exact") result.Exact[word] = word;
			}
		}
		return result;
	}

	// A word: letters and digits, joined by ' . / - & (L.L.C., a/k/a, O'Brien, Self-Employed, AT&T), and a period after it.
	static readonly Regex Word = new(@"[\p{L}\p{N}]+(?:['’./&-][\p{L}\p{N}]+)*\.?", RegexOptions.Compiled);
	static readonly Regex Roman = new(@"^(?=[ivx]{2,}$)x{0,3}(ix|iv|v?i{0,3})$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
	static readonly Regex Ordinal = new(@"^\d+(st|nd|rd|th|d)\.?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
	static readonly Regex Initials = new(@"^(\p{L}\.){2,}$|^(\p{L}\.)+\p{L}$", RegexOptions.Compiled);

	/// The text in headline case. hidden: the characters to pass over (left as they are), or null.
	public static string Apply(string text, bool[]? hidden = null) {
		var lists = Current();
		var culture = CultureInfo.CurrentCulture;
		// What is looked at: hidden text as spaces; each line (paragraph, table cell) on its own, a line break (\v) and a
		// tab within a line
		var seen = new StringBuilder(text);
		for (int i = 0; i < seen.Length; i++) if (hidden != null && i < hidden.Length && hidden[i]) seen[i] = ' ';
		var result = text.ToCharArray();
		int lineStart = 0;
		for (int i = 0; i <= seen.Length; i++) {
			if (i < seen.Length && (seen[i] >= ' ' || seen[i] == '\v' || seen[i] == '\t')) continue;
			Line(seen.ToString(lineStart, i - lineStart), lineStart);
			lineStart = i + 1;
		}
		return new string(result);

		void Line(string line, int offset) {
			var words = Word.Matches(line);
			if (words.Count == 0) return;
			bool allCaps = !line.Any(char.IsLower);
			bool phraseStart = true;
			for (int w = 0; w < words.Count; w++) {
				var m = words[w];
				// (what stands between this word and the one before: a colon, a semicolon, a dash, ? or ! start a phrase)
				if (w > 0) {
					string between = line[(words[w - 1].Index + words[w - 1].Length)..m.Index];
					if (between.IndexOfAny([':', ';', '?', '!', '—', '–']) >= 0 || between.Contains(" - ")) phraseStart = true;
				}
				bool first = w == 0, edge = phraseStart || w == words.Count - 1;
				string cased = Case(m.Value, first, edge, allCaps);
				for (int k = 0; k < cased.Length; k++) result[offset + m.Index + k] = cased[k];
				// (a period that ends a sentence starts a phrase: not an abbreviation's, an initial's or one in the lists)
				phraseStart = m.Value.EndsWith('.') && m.Value.Length > 2 && !m.Value[..^1].Contains('.')
					&& !lists.Exact.ContainsKey(m.Value) && !lists.Lower.Contains(m.Value) && !lists.Small.Contains(m.Value);
			}
		}

		string Case(string word, bool first, bool edge, bool allCaps) {
			string bare = word.EndsWith('.') ? word[..^1] : word;
			if (lists.Exact.TryGetValue(word, out var exact)) return exact;
			if (bare != word && lists.Exact.TryGetValue(bare, out exact)) return exact + ".";
			if (Ordinal.IsMatch(word)) return Lower(word);
			if (char.IsDigit(word[0])) return word;
			if (!allCaps && Kept(bare)) return word;
			if (lists.Lower.Contains(word) || lists.Lower.Contains(bare)) return first ? Single(word, true) : Lower(word);
			if (Initials.IsMatch(word)) return Upper(word);
			if (bare.IndexOfAny(['-', '/']) > 0) {
				// each part on its own: the first as the word would be, the last capitalized (Made-Up, Built-In), those
				// between unless small (Run-of-the-Mill)
				var parts = new StringBuilder();
				int start = 0;
				for (int i = 0; i <= word.Length; i++) {
					if (i < word.Length && word[i] != '-' && word[i] != '/') continue;
					string part = word[start..i];
					parts.Append(part.Length == 0 ? part : Single(part, start == 0 ? edge : i == word.Length && word[start - 1] == '-'));
					if (i < word.Length) parts.Append(word[i]);
					start = i + 1;
				}
				return parts.ToString();
			}
			return Single(word, edge);
		}

		string Single(string word, bool edge) {
			string bare = word.TrimEnd('.');
			if (lists.Exact.TryGetValue(word, out var exact) || lists.Exact.TryGetValue(bare, out exact)) return exact + word[exact.Length..];
			if (!edge && (lists.Small.Contains(word) || lists.Small.Contains(bare))) return Lower(word);
			if (Roman.IsMatch(bare)) return Upper(word);
			var chars = Lower(word).ToCharArray();
			chars[0] = char.ToUpper(chars[0], culture);
			// O'Brien, D'Angelo
			if (chars.Length > 2 && (chars[1] == '\'' || chars[1] == '’') && "OD".Contains(chars[0])) chars[2] = char.ToUpper(chars[2], culture);
			return new string(chars);
		}

		// In capitals (USA), or in mixed case (McDonald, iPhone): as written
		static bool Kept(string word) {
			bool upper = false, lower = false, upperAfterFirst = false;
			for (int i = 0; i < word.Length; i++) {
				if (char.IsUpper(word[i])) { upper = true; if (i > 0) upperAfterFirst = true; }
				else if (char.IsLower(word[i])) lower = true;
			}
			return upper && !lower && word.Count(char.IsLetter) > 1 || lower && upperAfterFirst;
		}

		string Lower(string s) { var c = s.ToCharArray(); for (int i = 0; i < c.Length; i++) c[i] = char.ToLower(c[i], culture); return new string(c); }
		string Upper(string s) { var c = s.ToCharArray(); for (int i = 0; i < c.Length; i++) c[i] = char.ToUpper(c[i], culture); return new string(c); }
	}
}
