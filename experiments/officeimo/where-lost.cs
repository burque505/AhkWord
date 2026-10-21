	// Experiment (iteration 47), made-up text: where OfficeIMO loses a paragraph's alignment and RichEdit's own list
	// numbering: its RTF model after parsing, then its Word document after converting, for a minimal RTF and for
	// AhkWord's own RTF of the same paragraphs.
	public string SelfTest(Action<int> pump) {
		var r = new StringBuilder();
		var doc = Editor.TextDocument;
		Run(() => api.View = "Rich");
		Run(() => api.New());
		Run(() => api.Rtf = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}\f0\fs24 \pard Plain.\par \pard\qc Centred.\par \pard\qr Right.\par \pard Alpha.\par \pard Bravo.\par }");
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		int alpha = story.IndexOf("Alpha"), bravoEnd = story.IndexOf("Bravo.") + 6;
		Run(() => { api.Select(alpha, bravoEnd - alpha); api.NumberedList(); });
		pump(300);
		string ahkword = ((RichTextBox)Editor).Rtf ?? "";
		r.AppendLine($"     AhkWord's RTF ({ahkword.Length} chars), its paragraph controls: {string.Join(" ", Regex.Matches(ahkword, @"\\(pard|q[lcrj]|pn\w*|ls\d+|ilvl\d+|par)\b").Select(m => m.Value).Distinct())}");
		foreach (var (name, rtf) in new[] {
			("minimal", @"{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}\f0\fs24 \pard Plain.\par \pard\qc Centred.\par \pard\qr Right.\par}"),
			("AhkWord's", ahkword) }) {
			var parsed = OfficeIMO.Rtf.RtfDocument.Parse(rtf);
			r.AppendLine($"     {name}: parsed paragraphs {string.Join(" | ", parsed.Paragraphs.Select(p => $"{p.Alignment}{(p.DirectAlignment == null ? "" : "*")}"))}");
			var conversion = OfficeIMO.Word.Rtf.WordRtfConverterExtensions.ToWordDocumentResult(parsed);
			using var word = conversion.Value;
			r.AppendLine($"     {name}: Word paragraphs {string.Join(" | ", word.Paragraphs.Select(p => $"{(p.Text.Length > 8 ? p.Text.Substring(0, 8) : p.Text)}:{p.ParagraphAlignment?.ToString() ?? "-"}{(p.IsListItem ? " list" : "")}"))}");
		}
		Editor.Modified = false;
		return r.ToString();
	}
	public void Capture(string path) { }
