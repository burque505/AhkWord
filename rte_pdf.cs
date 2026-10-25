#:package PDFsharp@6.2.4
// PDF bookmarks and links, for AhkWord (AhkWord.cs, through #load). Iteration 53.
//
// Export as PDF prints through Windows' "Microsoft Print to PDF", which writes the pages only. Once it has written the
// file, its bookmarks (the outline a PDF viewer shows beside the pages) are put in: one for each heading, under the one
// of the level above it, each opening its page at the heading; and each table of contents' entry becomes a link to its
// heading's page, as in a PDF that Word makes. The file opens with its bookmarks shown.
// Where each heading and entry is printed: DocumentPrinter.PdfPlaces (the printer's own layout of the pages).
// PDFsharp (MIT license up to 6.2.x; later versions are AGPL: this one, pinned) edits the file.
using System.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Annotations;

public static class PdfMarks
{
	/// A bookmark: its heading's level (1 the top) and text, the page it is printed on (0: the first) and where on the
	/// page (points from the page's top).
	public sealed record Bookmark(int Level, string Title, int Page, float Top);

	/// A link: the page it is on (0: the first) and its area there (points from the page's top left corner), and the
	/// page it goes to (0: the first) and where on it (points from the top).
	public sealed record Link(int Page, RectangleF Area, int ToPage, float ToTop);

	/// Puts the bookmarks and links into a PDF file; the file opens with its bookmarks shown (if it has any).
	public static void Add(string path, IReadOnlyList<Bookmark> bookmarks, IReadOnlyList<Link> links) {
		if (bookmarks.Count == 0 && links.Count == 0) return;
		using var document = PdfReader.Open(path, PdfDocumentOpenMode.Modify);
		int pages = document.PageCount;
		// the bookmarks, each under the last one of a level above it
		var parents = new List<(int Level, PdfOutline Outline)>();
		foreach (var mark in bookmarks) {
			if (mark.Page < 0 || mark.Page >= pages || string.IsNullOrWhiteSpace(mark.Title)) continue;
			while (parents.Count > 0 && parents[^1].Level >= mark.Level) parents.RemoveAt(parents.Count - 1);
			var into = parents.Count > 0 ? parents[^1].Outline.Outlines : document.Outlines;
			var page = document.Pages[mark.Page];
			var outline = into.Add(mark.Title.Trim(), page, true);
			outline.PageDestinationType = PdfPageDestinationType.Xyz;
			outline.Left = 0;
			// (PDF's height, from the page's foot: PDFsharp writes it as it is given; a little above the heading, as Word's)
			outline.Top = Math.Clamp(page.Height.Point - mark.Top + 6, 0, page.Height.Point);
			parents.Add((mark.Level, outline));
		}
		// the links (PDF's own coordinates: from the page's bottom left corner)
		foreach (var link in links) {
			if (link.Page < 0 || link.Page >= pages || link.ToPage < 0 || link.ToPage >= pages) continue;
			var page = document.Pages[link.Page];
			double height = page.Height.Point;
			var area = new PdfRectangle(new PdfSharp.Drawing.XPoint(link.Area.Left, height - link.Area.Bottom), new PdfSharp.Drawing.XPoint(link.Area.Right, height - link.Area.Top));
			page.AddDocumentLink(area, link.ToPage + 1);
		}
		if (document.Outlines.Count > 0) document.PageMode = PdfPageMode.UseOutlines;
		document.Save(path);
	}
}
