using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptNarrator;

/// <summary>
/// Extracts speaker notes text from each slide in a PPTX file.
/// </summary>
public static class NoteExtractor
{
    public record SlideNote(int SlideNumber, string Text);

    /// <summary>
    /// Returns a list of (slideNumber, noteText) for every slide in presentation order.
    /// Slides without notes return an empty string.
    /// </summary>
    public static List<SlideNote> Extract(string pptxPath)
    {
        var results = new List<SlideNote>();

        using var doc = PresentationDocument.Open(pptxPath, isEditable: false);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Not a valid PowerPoint file.");

        var slideIdList = presentationPart.Presentation.SlideIdList
            ?? throw new InvalidOperationException("Presentation contains no slides.");

        int slideNumber = 0;
        foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
        {
            slideNumber++;
            var relId = slideId.RelationshipId?.Value;
            if (relId is null) { results.Add(new SlideNote(slideNumber, "")); continue; }

            var slidePart = (SlidePart)presentationPart.GetPartById(relId);
            string noteText = GetNotesText(slidePart);
            results.Add(new SlideNote(slideNumber, noteText));
        }

        return results;
    }

    private static string GetNotesText(SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart;
        if (notesPart is null) return "";

        // Look for the body placeholder shape that contains speaker notes
        var shapes = notesPart.NotesSlide?.CommonSlideData?.ShapeTree?.Elements<Shape>();
        if (shapes is null) return "";

        foreach (var shape in shapes)
        {
            var ph = shape.NonVisualShapeProperties?
                          .ApplicationNonVisualDrawingProperties?
                          .GetFirstChild<PlaceholderShape>();

            // The notes body is type="body" or (no type with idx=1)
            bool isNotesBody = ph is not null &&
                (ph.Type?.Value == PlaceholderValues.Body ||
                 (ph.Type is null && ph.Index?.Value == 1));

            if (!isNotesBody) continue;

            return ExtractTextFromShape(shape);
        }

        return "";
    }

    private static string ExtractTextFromShape(Shape shape)
    {
        var textBody = shape.TextBody;
        if (textBody is null) return "";

        var paragraphs = textBody.Descendants<A.Paragraph>().ToList();
        var sb = new System.Text.StringBuilder();

        for (int i = 0; i < paragraphs.Count; i++)
        {
            if (i > 0) sb.AppendLine();
            foreach (var run in paragraphs[i].Elements<A.Run>())
            {
                sb.Append(run.Text?.Text);
            }
        }

        return sb.ToString().Trim();
    }
}
