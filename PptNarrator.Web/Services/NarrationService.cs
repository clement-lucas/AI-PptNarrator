using PptNarrator;

namespace PptNarrator.Web.Services;

public sealed class NarrationJob
{
    public string Id { get; } = Guid.NewGuid().ToString("N")[..12];
    public string FileName { get; set; } = "";
    public string Status { get; set; } = "Pending";
    public string CurrentStep { get; set; } = "";
    public int TotalSlides { get; set; }
    public int ProcessedSlides { get; set; }
    public string? OutputPath { get; set; }
    public string? ErrorMessage { get; set; }
    public List<string> Log { get; } = [];
}

public class NarrationService
{
    private readonly string _workDir = Path.Combine(Path.GetTempPath(), "PptNarratorWeb");

    public NarrationService()
    {
        Directory.CreateDirectory(_workDir);
    }

    public async Task ProcessAsync(
        Stream pptxStream,
        string originalFileName,
        AppOptions options,
        NarrationJob job,
        CancellationToken ct = default)
    {
        string jobDir = Path.Combine(_workDir, job.Id);
        Directory.CreateDirectory(jobDir);

        string inputPath = Path.Combine(jobDir, originalFileName);
        string outputPath = Path.Combine(jobDir,
            Path.GetFileNameWithoutExtension(originalFileName) + "-narrated.pptx");

        try
        {
            // Save uploaded file
            job.Status = "Processing";
            job.CurrentStep = "Saving uploaded file...";
            await using (var fs = File.Create(inputPath))
                await pptxStream.CopyToAsync(fs, ct);

            options.InputFile = inputPath;
            options.OutputFile = outputPath;

            // Extract notes
            job.CurrentStep = "Extracting speaker notes...";
            var slideNotes = NoteExtractor.Extract(inputPath);
            var notesWithText = slideNotes.Where(n => !string.IsNullOrWhiteSpace(n.Text)).ToList();
            job.TotalSlides = notesWithText.Count;

            if (notesWithText.Count == 0)
            {
                job.Status = "Completed";
                job.CurrentStep = "No speaker notes found — nothing to narrate.";
                job.ErrorMessage = "The uploaded PPTX has no speaker notes on any slide.";
                return;
            }

            job.Log.Add($"Found {notesWithText.Count} slide(s) with speaker notes.");

            // Generate media for each slide
            string tempMediaDir = Path.Combine(jobDir, "media");
            Directory.CreateDirectory(tempMediaDir);

            var mediaFiles = new Dictionary<int, string>();
            using var speechService = new SpeechService(options);

            foreach (var sn in notesWithText)
            {
                ct.ThrowIfCancellationRequested();

                job.ProcessedSlides++;
                string modeLabel = options.Mode == "avatar" ? "avatar video" : "audio";
                job.CurrentStep = $"Generating {modeLabel} for slide {sn.SlideNumber}...";
                job.Log.Add($"Slide {sn.SlideNumber}: generating {modeLabel}...");

                string ssml = SsmlBuilder.Build(sn.Text, options.Voice, options.Language);
                string ext = options.Mode == "avatar" ? ".mp4" : ".wav";
                string mediaPath = Path.Combine(tempMediaDir, $"slide{sn.SlideNumber}{ext}");

                if (options.Mode == "avatar")
                    await speechService.GenerateAvatarAsync(ssml, mediaPath, sn.SlideNumber);
                else
                    await speechService.GenerateAudioAsync(ssml, mediaPath);

                mediaFiles[sn.SlideNumber - 1] = mediaPath;
                job.Log.Add($"Slide {sn.SlideNumber}: done.");
            }

            // Embed media
            ct.ThrowIfCancellationRequested();
            job.CurrentStep = "Embedding media into PPTX...";
            job.Log.Add("Embedding media into presentation...");

            File.Copy(inputPath, outputPath, overwrite: true);
            SlideMediaEmbedder.Embed(outputPath, mediaFiles, options.Mode);

            job.OutputPath = outputPath;
            job.FileName = Path.GetFileName(outputPath);
            job.Status = "Completed";
            job.CurrentStep = "Done!";
            job.Log.Add("Narration complete.");
        }
        catch (OperationCanceledException)
        {
            job.Status = "Cancelled";
            job.CurrentStep = "Cancelled by user.";
        }
        catch (Exception ex)
        {
            job.Status = "Failed";
            job.CurrentStep = "Error occurred.";
            job.ErrorMessage = ex.Message;
            job.Log.Add($"Error: {ex.Message}");
        }
    }

    public FileStream? GetOutputFile(NarrationJob job)
    {
        if (job.OutputPath is null || !File.Exists(job.OutputPath))
            return null;
        return File.OpenRead(job.OutputPath);
    }

    public void CleanupJob(NarrationJob job)
    {
        string jobDir = Path.Combine(_workDir, job.Id);
        try { if (Directory.Exists(jobDir)) Directory.Delete(jobDir, true); } catch { }
    }
}
