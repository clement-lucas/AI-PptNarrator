namespace PptNarrator;

class Program
{
    static async Task<int> Main(string[] args)
    {
        var options = ParseArguments(args);
        if (options is null) return 1;

        // ── 1. Extract speaker notes ──
        Console.WriteLine($"Reading speaker notes from: {Path.GetFileName(options.InputFile)}");
        var slideNotes = NoteExtractor.Extract(options.InputFile);

        int slidesWithNotes = slideNotes.Count(n => !string.IsNullOrWhiteSpace(n.Text));
        Console.WriteLine($"Found {slidesWithNotes} slide(s) with speaker notes (out of {slideNotes.Count} total).");

        if (slidesWithNotes == 0)
        {
            Console.WriteLine("No speaker notes found. Nothing to do.");
            return 0;
        }

        // Show preview of notes
        foreach (var sn in slideNotes.Where(n => !string.IsNullOrWhiteSpace(n.Text)))
        {
            string preview = sn.Text.Length > 60 ? sn.Text[..60] + "..." : sn.Text;
            preview = preview.ReplaceLineEndings(" ");
            Console.WriteLine($"  Slide {sn.SlideNumber}: \"{preview}\"");
        }
        Console.WriteLine();

        // ── 2. Generate audio/video for each slide with notes ──
        string tempDir = Path.Combine(Path.GetTempPath(), $"PptNarrator_{Guid.NewGuid():N}"[..24]);
        Directory.CreateDirectory(tempDir);

        var mediaFiles = new Dictionary<int, string>(); // 0-based slide index → media path

        using var speechService = new SpeechService(options);

        try
        {
            int processed = 0;
            foreach (var sn in slideNotes)
            {
                if (string.IsNullOrWhiteSpace(sn.Text)) continue;

                processed++;
                string ssml = SsmlBuilder.Build(sn.Text, options.Voice, options.Language);
                string ext = options.Mode == "avatar" ? ".mp4" : ".wav";
                string mediaPath = Path.Combine(tempDir, $"slide{sn.SlideNumber}{ext}");

                string modeLabel = options.Mode == "avatar" ? "avatar video" : "audio";
                Console.Write($"  [{processed}/{slidesWithNotes}] Slide {sn.SlideNumber}: Generating {modeLabel}...");

                if (options.Mode == "avatar")
                    await speechService.GenerateAvatarAsync(ssml, mediaPath, sn.SlideNumber);
                else
                    await speechService.GenerateAudioAsync(ssml, mediaPath);

                Console.WriteLine(" done.");
                mediaFiles[sn.SlideNumber - 1] = mediaPath; // 0-based index
            }

            // ── 3. Create a narrated copy of the PPTX ──
            string outputDir = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "output");
            Directory.CreateDirectory(outputDir);

            string? outputFile = options.OutputFile ??
                Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(options.InputFile) + "-narrated.pptx");

            outputFile = PromptIfFileExists(outputFile);
            if (outputFile is null) { Console.WriteLine("Cancelled."); return 0; }

            Console.WriteLine();
            Console.WriteLine($"Embedding media into: {Path.GetFileName(outputFile)}");

            File.Copy(options.InputFile, outputFile, overwrite: true);
            SlideMediaEmbedder.Embed(outputFile, mediaFiles, options.Mode);

            Console.WriteLine($"Done! Output: {outputFile}");
            return 0;
        }
        finally
        {
            try { Directory.Delete(tempDir, true); } catch { /* cleanup best-effort */ }
        }
    }

    // ─────────────────────────────────────────────────────────────────
    //  DUPLICATE OUTPUT FILE PROMPT
    // ─────────────────────────────────────────────────────────────────

    static string? PromptIfFileExists(string outputFile)
    {
        if (!File.Exists(outputFile)) return outputFile;

        string dir  = Path.GetDirectoryName(outputFile)!;
        string name = Path.GetFileNameWithoutExtension(outputFile);
        string ext  = Path.GetExtension(outputFile);

        Console.WriteLine();
        Console.WriteLine($"File already exists: {Path.GetFileName(outputFile)}");
        Console.WriteLine("  [O] Overwrite");
        Console.WriteLine("  [R] Rename (enter new name)");
        Console.WriteLine("  [C] Cancel");
        Console.Write("Choice [O/R/C]: ");

        string? choice = Console.ReadLine()?.Trim().ToUpperInvariant();

        switch (choice)
        {
            case "O":
                return outputFile;

            case "R":
                Console.Write("New filename (without path): ");
                string? newName = Console.ReadLine()?.Trim();
                if (string.IsNullOrWhiteSpace(newName)) return null;
                if (!newName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                    newName += ".pptx";
                string newPath = Path.Combine(dir, newName);
                // If the renamed file also exists, recurse
                return PromptIfFileExists(newPath);

            default:
                return null;
        }
    }

    // ─────────────────────────────────────────────────────────────────
    //  ARGUMENT PARSING
    // ─────────────────────────────────────────────────────────────────

    static AppOptions? ParseArguments(string[] args)
    {
        if (args.Length == 0 || args[0] is "-h" or "--help")
        {
            PrintUsage();
            return null;
        }

        string inputFile = args[0];
        if (!File.Exists(inputFile))
        {
            Console.Error.WriteLine($"Error: File not found: {inputFile}");
            return null;
        }

        var opts = new AppOptions { InputFile = Path.GetFullPath(inputFile) };

        for (int i = 1; i < args.Length; i++)
        {
            string key = args[i].ToLowerInvariant();
            string? value = i + 1 < args.Length ? args[i + 1] : null;

            switch (key)
            {
                case "--mode":
                    opts.Mode = value ?? "audio"; i++; break;
                case "--auth":
                    opts.Auth = value ?? "entraid"; i++; break;
                case "--voice":
                    opts.Voice = value ?? opts.Voice; i++; break;
                case "--lang":
                    opts.Language = value ?? opts.Language; i++; break;
                case "--region":
                    opts.Region = value; i++; break;
                case "--api-key":
                    opts.ApiKey = value; i++; break;
                case "--resource-name":
                    opts.ResourceName = value; i++; break;
                case "--avatar-character":
                    opts.AvatarCharacter = value ?? opts.AvatarCharacter; i++; break;
                case "--avatar-style":
                    opts.AvatarStyle = value ?? opts.AvatarStyle; i++; break;
                case "--output":
                    opts.OutputFile = value; i++; break;
                default:
                    Console.Error.WriteLine($"Unknown option: {key}");
                    PrintUsage();
                    return null;
            }
        }

        // Validate
        if (opts.Auth == "apikey")
        {
            if (string.IsNullOrEmpty(opts.Region))
            {
                Console.Error.WriteLine("Error: --region is required when using --auth apikey");
                return null;
            }
            if (string.IsNullOrEmpty(opts.ApiKey))
            {
                Console.Error.WriteLine("Error: --api-key is required when using --auth apikey");
                return null;
            }
        }
        else if (opts.Auth == "entraid")
        {
            if (string.IsNullOrEmpty(opts.ResourceName))
            {
                Console.Error.WriteLine("Error: --resource-name is required when using --auth entraid");
                return null;
            }
        }
        else
        {
            Console.Error.WriteLine($"Error: Unknown auth method: {opts.Auth}. Use 'apikey' or 'entraid'.");
            return null;
        }

        if (opts.Mode is not "audio" and not "avatar")
        {
            Console.Error.WriteLine($"Error: Unknown mode: {opts.Mode}. Use 'audio' or 'avatar'.");
            return null;
        }

        return opts;
    }

    static void PrintUsage()
    {
        Console.WriteLine("""
            PptNarrator — Generate audio/avatar narration from PowerPoint speaker notes

            Usage:
              PptNarrator <input.pptx> [options]

            Options:
              --mode <audio|avatar>         Output mode (default: audio)
              --auth <apikey|entraid>       Authentication method (default: entraid)
              --voice <name>               TTS voice name (default: en-US-JennyNeural)
              --lang <code>                Language code (default: en-US)
              --region <region>            Azure region (required for apikey auth)
              --api-key <key>              Speech API key (required for apikey auth)
              --resource-name <name>       Speech resource name (required for entraid auth)
              --avatar-character <name>    Avatar character (default: lisa)
              --avatar-style <style>       Avatar style (default: graceful-sitting)
              --output <path>              Output file path (default: <input>-narrated.pptx)

            Examples:
              # Audio narration with Entra ID auth
              PptNarrator presentation.pptx --auth entraid --resource-name my-speech

              # Audio narration with API key
              PptNarrator presentation.pptx --auth apikey --region swedencentral --api-key abc123

              # Avatar video with Entra ID auth
              PptNarrator presentation.pptx --mode avatar --auth entraid --resource-name my-speech

              # Japanese voice
              PptNarrator slides.pptx --auth entraid --resource-name my-speech --voice ja-JP-NanamiNeural --lang ja-JP
            """);
    }
}
