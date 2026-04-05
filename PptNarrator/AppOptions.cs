namespace PptNarrator;

class AppOptions
{
    public required string InputFile { get; set; }
    public string Mode { get; set; } = "audio";                   // "audio" or "avatar"
    public string Auth { get; set; } = "entraid";                 // "apikey" or "entraid"
    public string Voice { get; set; } = "en-US-JennyNeural";
    public string Language { get; set; } = "en-US";
    public string? Region { get; set; }
    public string? ApiKey { get; set; }
    public string? ResourceName { get; set; }
    public string AvatarCharacter { get; set; } = "lisa";
    public string AvatarStyle { get; set; } = "graceful-sitting";
    public string? OutputFile { get; set; }
}
