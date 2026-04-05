using System.Security;

namespace PptNarrator;

/// <summary>
/// Builds SSML markup from plain text and voice settings.
/// </summary>
static class SsmlBuilder
{
    public static string Build(string text, string voice, string language)
    {
        // Escape XML special characters in the text
        string escaped = SecurityElement.Escape(text) ?? text;

        return $"""
            <speak version="1.0" xml:lang="{language}">
              <voice name="{voice}">
                <prosody rate="0%" pitch="0%">
                  {escaped}
                </prosody>
              </voice>
            </speak>
            """;
    }
}
