using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;

namespace PptNarrator;

/// <summary>
/// Calls Azure AI Speech REST APIs for TTS audio and avatar video generation.
/// </summary>
sealed class SpeechService : IDisposable
{
    private readonly AppOptions _opts;
    private readonly HttpClient _http = new();
    private string? _cachedToken;

    public SpeechService(AppOptions opts)
    {
        _opts = opts;
    }

    // ───── TTS (synchronous POST → WAV) ─────

    public async Task GenerateAudioAsync(string ssml, string outputPath)
    {
        string uri = _opts.Auth == "apikey"
            ? $"https://{_opts.Region}.tts.speech.microsoft.com/cognitiveservices/v1"
            : $"https://{_opts.ResourceName}.cognitiveservices.azure.com/tts/cognitiveservices/v1";

        using var request = new HttpRequestMessage(HttpMethod.Post, uri);
        request.Content = new StringContent(ssml, Encoding.UTF8, "application/ssml+xml");
        request.Headers.Add("X-Microsoft-OutputFormat", "riff-16khz-16bit-mono-pcm");
        request.Headers.Add("User-Agent", "PptNarrator");
        await AddAuth(request);

        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        await using var fs = File.Create(outputPath);
        await response.Content.CopyToAsync(fs);
    }

    // ───── Avatar (async batch synthesis: PUT → poll → download) ─────

    public async Task GenerateAvatarAsync(string ssml, string outputPath, int slideNumber)
    {
        string apiBase = _opts.Auth == "apikey"
            ? $"https://{_opts.Region}.api.cognitive.microsoft.com"
            : $"https://{_opts.ResourceName}.cognitiveservices.azure.com";

        string synthesisId = $"slide{slideNumber}-{Guid.NewGuid():N}"[..32];
        string uri = $"{apiBase}/avatar/batchsyntheses/{synthesisId}?api-version=2024-08-01";

        // Build request body
        var avatarConfig = new Dictionary<string, object>
        {
            ["talkingAvatarCharacter"] = _opts.AvatarCharacter,
            ["videoFormat"] = "Mp4",
            ["videoCodec"] = "h264",
            ["subtitleType"] = "soft_embedded",
            ["backgroundColor"] = "#FFFFFFFF",
            ["bitrateKbps"] = 2000,
            ["customized"] = false
        };
        if (!string.IsNullOrEmpty(_opts.AvatarStyle))
            avatarConfig["talkingAvatarStyle"] = _opts.AvatarStyle;

        var body = new Dictionary<string, object>
        {
            ["inputKind"] = "SSML",
            ["inputs"] = new[] { new { content = ssml } },
            ["avatarConfig"] = avatarConfig
        };

        string json = JsonSerializer.Serialize(body);

        // PUT to create job
        using var createReq = new HttpRequestMessage(HttpMethod.Put, uri);
        createReq.Content = new StringContent(json, Encoding.UTF8, "application/json");
        await AddAuth(createReq);

        using var createResp = await _http.SendAsync(createReq);
        if (!createResp.IsSuccessStatusCode)
        {
            string err = await createResp.Content.ReadAsStringAsync();
            throw new HttpRequestException($"Avatar job submit failed ({createResp.StatusCode}): {err}");
        }

        // Poll until Succeeded or Failed
        string status = "NotStarted";
        JsonElement result = default;
        while (status is not "Succeeded" and not "Failed")
        {
            await Task.Delay(TimeSpan.FromSeconds(10));

            using var pollReq = new HttpRequestMessage(HttpMethod.Get, uri);
            await AddAuth(pollReq);

            using var pollResp = await _http.SendAsync(pollReq);
            pollResp.EnsureSuccessStatusCode();

            string pollJson = await pollResp.Content.ReadAsStringAsync();
            result = JsonSerializer.Deserialize<JsonElement>(pollJson);
            status = result.GetProperty("status").GetString() ?? "Unknown";
            Console.Write($" [{status}]");
        }

        if (status == "Failed")
            throw new InvalidOperationException($"Avatar synthesis failed for slide {slideNumber}.");

        // Download video
        string videoUrl = result.GetProperty("outputs").GetProperty("result").GetString()
            ?? throw new InvalidOperationException("No result URL in avatar response.");

        using var dlReq = new HttpRequestMessage(HttpMethod.Get, videoUrl);
        using var dlResp = await _http.SendAsync(dlReq, HttpCompletionOption.ResponseHeadersRead);
        dlResp.EnsureSuccessStatusCode();

        await using var fs = File.Create(outputPath);
        await dlResp.Content.CopyToAsync(fs);
    }

    // ───── Authentication ─────

    private async Task AddAuth(HttpRequestMessage request)
    {
        if (_opts.Auth == "apikey")
        {
            request.Headers.Add("Ocp-Apim-Subscription-Key", _opts.ApiKey);
        }
        else
        {
            string token = await GetEntraTokenAsync();
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }
    }

    private async Task<string> GetEntraTokenAsync()
    {
        if (_cachedToken is not null) return _cachedToken;

        var credential = new DefaultAzureCredential();
        var context = new TokenRequestContext(["https://cognitiveservices.azure.com/.default"]);
        var token = await credential.GetTokenAsync(context);
        _cachedToken = token.Token;
        return _cachedToken;
    }

    public void Dispose() => _http.Dispose();
}
