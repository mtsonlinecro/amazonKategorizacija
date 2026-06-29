using System.Diagnostics;
using System.Net.Http;
using System.Text.Json;

namespace AmazonCategoryMapperWinForms;

public static class BackendRunner
{
    private static Process? _process;

    public static string? FindBackendDirectory()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, "backend");
            if (Directory.Exists(candidate) && File.Exists(Path.Combine(candidate, "run_backend.bat")))
                return candidate;
            dir = dir.Parent;
        }
        return null;
    }

    public static void WriteEnvIfNeeded(string backendDir, string apiKey, string model)
    {
        var envPath = Path.Combine(backendDir, ".env");
        var content = $"OPENAI_API_KEY={apiKey.Trim()}\r\nOPENAI_MODEL={model.Trim()}\r\nDATABASE_URL=sqlite:///data/learning.db\r\n";
        File.WriteAllText(envPath, content);
    }

    public static void StartBackend(string backendDir)
    {
        if (_process != null && !_process.HasExited) return;

        var psi = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = "/c run_backend.bat",
            WorkingDirectory = backendDir,
            UseShellExecute = true,
            CreateNoWindow = false
        };
        _process = Process.Start(psi);
    }

    public static async Task<bool> IsBackendHealthyAsync()
    {
        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };
            var response = await client.GetAsync("http://127.0.0.1:8008/health");
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    public static async Task StopBackendAsync(string? jobId = null)
    {
        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };
            if (!string.IsNullOrWhiteSpace(jobId))
            {
                var payload = JsonSerializer.Serialize(new { job_id = jobId });
                try { await client.PostAsync("http://127.0.0.1:8008/api/cancel", new StringContent(payload, System.Text.Encoding.UTF8, "application/json")); } catch { }
            }
            try { await client.PostAsync("http://127.0.0.1:8008/api/shutdown", new StringContent("{}", System.Text.Encoding.UTF8, "application/json")); } catch { }
        }
        catch
        {
            // Ako HTTP shutdown ne uspije, ispod svejedno gasimo proces koji je aplikacija pokrenula.
        }

        await Task.Delay(500);
        try
        {
            if (_process != null && !_process.HasExited)
            {
                _process.Kill(entireProcessTree: true);
                _process.Dispose();
                _process = null;
            }
        }
        catch
        {
            // Namjerno ignoriramo jer proces možda već više ne postoji.
        }
    }

    public static async Task<string> GetBackendInfoAsync()
    {
        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            var json = await client.GetStringAsync("http://127.0.0.1:8008/health");
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var btg = root.TryGetProperty("btg_folder", out var btgEl) ? btgEl.GetString() ?? "" : "";
            var mapping = root.TryGetProperty("mapping_file", out var mapEl) ? mapEl.GetString() ?? "" : "";

            var btgText = string.IsNullOrWhiteSpace(btg) ? "BTG folder nije pronađen" : $"BTG: {btg}";
            var mappingText = string.IsNullOrWhiteSpace(mapping) ? "mapping_file nije pronađen" : $"Mapping: {mapping}";
            return btgText + " | " + mappingText;
        }
        catch
        {
            return "Ne mogu dohvatiti auto-info s backend-a.";
        }
    }

}
