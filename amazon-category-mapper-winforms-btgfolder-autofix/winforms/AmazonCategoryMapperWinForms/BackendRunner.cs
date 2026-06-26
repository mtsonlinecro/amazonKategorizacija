using System.Diagnostics;
using System.Net.Http;

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
}
