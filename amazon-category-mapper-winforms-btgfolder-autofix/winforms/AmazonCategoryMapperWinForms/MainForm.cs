using System.Data;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace AmazonCategoryMapperWinForms;

public class MainForm : Form
{
    private TextBox _inputPath = null!;
    private Button _startBackendButton = null!;
    private Button _mapButton = null!;
    private Button _saveCsvButton = null!;
    private Button _saveCorrectionsButton = null!;
    private Button _clearCacheButton = null!;
    private Button _stopBackendButton = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressText = null!;
    private Label _status = null!;
    private Label _autoInfo = null!;
    private DataGridView _grid = null!;
    private CheckBox _overwrite = null!;
    private readonly Dictionary<string, CheckBox> _marketChecks = new();
    private DataTable? _currentTable;
    private byte[]? _lastCsvBytes;
    private CancellationTokenSource? _progressPollingCts;
    private string? _activeJobId;

    public MainForm()
    {
        try
        {
            BuildUi();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Greška u MainForm konstruktoru");
            throw;
        }
    }

    private void BuildUi()
    {
        Text = "Amazon Category Mapper - DE + FR/IT/ES + NL/PL/IE/SE";
        Width = 1300;
        Height = 760;
        StartPosition = FormStartPosition.CenterScreen;

        _inputPath = new TextBox { Width = 620 };
        _startBackendButton = new Button { Text = "1. Pokreni Python backend", Width = 190 };
        _mapButton = new Button { Text = "2. Mapiraj", Width = 120 };
        _saveCsvButton = new Button { Text = "Spremi novi CSV", Width = 150 };
        _saveCorrectionsButton = new Button { Text = "Spremi ručne ispravke", Width = 180 };
        _clearCacheButton = new Button { Text = "Obriši BTG cache", Width = 150 };
        _stopBackendButton = new Button { Text = "Zaustavi Python", Width = 140 };
        _progressBar = new ProgressBar { Width = 520, Height = 22, Minimum = 0, Maximum = 100, Value = 0 };
        _progressText = new Label { AutoSize = true, Text = "Progress: 0%" };
        _status = new Label { AutoSize = true, Text = "Status: spremno" };
        _autoInfo = new Label
        {
            AutoSize = true,
            Text = "Program automatski traži BTG folder i mapping_file.xls/xlsx. DE kategorija ide u izlaz. UK/source stupci su izbačeni. Učitava se samo CSV/XLSX/TXT s DE node ID-evima.",
        };
        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = false
        };
        _overwrite = new CheckBox { Text = "Prepiši postojeće vrijednosti", Checked = true, AutoSize = true };

        var top = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            Padding = new Padding(10)
        };

        top.Controls.Add(Row(
            new Label { Text = "Ulazni CSV/XLSX/TXT:", Width = 150 },
            _inputPath,
            BrowseButton(_inputPath, "CSV/XLSX/TXT|*.csv;*.xlsx;*.xlsm;*.xls;*.txt|All files|*.*"),
            new Label { Text = "TXT može imati samo DE node ID-eve, jedan po retku." }
        ));

        var marketPanel = new FlowLayoutPanel { AutoSize = true };
        foreach (var mp in new[] { "FR", "IT", "ES", "NL", "PL", "IE", "SE" })
        {
            // FR/IT/ES idu preko European mappinga, NL/PL/IE/SE preko BTG similarity prijedloga.
            var cb = new CheckBox { Text = mp, AutoSize = true, Checked = true };
            _marketChecks[mp] = cb;
            marketPanel.Controls.Add(cb);
        }
        top.Controls.Add(Row(new Label { Text = "Države:", Width = 150 }, marketPanel));
        top.Controls.Add(Row(_overwrite, _startBackendButton, _mapButton, _saveCsvButton, _saveCorrectionsButton, _clearCacheButton, _stopBackendButton, _status));
        top.Controls.Add(Row(new Label { Text = "Obrada:", Width = 150 }, _progressBar, _progressText));
        top.Controls.Add(Row(new Label { Text = "Napomena:", Width = 150 }, _autoInfo));

        Controls.Add(_grid);
        Controls.Add(top);

        _startBackendButton.Click += async (_, _) => await StartBackendClicked();
        _mapButton.Click += async (_, _) => await MapClicked();
        _clearCacheButton.Click += async (_, _) => await ClearCacheClicked();
        _stopBackendButton.Click += async (_, _) => await StopBackendClicked();
        FormClosing += (_, _) => StopBackendOnExit();
        _saveCsvButton.Click += (_, _) => SaveCsvClicked();
        _saveCorrectionsButton.Click += async (_, _) => await SaveCorrectionsClicked();
    }

    private static FlowLayoutPanel Row(params Control[] controls)
    {
        var panel = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = new Padding(0, 3, 0, 3)
        };
        foreach (var c in controls)
        {
            c.Margin = new Padding(3);
            panel.Controls.Add(c);
        }
        return panel;
    }

    private static Button BrowseButton(TextBox target, string filter)
    {
        var btn = new Button { Text = "Odaberi", Width = 80 };
        btn.Click += (_, _) =>
        {
            using var dialog = new OpenFileDialog { Filter = filter };
            if (dialog.ShowDialog() == DialogResult.OK) target.Text = dialog.FileName;
        };
        return btn;
    }

    private async Task StartBackendClicked()
    {
        try
        {
            var backendDir = BackendRunner.FindBackendDirectory();
            if (backendDir == null)
            {
                MessageBox.Show("Ne mogu pronaći backend folder. Otvori cijeli folder projekta, ne samo winforms mapu. Alternativno ručno pokreni backend/run_backend.bat.");
                return;
            }

            BackendRunner.StartBackend(backendDir);
            _status.Text = "Status: backend se pokreće...";
            await Task.Delay(3000);

            if (await BackendRunner.IsBackendHealthyAsync())
            {
                _status.Text = "Status: backend radi";
                _autoInfo.Text = await BackendRunner.GetBackendInfoAsync();
            }
            else
            {
                _status.Text = "Status: backend se još pokreće ili Python nije instaliran";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Greška kod pokretanja backend-a");
        }
    }

    private async Task MapClicked()
    {
        string? jobId = null;
        try
        {
            if (!File.Exists(_inputPath.Text))
            {
                MessageBox.Show("Odaberi ulazni CSV/XLSX/TXT. Za TXT je dovoljno staviti DE node ID-eve, jedan po retku.");
                return;
            }

            if (!await BackendRunner.IsBackendHealthyAsync())
            {
                MessageBox.Show("Python backend ne radi. Klikni 'Pokreni Python backend' ili ručno pokreni backend/run_backend.bat.");
                return;
            }

            var markets = string.Join(",", _marketChecks.Where(kv => kv.Value.Checked).Select(kv => kv.Key));
            if (string.IsNullOrWhiteSpace(markets))
            {
                MessageBox.Show("Odaberi barem jednu državu.");
                return;
            }

            jobId = Guid.NewGuid().ToString("N");
            _activeJobId = jobId;
            UpdateProgress(0, "Priprema obrade...");
            _status.Text = "Status: mapiranje u tijeku...";
            _mapButton.Enabled = false;
            _clearCacheButton.Enabled = false;
            _stopBackendButton.Enabled = true;

            _progressPollingCts?.Cancel();
            _progressPollingCts = new CancellationTokenSource();
            var pollingTask = PollProgressAsync(jobId, _progressPollingCts.Token);

            using var client = new HttpClient { Timeout = TimeSpan.FromHours(2) };
            using var form = new MultipartFormDataContent();

            form.Add(FileContent(_inputPath.Text, "input_file"));
            form.Add(new StringContent(markets), "marketplaces");
            form.Add(new StringContent(_overwrite.Checked ? "true" : "false"), "overwrite_existing");
            form.Add(new StringContent(jobId), "job_id");

            var response = await client.PostAsync("http://127.0.0.1:8008/api/map-csv", form);
            var bytes = await response.Content.ReadAsByteArrayAsync();

            _progressPollingCts.Cancel();
            try { await pollingTask; } catch { /* ignore */ }

            if (!response.IsSuccessStatusCode)
            {
                var err = Encoding.UTF8.GetString(bytes);
                MessageBox.Show(err, "Greška iz backend-a");
                _status.Text = err.Contains("zaustavljena", StringComparison.OrdinalIgnoreCase) ? "Status: zaustavljeno" : "Status: greška";
                UpdateProgress(0, "Greška kod obrade.");
                return;
            }

            _lastCsvBytes = bytes;
            var tempPath = Path.Combine(Path.GetTempPath(), "amazon_categories_mapped.csv");
            File.WriteAllBytes(tempPath, bytes);
            _currentTable = CsvUtil.ReadCsv(tempPath);
            _grid.DataSource = _currentTable;
            UpdateProgress(100, "Gotovo.");
            _status.Text = $"Status: gotovo. Redova: {_currentTable.Rows.Count}";
            _autoInfo.Text = await BackendRunner.GetBackendInfoAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Greška kod mapiranja");
            _status.Text = "Status: greška";
            UpdateProgress(0, "Greška kod obrade.");
        }
        finally
        {
            _progressPollingCts?.Cancel();
            _activeJobId = null;
            _mapButton.Enabled = true;
            _clearCacheButton.Enabled = true;
            _stopBackendButton.Enabled = true;
        }
    }


    private void StopBackendOnExit()
    {
        try
        {
            _progressPollingCts?.Cancel();
            var jobId = _activeJobId;
            _activeJobId = null;
            BackendRunner.StopBackendAsync(jobId).GetAwaiter().GetResult();
        }
        catch
        {
            // Kod izlaza ne prikazujemo MessageBox da aplikacija može zatvoriti prozor.
        }
    }

    private async Task StopBackendClicked(bool silent = false)
    {
        try
        {
            _progressPollingCts?.Cancel();
            var jobId = _activeJobId;
            _activeJobId = null;
            if (!silent)
            {
                _status.Text = "Status: zaustavljam Python backend...";
                UpdateProgress(0, "Zaustavljanje Pythona...");
            }
            await BackendRunner.StopBackendAsync(jobId);
            if (!silent)
            {
                _status.Text = "Status: Python backend zaustavljen";
            }
        }
        catch (Exception ex)
        {
            if (!silent) MessageBox.Show(ex.ToString(), "Greška kod zaustavljanja backend-a");
        }
    }

    private async Task PollProgressAsync(string jobId, CancellationToken token)
    {
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        while (!token.IsCancellationRequested)
        {
            try
            {
                var json = await client.GetStringAsync($"http://127.0.0.1:8008/api/progress?job_id={Uri.EscapeDataString(jobId)}", token);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                var percent = root.TryGetProperty("percent", out var p) ? p.GetInt32() : 0;
                var message = root.TryGetProperty("message", out var m) ? (m.GetString() ?? "") : "";
                UpdateProgress(percent, message);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
                // Ako backend trenutno ne odgovori, ne ruši mapiranje. Samo probaj opet.
            }

            try
            {
                await Task.Delay(600, token);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    private void UpdateProgress(int percent, string message)
    {
        percent = Math.Max(0, Math.Min(100, percent));
        if (InvokeRequired)
        {
            BeginInvoke(() => UpdateProgress(percent, message));
            return;
        }
        _progressBar.Value = percent;
        _progressText.Text = string.IsNullOrWhiteSpace(message)
            ? $"Progress: {percent}%"
            : $"Progress: {percent}% - {message}";
    }

    private async Task ClearCacheClicked()
    {
        try
        {
            if (!await BackendRunner.IsBackendHealthyAsync())
            {
                MessageBox.Show("Python backend ne radi. Prvo ga pokreni.");
                return;
            }

            var answer = MessageBox.Show(
                "Ovo briše samo BTG/category cache. Ručne ispravke i learning baza ostaju. Nastaviti?",
                "Obriši BTG cache",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (answer != DialogResult.Yes) return;

            _clearCacheButton.Enabled = false;
            using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
            var response = await client.PostAsync("http://127.0.0.1:8008/api/clear-cache", new StringContent("{}", Encoding.UTF8, "application/json"));
            var text = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                MessageBox.Show(text, "Greška kod brisanja cache-a");
                return;
            }

            _status.Text = "Status: BTG cache obrisan";
            UpdateProgress(0, "Cache je obrisan. Sljedeće mapiranje ponovno učitava BTG datoteke.");
            _autoInfo.Text = await BackendRunner.GetBackendInfoAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Greška kod brisanja cache-a");
        }
        finally
        {
            _clearCacheButton.Enabled = true;
        }
    }

    private static ByteArrayContent FileContent(string path, string fieldName)
    {
        var bytes = File.ReadAllBytes(path);
        var content = new ByteArrayContent(bytes);
        content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
        content.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data")
        {
            Name = $"\"{fieldName}\"",
            FileName = $"\"{Path.GetFileName(path)}\""
        };
        return content;
    }

    private void SaveCsvClicked()
    {
        if (_currentTable == null && _lastCsvBytes == null)
        {
            MessageBox.Show("Prvo mapiraj ulaznu datoteku.");
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV|*.csv",
            FileName = "amazon_categories_mapped.csv"
        };
        if (dialog.ShowDialog() != DialogResult.OK) return;

        _grid.EndEdit();
        if (_currentTable != null) CsvUtil.WriteCsv(_currentTable, dialog.FileName);
        else File.WriteAllBytes(dialog.FileName, _lastCsvBytes!);
        _status.Text = "Status: CSV spremljen";
    }

    private async Task SaveCorrectionsClicked()
    {
        if (_currentTable == null)
        {
            MessageBox.Show("Nema podataka za spremanje.");
            return;
        }
        if (!await BackendRunner.IsBackendHealthyAsync())
        {
            MessageBox.Show("Python backend ne radi.");
            return;
        }

        _grid.EndEdit();
        var rows = CsvUtil.ToRows(_currentTable);
        var markets = _marketChecks.Where(kv => kv.Value.Checked).Select(kv => kv.Key).ToArray();
        var payload = new { rows, marketplaces = markets };
        var json = JsonSerializer.Serialize(payload);

        using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
        var response = await client.PostAsync("http://127.0.0.1:8008/api/save-corrections", new StringContent(json, Encoding.UTF8, "application/json"));
        var text = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            MessageBox.Show(text, "Greška kod spremanja ispravaka");
            return;
        }
        _status.Text = "Status: ispravke spremljene u learning bazu: " + text;
    }
}
