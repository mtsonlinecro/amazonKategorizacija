using System.Data;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace AmazonCategoryMapperWinForms;

public class MainForm : Form
{
    private TextBox _csvPath = null!;
    private TextBox _mappingPath = null!;
    private TextBox _categoryPath = null!;
    private TextBox _btgFolderPath = null!;
    private TextBox _apiKey = null!;
    private TextBox _model = null!;
    private Button _startBackendButton = null!;
    private Button _mapButton = null!;
    private Button _saveCsvButton = null!;
    private Button _saveCorrectionsButton = null!;
    private Button _importCatalogButton = null!;
    private Button _importFolderButton = null!;
    private ComboBox _catalogMarketplace = null!;
    private Label _status = null!;
    private DataGridView _grid = null!;
    private CheckBox _useAi = null!;
    private CheckBox _overwrite = null!;
    private NumericUpDown _maxAiRows = null!;
    private readonly Dictionary<string, CheckBox> _marketChecks = new();
    private DataTable? _currentTable;
    private byte[]? _lastCsvBytes;

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
        Text = "Amazon Category Mapper - CSV + Browse Node Mapping";
        Width = 1300;
        Height = 800;
        StartPosition = FormStartPosition.CenterScreen;

        _csvPath = new TextBox { Width = 520 };
        _mappingPath = new TextBox { Width = 520 };
        _categoryPath = new TextBox { Width = 520 };
        _btgFolderPath = new TextBox { Width = 520 };
        _apiKey = new TextBox { Width = 300, UseSystemPasswordChar = true };
        _model = new TextBox { Width = 150, Text = "gpt-4.1-mini" };
        _startBackendButton = new Button { Text = "1. Pokreni Python backend", Width = 190 };
        _mapButton = new Button { Text = "3. Mapiraj CSV", Width = 130 };
        _saveCsvButton = new Button { Text = "Preuzmi / spremi novi CSV", Width = 190 };
        _saveCorrectionsButton = new Button { Text = "Spremi ispravke u learning bazu", Width = 220 };
        _importCatalogButton = new Button { Text = "Uvezi 1 BTG u bazu", Width = 160 };
        _importFolderButton = new Button { Text = "Uvezi cijeli BTG folder", Width = 190 };
        _catalogMarketplace = new ComboBox { Width = 90, DropDownStyle = ComboBoxStyle.DropDownList };
        _catalogMarketplace.Items.AddRange(new object[] { "AUTO", "FR", "DE", "IT", "ES", "UK", "NL", "PL", "IE", "SE" });
        _catalogMarketplace.SelectedIndex = 0;
        _status = new Label { AutoSize = true, Text = "Status: spremno" };
        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = false
        };
        _useAi = new CheckBox { Text = "Koristi OpenAI fallback za NL/PL/IE/SE", AutoSize = true };
        _overwrite = new CheckBox { Text = "Prepiši postojeće vrijednosti", Checked = true, AutoSize = true };
        _maxAiRows = new NumericUpDown { Minimum = 1, Maximum = 10000, Value = 50, Width = 80 };

        var top = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            Padding = new Padding(10)
        };

        top.Controls.Add(Row(new Label { Text = "Ulazni CSV/XLSX:", Width = 140 }, _csvPath, BrowseButton(_csvPath, "CSV/XLSX|*.csv;*.xlsx;*.xlsm;*.xls|All files|*.*")));
        top.Controls.Add(Row(new Label { Text = "Amazon MAPPINGS Excel:", Width = 140 }, _mappingPath, BrowseButton(_mappingPath, "Excel|*.xlsx;*.xlsm;*.xls|CSV|*.csv|All files|*.*")));
        top.Controls.Add(Row(
            new Label { Text = "Category catalog/BTG:", Width = 140 },
            _categoryPath,
            BrowseButton(_categoryPath, "CSV/Excel|*.csv;*.xlsx;*.xlsm;*.xls|All files|*.*"),
            new Label { Text = "Marketplace:" },
            _catalogMarketplace,
            _importCatalogButton,
            new Label { Text = "opcionalno; jednom uvezi BTG pa ostaje u lokalnoj bazi" }
        ));
        top.Controls.Add(Row(
            new Label { Text = "BTG folder:", Width = 140 },
            _btgFolderPath,
            BrowseFolderButton(_btgFolderPath),
            _importFolderButton,
            new Label { Text = "ubaci folder koji sadrži FR/IT/ES/DE BTG datoteke; može imati podfoldere FR, IT, ES..." }
        ));

        var marketPanel = new FlowLayoutPanel { AutoSize = true };
        foreach (var mp in new[] { "FR", "IT", "ES", "NL", "PL", "IE", "SE" })
        {
            var isDefault = mp == "FR" || mp == "IT" || mp == "ES";
            var cb = new CheckBox { Text = mp, AutoSize = true, Checked = isDefault };
            _marketChecks[mp] = cb;
            marketPanel.Controls.Add(cb);
        }
        top.Controls.Add(Row(new Label { Text = "Države:", Width = 140 }, marketPanel));

        top.Controls.Add(Row(new Label { Text = "OpenAI API key:", Width = 140 }, _apiKey, new Label { Text = "Model:" }, _model, _useAi, new Label { Text = "Max AI redova:" }, _maxAiRows));
        top.Controls.Add(Row(_overwrite, _startBackendButton, _mapButton, _saveCsvButton, _saveCorrectionsButton, _status));

        Controls.Add(_grid);
        Controls.Add(top);

        _startBackendButton.Click += async (_, _) => await StartBackendClicked();
        _mapButton.Click += async (_, _) => await MapClicked();
        _saveCsvButton.Click += (_, _) => SaveCsvClicked();
        _saveCorrectionsButton.Click += async (_, _) => await SaveCorrectionsClicked();
        _importCatalogButton.Click += async (_, _) => await ImportCatalogClicked();
        _importFolderButton.Click += async (_, _) => await ImportFolderClicked();
    }

    private static FlowLayoutPanel Row(params Control[] controls)
    {
        var panel = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0, 3, 0, 3) };
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

    private static Button BrowseFolderButton(TextBox target)
    {
        var btn = new Button { Text = "Odaberi folder", Width = 110 };
        btn.Click += (_, _) =>
        {
            using var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK) target.Text = dialog.SelectedPath;
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
            BackendRunner.WriteEnvIfNeeded(backendDir, _apiKey.Text, _model.Text);
            BackendRunner.StartBackend(backendDir);
            _status.Text = "Status: backend se pokreće...";
            await Task.Delay(3000);
            _status.Text = await BackendRunner.IsBackendHealthyAsync() ? "Status: backend radi" : "Status: backend se još pokreće ili Python nije instaliran";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Greška kod pokretanja backend-a");
        }
    }

    private async Task ImportCatalogClicked()
    {
        try
        {
            if (!File.Exists(_categoryPath.Text))
            {
                MessageBox.Show("Odaberi BTG/category catalog Excel ili CSV datoteku.");
                return;
            }
            if (!await BackendRunner.IsBackendHealthyAsync())
            {
                MessageBox.Show("Python backend ne radi. Klikni 'Pokreni Python backend' ili ručno pokreni backend/run_backend.bat.");
                return;
            }

            _status.Text = "Status: uvoz BTG/category catalog datoteke u tijeku...";
            using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(30) };
            using var form = new MultipartFormDataContent();
            form.Add(FileContent(_categoryPath.Text, "category_file"));
            form.Add(new StringContent(Convert.ToString(_catalogMarketplace.SelectedItem ?? "AUTO")), "marketplace");

            var response = await client.PostAsync("http://127.0.0.1:8008/api/import-category-catalog", form);
            var text = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                MessageBox.Show(text, "Greška kod uvoza catalog/BTG datoteke");
                _status.Text = "Status: greška kod uvoza catalog-a";
                return;
            }
            _status.Text = "Status: catalog/BTG uvezen: " + text;
            MessageBox.Show(text, "BTG/category catalog uvezen");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Greška kod uvoza catalog/BTG datoteke");
            _status.Text = "Status: greška kod uvoza catalog-a";
        }
    }

    private async Task ImportFolderClicked()
    {
        try
        {
            if (!Directory.Exists(_btgFolderPath.Text))
            {
                MessageBox.Show("Odaberi folder koji sadrži BTG/category catalog datoteke.");
                return;
            }
            if (!await BackendRunner.IsBackendHealthyAsync())
            {
                MessageBox.Show("Python backend ne radi. Klikni 'Pokreni Python backend' ili ručno pokreni backend/run_backend.bat.");
                return;
            }

            _status.Text = "Status: uvoz cijelog BTG foldera u tijeku...";
            using var client = new HttpClient { Timeout = TimeSpan.FromHours(2) };
            var payload = new
            {
                folder_path = _btgFolderPath.Text,
                marketplace = Convert.ToString(_catalogMarketplace.SelectedItem ?? "AUTO")
            };
            var json = JsonSerializer.Serialize(payload);
            var response = await client.PostAsync(
                "http://127.0.0.1:8008/api/import-category-folder",
                new StringContent(json, Encoding.UTF8, "application/json")
            );
            var text = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                MessageBox.Show(text, "Greška kod uvoza BTG foldera");
                _status.Text = "Status: greška kod uvoza foldera";
                return;
            }
            _status.Text = "Status: BTG folder uvezen: " + text;
            MessageBox.Show(text, "BTG folder uvezen");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Greška kod uvoza BTG foldera");
            _status.Text = "Status: greška kod uvoza foldera";
        }
    }

    private async Task MapClicked()
    {
        try
        {
            if (!File.Exists(_csvPath.Text)) { MessageBox.Show("Odaberi ulazni CSV/XLSX."); return; }
            if (!File.Exists(_mappingPath.Text)) { MessageBox.Show("Odaberi Amazon European Browse Node Mapping Excel s MAPPINGS tabom."); return; }

            if (!await BackendRunner.IsBackendHealthyAsync())
            {
                MessageBox.Show("Python backend ne radi. Klikni 'Pokreni Python backend' ili ručno pokreni backend/run_backend.bat.");
                return;
            }

            var markets = string.Join(",", _marketChecks.Where(kv => kv.Value.Checked).Select(kv => kv.Key));
            if (string.IsNullOrWhiteSpace(markets)) { MessageBox.Show("Odaberi barem jednu državu."); return; }

            _status.Text = "Status: mapiranje u tijeku...";
            using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(30) };
            using var form = new MultipartFormDataContent();

            form.Add(FileContent(_csvPath.Text, "input_file"));
            form.Add(FileContent(_mappingPath.Text, "mapping_file"));
            if (File.Exists(_categoryPath.Text))
            {
                form.Add(FileContent(_categoryPath.Text, "category_file"));
            }
            if (Directory.Exists(_btgFolderPath.Text))
            {
                // Backend ce automatski uvesti/azurirati BTG folder prije mapiranja,
                // pa korisnik ne mora svaki put posebno kliknuti uvoz ako je folder odabran.
                form.Add(new StringContent(_btgFolderPath.Text), "btg_folder_path");
                form.Add(new StringContent(Convert.ToString(_catalogMarketplace.SelectedItem ?? "AUTO")), "btg_marketplace");
            }
            form.Add(new StringContent(markets), "marketplaces");
            form.Add(new StringContent(_useAi.Checked ? "true" : "false"), "use_ai");
            form.Add(new StringContent(_overwrite.Checked ? "true" : "false"), "overwrite_existing");
            form.Add(new StringContent(Convert.ToString(_maxAiRows.Value)), "max_ai_rows");

            var response = await client.PostAsync("http://127.0.0.1:8008/api/map-csv", form);
            var bytes = await response.Content.ReadAsByteArrayAsync();
            if (!response.IsSuccessStatusCode)
            {
                var err = Encoding.UTF8.GetString(bytes);
                MessageBox.Show(err, "Greška iz backend-a");
                _status.Text = "Status: greška";
                return;
            }

            _lastCsvBytes = bytes;
            var tempPath = Path.Combine(Path.GetTempPath(), "amazon_categories_mapped.csv");
            File.WriteAllBytes(tempPath, bytes);
            _currentTable = CsvUtil.ReadCsv(tempPath);
            _grid.DataSource = _currentTable;
            _status.Text = $"Status: gotovo. Redova: {_currentTable.Rows.Count}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Greška kod mapiranja");
            _status.Text = "Status: greška";
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
            MessageBox.Show("Prvo mapiraj CSV.");
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
