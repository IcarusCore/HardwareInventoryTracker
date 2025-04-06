using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HardwareInventoryTrackerNew.Properties;
using OfficeOpenXml;
using System.Data.SQLite;
using System.Diagnostics;

namespace HardwareInventoryTrackerNew
{
    public static class AppInfo
    {
        public const string CurrentVersion = "1.1.9";
    }

    public static class InventoryFields
    {
        public const string AssetTag = "asset_tag";
        public const string Description = "description";
        public const string SerialNumber = "serial_number";
        public const string TransferSheet = "transfer_sheet";
        public const string Notes = "notes";
        public const string Date = "date";
        public const string Time = "time";
        public const string Location = "location";
        public const string TransferredBy = "transferred_by";
        public const string ReceivedBy = "received_by";
        public const string Color = "color";
        public const string Id = "id";
    }

    public class InventoryAppForm : Form
    {
        private readonly InventoryDatabase db;
        private ComboBox cmbTheme = null!;
        private Button btnHelp = null!, btnUpdateApp = null!;
        private Dictionary<string, Control> entries = new Dictionary<string, Control>();
        private DataGridView dgvInventory = null!;
        private Button btnAddEntry = null!, btnBatchAdd = null!, btnUpdate = null!,
                       btnDelete = null!, btnSearch = null!, btnReset = null!,
                       btnRefresh = null!, btnExportCSV = null!, btnImportCSV = null!,
                       btnUpdateDB = null!;
        private Dictionary<string, Dictionary<string, string>> colorSchemes = new Dictionary<string, Dictionary<string, string>>();
        private string currentScheme = "Dark";
        private int currentPage = 1;
        private const int PageSize = 25;

        public InventoryAppForm()
        {
            Text = "Asset Management";
            Size = new Size(1900, 900);

            try
            {
                this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading icon: " + ex.Message);
            }

            db = new InventoryDatabase();
            db.SetupDatabase();
            SetupColorSchemes();
            RetrieveThemeFromSettings();
            CreateWidgets();
            InitializeTheme();
            LoadData();
        }

        #region Color Schemes & Styling

        private void SetupColorSchemes()
        {
            colorSchemes["Light"] = new Dictionary<string, string>
            {
                {"frame_bg", "#FFFFFF"}, {"label_fg", "#333333"}, {"label_bg", "#FFFFFF"},
                {"tree_bg", "#FAFAFA"}, {"tree_fg", "#333333"}, {"button_fg", "#333333"},
                {"select_bg", "#3498DB"}, {"select_fg", "#FFFFFF"}
            };

            colorSchemes["Dark"] = new Dictionary<string, string>
            {
                {"frame_bg", "#2C3E50"}, {"label_fg", "#ECF0F1"}, {"label_bg", "#2C3E50"},
                {"tree_bg", "#3B4B5A"}, {"tree_fg", "#ECF0F1"}, {"button_fg", "#2C3E50"},
                {"select_bg", "#3498DB"}, {"select_fg", "#FFFFFF"}
            };

            colorSchemes["Midnight"] = new Dictionary<string, string>
            {
                {"frame_bg", "#1B1F22"}, {"label_fg", "#F2F2F2"}, {"label_bg", "#1B1F22"},
                {"tree_bg", "#2A2F33"}, {"tree_fg", "#F2F2F2"}, {"button_fg", "#F2F2F2"},
                {"select_bg", "#8E44AD"}, {"select_fg", "#FFFFFF"}
            };

            colorSchemes["Green"] = new Dictionary<string, string>
            {
                {"frame_bg", "#C8E6C9"}, {"label_fg", "#2E7D32"}, {"label_bg", "#C8E6C9"},
                {"tree_bg", "#A5D6A7"}, {"tree_fg", "#2E7D32"}, {"button_fg", "#2E7D32"},
                {"select_bg", "#388E3C"}, {"select_fg", "#FFFFFF"}
            };

            colorSchemes["Purple"] = new Dictionary<string, string>
            {
                {"frame_bg", "#E1BEE7"}, {"label_fg", "#311B92"}, {"label_bg", "#E1BEE7"},
                {"tree_bg", "#CE93D8"}, {"tree_fg", "#311B92"}, {"button_fg", "#311B92"},
                {"select_bg", "#6A1B9A"}, {"select_fg", "#FFFFFF"}
            };

            colorSchemes["Blue"] = new Dictionary<string, string>
            {
                {"frame_bg", "#35a5ff"}, {"label_fg", "#00007b"}, {"label_bg", "#35a5ff"},
                {"tree_bg", "#90CAF9"}, {"tree_fg", "#00007b"}, {"button_fg", "#00007b"},
                {"select_bg", "#1565C0"}, {"select_fg", "#FFFFFF"}
            };

            colorSchemes["Red"] = new Dictionary<string, string>
            {
                {"frame_bg", "#FFCDD2"}, {"label_fg", "#B71C1C"}, {"label_bg", "#FFCDD2"},
                {"tree_bg", "#EF9A9A"}, {"tree_fg", "#B71C1C"}, {"button_fg", "#B71C1C"},
                {"select_bg", "#D32F2F"}, {"select_fg", "#FFFFFF"}
            };

            colorSchemes["USC Gamecocks"] = new Dictionary<string, string>
            {
                {"frame_bg", "#73000A"}, {"label_fg", "#FFFFFF"}, {"label_bg", "#73000A"},
                {"tree_bg", "#C4C4C4"}, {"tree_fg", "#000000"}, {"button_fg", "#000000"},
                {"select_bg", "#000000"}, {"select_fg", "#FFFFFF"}
            };

            colorSchemes["Clemson"] = new Dictionary<string, string>
            {
                {"frame_bg", "#F56600"}, {"label_fg", "#FFFFFF"}, {"label_bg", "#F56600"},
                {"tree_bg", "#522D80"}, {"tree_fg", "#FFFFFF"}, {"button_fg", "#F56600"},
                {"select_bg", "#522D80"}, {"select_fg", "#FFFFFF"}
            };
        }

        private void RetrieveThemeFromSettings()
        {
            try
            {
                var rows = db.ExecuteQuery("SELECT theme FROM settings WHERE id=1");
                currentScheme = rows.Any() && rows[0][0] != DBNull.Value ? rows[0][0].ToString()! : "Dark";
            }
            catch (Exception ex)
            {
                File.AppendAllText("error.log", $"{DateTime.Now}: Theme retrieval failed - {ex.Message}\n");
                currentScheme = "Dark";
            }
        }

        private void InitializeTheme()
        {
            cmbTheme.SelectedIndexChanged -= CmbTheme_SelectedIndexChanged;
            ApplyColorScheme(currentScheme);
            cmbTheme.SelectedIndexChanged += CmbTheme_SelectedIndexChanged;
        }

        private void ApplyColorScheme(string schemeName)
        {
            if (!colorSchemes.ContainsKey(schemeName))
                schemeName = "Dark";

            var scheme = colorSchemes[schemeName];
            BackColor = ColorTranslator.FromHtml(scheme["frame_bg"]);

            foreach (Control ctrl in Controls)
                ApplyControlColors(ctrl, scheme);

            UpdateButtonStyles(scheme);

            try
            {
                db.ExecuteNonQuery(
                    "INSERT OR REPLACE INTO settings (id, theme) VALUES (1, @theme)",
                    new Dictionary<string, object> { { "@theme", schemeName } }
                );
            }
            catch (Exception ex)
            {
                File.AppendAllText("error.log", $"{DateTime.Now}: Theme save failed - {ex.Message}\n");
            }
        }

        private void ApplyControlColors(Control ctrl, Dictionary<string, string> scheme)
        {
            if (ctrl is Label lbl)
            {
                lbl.Dock = DockStyle.Fill;
                lbl.TextAlign = ContentAlignment.TopRight;
                lbl.BackColor = ColorTranslator.FromHtml(scheme["label_bg"]);
                lbl.ForeColor = ColorTranslator.FromHtml(scheme["label_fg"]);
                lbl.Font = new Font("Segoe UI", 11);
            }
            else if (ctrl is TextBox || ctrl is ComboBox || ctrl is DateTimePicker)
            {
                ctrl.Dock = DockStyle.Fill;
                if (entries.ContainsValue(ctrl))
                {
                    ctrl.BackColor = Color.White;
                    ctrl.ForeColor = Color.Black;
                }
                else
                {
                    if (currentScheme == "Light")
                    {
                        ctrl.BackColor = Color.White;
                        ctrl.ForeColor = Color.Black;
                    }
                    else
                    {
                        ctrl.BackColor = ColorTranslator.FromHtml(scheme["frame_bg"]);
                        ctrl.ForeColor = ColorTranslator.FromHtml(scheme["label_fg"]);
                    }
                }
                ctrl.Font = new Font("Segoe UI", 10);
            }

            foreach (Control child in ctrl.Controls)
                ApplyControlColors(child, scheme);
        }

        private void UpdateButtonStyles(Dictionary<string, string> scheme)
        {
            UpdateButtonStylesRecursive(this, scheme);
        }

        private void UpdateButtonStylesRecursive(Control ctrl, Dictionary<string, string> scheme)
        {
            if (ctrl is Button btn)
            {
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.BackColor = ColorTranslator.FromHtml(scheme["select_bg"]);
                btn.ForeColor = ColorTranslator.FromHtml(scheme["select_fg"]);
                btn.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            }
            foreach (Control child in ctrl.Controls)
                UpdateButtonStylesRecursive(child, scheme);
        }

        #endregion

        #region UI Creation

        private void CreateWidgets()
        {
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 4
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(mainLayout);

            var topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true
            };
            topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            FlowLayoutPanel leftFlow = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            Label lblTheme = new Label
            {
                Text = "Choose Color Theme:",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(5)
            };
            cmbTheme = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 120,
                Margin = new Padding(5)
            };
            cmbTheme.Items.AddRange(colorSchemes.Keys.ToArray());
            cmbTheme.SelectedItem = currentScheme;
            cmbTheme.SelectedIndexChanged += CmbTheme_SelectedIndexChanged;

            btnHelp = new Button
            {
                Text = "Help",
                AutoSize = false,
                Size = new Size(60, 25),
                Margin = new Padding(5)
            };
            btnHelp.Click += BtnHelp_Click;

            leftFlow.Controls.Add(lblTheme);
            leftFlow.Controls.Add(cmbTheme);
            leftFlow.Controls.Add(btnHelp);

            btnUpdateApp = new Button
            {
                Text = "Update App",
                AutoSize = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Margin = new Padding(10)
            };
            btnUpdateApp.Click += BtnUpdateApp_Click;

            topTable.Controls.Add(leftFlow, 0, 0);
            topTable.Controls.Add(btnUpdateApp, 1, 0);
            mainLayout.Controls.Add(topTable, 0, 0);

            TableLayoutPanel formPanel = new TableLayoutPanel
            {
                ColumnCount = 4,
                RowCount = 5,
                Dock = DockStyle.Top,
                Padding = new Padding(10),
                AutoSize = true
            };
            formPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            formPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            formPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            formPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            for (int i = 0; i < 5; i++)
                formPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            string[] fieldLabels = {
                "Asset Tag", "Description", "Serial Number", "Transfer Sheet",
                "Notes", "Date", "Location", "Transferred By", "Received By", "Color"
            };

            for (int i = 0; i < fieldLabels.Length; i++)
            {
                int row = i / 2;
                int colLabel = (i % 2) * 2;
                int colControl = colLabel + 1;

                Label lbl = new Label
                {
                    Text = fieldLabels[i] + ":",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.TopRight,
                    Margin = new Padding(2)
                };

                Control ctrl;
                if (fieldLabels[i] == "Date")
                {
                    FlowLayoutPanel dateTimePanel = new FlowLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        FlowDirection = FlowDirection.LeftToRight,
                        Margin = new Padding(2),
                        AutoSize = true
                    };

                    DateTimePicker dtpDate = new DateTimePicker
                    {
                        Format = DateTimePickerFormat.Custom,
                        CustomFormat = "MM/dd/yyyy",
                        Width = 150,
                        Margin = new Padding(2)
                    };

                    Label lblTime = new Label
                    {
                        Text = "Time:",
                        AutoSize = true,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Margin = new Padding(5, 2, 2, 2)
                    };

                    DateTimePicker dtpTime = new DateTimePicker
                    {
                        Format = DateTimePickerFormat.Custom,
                        CustomFormat = "hh:mm tt",
                        ShowUpDown = true,
                        Width = 120,
                        Margin = new Padding(2),
                        Font = new Font("Segoe UI", 10)
                    };

                    dateTimePanel.Controls.Add(dtpDate);
                    dateTimePanel.Controls.Add(lblTime);
                    dateTimePanel.Controls.Add(dtpTime);
                    ctrl = dateTimePanel;

                    entries[InventoryFields.Date] = dtpDate;
                    entries[InventoryFields.Time] = dtpTime;
                }
                else if (fieldLabels[i] == "Notes")
                {
                    TextBox tb = new TextBox
                    {
                        Dock = DockStyle.Fill,
                        Margin = new Padding(2)
                    };
                    ctrl = tb;
                }
                else if (fieldLabels[i] == "Color")
                {
                    ComboBox cb = new ComboBox
                    {
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Dock = DockStyle.Fill,
                        Margin = new Padding(2)
                    };
                    cb.Items.AddRange(new string[] { "Blue", "Pink" });
                    ctrl = cb;
                }
                else
                {
                    TextBox tb = new TextBox
                    {
                        Dock = DockStyle.Fill,
                        Margin = new Padding(2)
                    };
                    ctrl = tb;
                }

                formPanel.Controls.Add(lbl, colLabel, row);
                formPanel.Controls.Add(ctrl, colControl, row);
                if (fieldLabels[i] != "Date")
                {
                    string key = fieldLabels[i].ToLower().Replace(" ", "_");
                    entries[key] = ctrl;
                }
            }
            mainLayout.Controls.Add(formPanel, 0, 1);

            TableLayoutPanel buttonsPanel = new TableLayoutPanel
            {
                ColumnCount = 10,
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                AutoSize = false,
                Height = 70
            };
            for (int i = 0; i < 10; i++)
                buttonsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));

            string[] btnNames = {
                "Add Entry", "Batch Add", "Update", "Delete", "Search",
                "Reset Form", "Refresh", "Export CSV", "Import CSV", "Update DB"
            };
            for (int i = 0; i < btnNames.Length; i++)
            {
                Button btn = new Button
                {
                    Text = btnNames[i],
                    Dock = DockStyle.Fill,
                    Margin = new Padding(5)
                };
                btn.Click += Button_Click;
                buttonsPanel.Controls.Add(btn, i, 0);

                switch (btnNames[i])
                {
                    case "Add Entry": btnAddEntry = btn; break;
                    case "Batch Add": btnBatchAdd = btn; break;
                    case "Update": btnUpdate = btn; break;
                    case "Delete": btnDelete = btn; break;
                    case "Search": btnSearch = btn; break;
                    case "Reset Form": btnReset = btn; break;
                    case "Refresh": btnRefresh = btn; break;
                    case "Export CSV": btnExportCSV = btn; break;
                    case "Import CSV": btnImportCSV = btn; break;
                    case "Update DB": btnUpdateDB = btn; break;
                }
            }
            mainLayout.Controls.Add(buttonsPanel, 0, 2);

            dgvInventory = new DataGridView
            {
                Dock = DockStyle.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                RowTemplate = { Height = 30 },
                Margin = new Padding(10)
            };
            string[] columns = {
                "ID", "Asset Tag", "Description", "Serial Number",
                "Transfer Sheet", "Notes", "Date", "Time", "Location",
                "Transferred By", "Received By", "Color"
            };
            foreach (string colName in columns)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn
                {
                    Name = colName,
                    HeaderText = colName,
                    Width = colName == "Time" ? 120 : 185,
                    Visible = (colName != "ID")
                };
                dgvInventory.Columns.Add(col);
            }
            dgvInventory.CellClick += DgvInventory_CellClick;
            dgvInventory.SortCompare += DgvInventory_SortCompare;
            mainLayout.Controls.Add(dgvInventory, 0, 3);

            var paginationPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 40,
                Padding = new Padding(10)
            };
            var btnPrev = new Button { Text = "Previous", Width = 80, Margin = new Padding(5, -5, 5, 10) };
            var btnNext = new Button { Text = "Next", Width = 80, Margin = new Padding(5, -5, 5, 10) };
            var spacer = new Label { AutoSize = false, Width = 1100, Margin = new Padding(0) };
            var lblPageInfo = new Label { Name = "lblPageInfo", AutoSize = true, Margin = new Padding(5, 5, 5, 5) };
            btnPrev.Click += (s, e) => { if (currentPage > 1) { currentPage--; LoadData(); UpdatePageInfo(lblPageInfo); } };
            btnNext.Click += (s, e) => { currentPage++; LoadData(); UpdatePageInfo(lblPageInfo); };
            paginationPanel.Controls.Add(btnPrev);
            paginationPanel.Controls.Add(btnNext);
            paginationPanel.Controls.Add(spacer);
            paginationPanel.Controls.Add(lblPageInfo);
            Controls.Add(paginationPanel);
        }

        #endregion

        #region Event Handlers

        private void CmbTheme_SelectedIndexChanged(object? sender, EventArgs e)
        {
            currentScheme = cmbTheme.SelectedItem?.ToString() ?? "Dark";
            ApplyColorScheme(currentScheme);
        }

        private void BtnHelp_Click(object? sender, EventArgs e)
        {
            ShowHelp();
        }

        private void BtnUpdateApp_Click(object? sender, EventArgs e)
        {
            try
            {
                string versionFilePath = @"\\Server\File\Path\latest_version.txt";
                string latestVersion = File.ReadAllText(versionFilePath).Trim();

                Version current = new Version(AppInfo.CurrentVersion);
                Version remote = new Version(latestVersion);

                if (remote <= current)
                {
                    MessageBox.Show("You already have the latest version.",
                                  "No Update Needed",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Information);
                    return;
                }

                DialogResult dr = MessageBox.Show(
                    $"A new version ({remote}) is available. Would you like to update?",
                    "Update Available",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (dr == DialogResult.Yes)
                {
                    string serverExePath = $@"C:\Temp\HardwareInventoryTrackerNew.exe";
                    string localTempPath = Path.Combine(Path.GetTempPath(), $"HardwareInventoryTrackerNew.exe");

                    File.Copy(serverExePath, localTempPath, true);

                    string updaterPath = Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        "MyApp.Updater.exe"
                    );

                    Process.Start(updaterPath, $"\"{localTempPath}\" \"{Application.ExecutablePath}\"");
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while updating:\n" + ex.Message,
                              "Update Error",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error);
            }
        }

        private void Button_Click(object? sender, EventArgs e)
        {
            if (sender is not Button btn)
                return;

            try
            {
                switch (btn.Text)
                {
                    case "Add Entry": AddEntry(); break;
                    case "Batch Add": BatchAdd(); break;
                    case "Update": UpdateEntry(); break;
                    case "Delete": DeleteEntry(); break;
                    case "Search": SearchEntry(); break;
                    case "Reset Form": ClearFields(); break;
                    case "Refresh": LoadData(); break;
                    case "Export CSV": ExportData(); break;
                    case "Import CSV": ImportInventoryCSV(); break;
                    case "Update DB": LoadKnownInventory(); break;
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText("error.log", $"{DateTime.Now}: Button {btn.Text} failed - {ex.Message}\n");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvInventory_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                SelectEntry();
        }

        private void DgvInventory_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if (e.Column.Name == "Date" || e.Column.Name == "Time")
            {
                string date1 = dgvInventory.Rows[e.RowIndex1].Cells["Date"].Value?.ToString() ?? "";
                string time1 = dgvInventory.Rows[e.RowIndex1].Cells["Time"].Value?.ToString() ?? "";
                string date2 = dgvInventory.Rows[e.RowIndex2].Cells["Date"].Value?.ToString() ?? "";
                string time2 = dgvInventory.Rows[e.RowIndex2].Cells["Time"].Value?.ToString() ?? "";

                DateTime dt1, dt2;
                bool parsed1 = DateTime.TryParseExact($"{date1} {time1}", "MM/dd/yyyy hh:mm tt", null, System.Globalization.DateTimeStyles.None, out dt1);
                bool parsed2 = DateTime.TryParseExact($"{date2} {time2}", "MM/dd/yyyy hh:mm tt", null, System.Globalization.DateTimeStyles.None, out dt2);

                if (!parsed1) dt1 = DateTime.MinValue;
                if (!parsed2) dt2 = DateTime.MinValue;

                e.SortResult = dt1.CompareTo(dt2); // Newest first (descending)
                e.Handled = true;
            }
        }

        #endregion

        #region CRUD and Batch Methods

        private void AddEntry() { ProcessEntries(batch: false); }
        private void BatchAdd() { ProcessEntries(batch: true); }

        private void ProcessEntries(bool batch)
        {
            Dictionary<string, string> raw = GetFormData();
            if (!ValidateEntry(raw, out string error))
            {
                MessageBox.Show(error, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string[] assetTags = raw[InventoryFields.AssetTag].Split(',', StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
            string[] descriptions = raw[InventoryFields.Description].Split(',', StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
            string[] serials = raw[InventoryFields.SerialNumber].Split(',', StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
            string[] sheets = raw[InventoryFields.TransferSheet].Split(',', StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();

            if (assetTags.Length == 0) assetTags = new string[] { "" };
            if (descriptions.Length == 0) descriptions = new string[] { "" };
            if (serials.Length == 0) serials = new string[] { "" };
            if (sheets.Length == 0) sheets = new string[] { "" };

            int maxEntries = new[] { assetTags.Length, descriptions.Length, serials.Length, sheets.Length }.Max();

            if (assetTags.Length == 1 && maxEntries > 1) assetTags = Enumerable.Repeat(assetTags[0], maxEntries).ToArray();
            if (descriptions.Length == 1 && maxEntries > 1) descriptions = Enumerable.Repeat(descriptions[0], maxEntries).ToArray();
            if (serials.Length == 1 && maxEntries > 1) serials = Enumerable.Repeat(serials[0], maxEntries).ToArray();
            if (sheets.Length == 1 && maxEntries > 1) sheets = Enumerable.Repeat(sheets[0], maxEntries).ToArray();

            var entriesList = new List<Dictionary<string, string>>();
            for (int i = 0; i < maxEntries; i++)
            {
                var entry = new Dictionary<string, string>
                {
                    [InventoryFields.AssetTag] = i < assetTags.Length ? assetTags[i] : "",
                    [InventoryFields.Description] = i < descriptions.Length ? descriptions[i] : "",
                    [InventoryFields.SerialNumber] = i < serials.Length ? serials[i] : "",
                    [InventoryFields.TransferSheet] = i < sheets.Length ? sheets[i] : "",
                    [InventoryFields.Notes] = raw[InventoryFields.Notes],
                    [InventoryFields.Date] = raw[InventoryFields.Date],
                    [InventoryFields.Time] = raw[InventoryFields.Time],
                    [InventoryFields.Location] = raw[InventoryFields.Location],
                    [InventoryFields.TransferredBy] = raw[InventoryFields.TransferredBy],
                    [InventoryFields.ReceivedBy] = raw[InventoryFields.ReceivedBy],
                    [InventoryFields.Color] = raw[InventoryFields.Color]
                };
                entriesList.Add(entry);
            }

            try
            {
                if (batch && entriesList.Count > 1)
                {
                    db.BulkInsert(entriesList);
                }
                else
                {
                    foreach (var entry in entriesList)
                    {
                        db.ExecuteNonQuery(
                            $"INSERT INTO inventory ({string.Join(", ", entry.Keys)}) VALUES ({string.Join(", ", entry.Keys.Select(k => $"@{k}"))})",
                            entry.ToDictionary(kvp => $"@{kvp.Key}", kvp => (object)kvp.Value)
                        );
                    }
                }
                LoadData();
                MessageBox.Show($"{entriesList.Count} entries added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding entries: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Dictionary<string, string> GetFormData()
        {
            var raw = new Dictionary<string, string>();
            foreach (var pair in entries)
            {
                if (pair.Value is DateTimePicker dtp)
                {
                    if (pair.Key == InventoryFields.Date)
                        raw[pair.Key] = dtp.Value.ToString("MM/dd/yyyy");
                    else if (pair.Key == InventoryFields.Time)
                        raw[pair.Key] = dtp.Value.ToString("hh:mm tt");
                }
                else if (pair.Value is TextBox tb)
                    raw[pair.Key] = tb.Text;
                else if (pair.Value is ComboBox cb)
                    raw[pair.Key] = cb.Text;
            }
            return raw;
        }

        private bool ValidateEntry(Dictionary<string, string> raw, out string error)
        {
            if (string.IsNullOrWhiteSpace(raw[InventoryFields.SerialNumber]))
            {
                error = "Serial Number is required.";
                return false;
            }
            if (!DateTime.TryParseExact(raw[InventoryFields.Date], "MM/dd/yyyy", null, System.Globalization.DateTimeStyles.None, out _))
            {
                error = "Invalid Date format.";
                return false;
            }
            if (!DateTime.TryParseExact(raw[InventoryFields.Time], "hh:mm tt", null, System.Globalization.DateTimeStyles.None, out _))
            {
                error = "Invalid Time format.";
                return false;
            }
            error = null;
            return true;
        }

        private void UpdateEntry()
        {
            if (dgvInventory.SelectedRows.Count == 0)
            {
                MessageBox.Show("No entry selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int id = Convert.ToInt32(dgvInventory.SelectedRows[0].Cells["ID"].Value);
            var raw = GetFormData();
            if (!ValidateEntry(raw, out string error))
            {
                MessageBox.Show(error, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var parameters = raw.ToDictionary(kvp => $"@{kvp.Key}", kvp => (object)kvp.Value);
            parameters["@id"] = id;
            db.ExecuteNonQuery(
                $"UPDATE inventory SET {string.Join(", ", raw.Keys.Select(k => $"{k}=@{k}"))} WHERE id=@id",
                parameters
            );
            LoadData();
            MessageBox.Show("Entry updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DeleteEntry()
        {
            if (dgvInventory.SelectedRows.Count == 0)
            {
                MessageBox.Show("No entry selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int id = Convert.ToInt32(dgvInventory.SelectedRows[0].Cells["ID"].Value);
            db.ExecuteNonQuery("DELETE FROM inventory WHERE id=@id", new Dictionary<string, object> { { "@id", id } });
            LoadData();
            MessageBox.Show("Entry deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SearchEntry()
        {
            string searchAssetTag = (entries[InventoryFields.AssetTag] as TextBox)?.Text.Trim() ?? "";
            string searchSerial = (entries[InventoryFields.SerialNumber] as TextBox)?.Text.Trim() ?? "";
            string searchTransferSheet = (entries[InventoryFields.TransferSheet] as TextBox)?.Text.Trim() ?? "";

            if (string.IsNullOrEmpty(searchAssetTag) && string.IsNullOrEmpty(searchSerial) && string.IsNullOrEmpty(searchTransferSheet))
            {
                MessageBox.Show("Enter an Asset Tag, Serial Number, or Transfer Sheet to search.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var parameters = new Dictionary<string, object>();
            string query = "SELECT * FROM inventory WHERE ";
            var conditions = new List<string>();

            if (!string.IsNullOrEmpty(searchAssetTag))
            {
                conditions.Add($"{InventoryFields.AssetTag} = @asset_tag");
                parameters["@asset_tag"] = searchAssetTag;
            }
            if (!string.IsNullOrEmpty(searchSerial))
            {
                conditions.Add($"{InventoryFields.SerialNumber} = @serial_number");
                parameters["@serial_number"] = searchSerial;
            }
            if (!string.IsNullOrEmpty(searchTransferSheet))
            {
                conditions.Add($"{InventoryFields.TransferSheet} = @transfer_sheet");
                parameters["@transfer_sheet"] = searchTransferSheet;
            }

            query += string.Join(" OR ", conditions);
            query += @" ORDER BY STRFTIME('%Y-%m-%d %H:%M', 
                                        date || ' ' || 
                                        CASE 
                                            WHEN time LIKE '%PM' AND SUBSTR(time, 1, INSTR(time, ':') - 1) != '12' THEN 
                                                CAST(CAST(SUBSTR(time, 1, INSTR(time, ':') - 1) AS INTEGER) + 12 AS TEXT) 
                                            WHEN time LIKE '%AM' AND SUBSTR(time, 1, INSTR(time, ':') - 1) = '12' THEN 
                                                '00' 
                                            ELSE 
                                                SUBSTR(time, 1, INSTR(time, ':') - 1) 
                                        END || SUBSTR(time, INSTR(time, ':'), 6)) DESC";

            var foundMain = db.ExecuteQuery(query, parameters);
            var foundKnown = new List<object[]>();

            if (foundMain.Count == 0)
            {
                string knownQuery = "SELECT * FROM known_inventory WHERE ";
                var knownConditions = new List<string>();
                var knownParams = new Dictionary<string, object>();

                if (!string.IsNullOrEmpty(searchAssetTag))
                {
                    knownConditions.Add($"{InventoryFields.AssetTag} = @asset_tag");
                    knownParams["@asset_tag"] = searchAssetTag;
                }
                if (!string.IsNullOrEmpty(searchSerial))
                {
                    knownConditions.Add($"{InventoryFields.SerialNumber} = @serial_number");
                    knownParams["@serial_number"] = searchSerial;
                }

                if (knownConditions.Any())
                {
                    knownQuery += string.Join(" OR ", knownConditions);
                    foundKnown = db.ExecuteQuery(knownQuery, knownParams);
                }
            }

            if (foundMain.Count > 0)
            {
                UpdateGrid(foundMain);
                var combinedMain = CombineRows(foundMain);
                FillFormWithSearchResults(combinedMain);
            }
            else if (foundKnown.Count > 0)
            {
                UpdateGrid(new List<object[]>());
                var combinedKnown = CombineRowsKnown(foundKnown);
                FillFormWithSearchResults(combinedKnown);
            }
            else
            {
                UpdateGrid(new List<object[]>());
                MessageBox.Show("No matching item found in any inventory.", "Asset Database", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private Dictionary<int, string> CombineRows(List<object[]> rows)
        {
            var combined = new Dictionary<int, List<string>>();
            for (int i = 1; i <= 11; i++)
                combined[i] = new List<string>();

            foreach (var row in rows)
            {
                for (int i = 1; i <= 11; i++)
                {
                    string val = row[i]?.ToString().Trim() ?? "";
                    if (!string.IsNullOrEmpty(val))
                        combined[i].Add(val);
                }
            }
            var result = new Dictionary<int, string>();
            foreach (var pair in combined)
                result[pair.Key] = string.Join(", ", pair.Value);
            return result;
        }

        private Dictionary<int, string> CombineRowsKnown(List<object[]> rows)
        {
            var combined = new Dictionary<int, List<string>>();
            for (int i = 1; i <= 3; i++)
                combined[i] = new List<string>();

            foreach (var row in rows)
            {
                if (row.Length > 1)
                {
                    string atag = row[1]?.ToString().Trim() ?? "";
                    if (!string.IsNullOrEmpty(atag))
                        combined[1].Add(atag);
                }
                if (row.Length > 2)
                {
                    string desc = row[2]?.ToString().Trim() ?? "";
                    if (!string.IsNullOrEmpty(desc))
                        combined[2].Add(desc);
                }
                if (row.Length > 3)
                {
                    string serial = row[3]?.ToString().Trim() ?? "";
                    if (!string.IsNullOrEmpty(serial))
                        combined[3].Add(serial);
                }
            }
            var result = new Dictionary<int, string>();
            foreach (var pair in combined)
                result[pair.Key] = string.Join(", ", pair.Value);
            return result;
        }

        private void FillFormWithSearchResults(Dictionary<int, string> combined)
        {
            var colMap = new Dictionary<string, int>
            {
                {InventoryFields.AssetTag, 1},
                {InventoryFields.Description, 2},
                {InventoryFields.SerialNumber, 3}
            };

            foreach (var pair in entries)
            {
                if (colMap.ContainsKey(pair.Key))
                {
                    int idx = colMap[pair.Key];
                    string val = combined.ContainsKey(idx) ? combined[idx] : "";
                    if (pair.Value is TextBox tb)
                        tb.Text = val;
                }
            }
        }

        private void FillFormWithCombined(Dictionary<int, string> combined)
        {
            var colMap = new Dictionary<string, int>
            {
                {InventoryFields.AssetTag, 1}, {InventoryFields.Description, 2}, {InventoryFields.SerialNumber, 3},
                {InventoryFields.TransferSheet, 4}, {InventoryFields.Notes, 5}, {InventoryFields.Date, 6},
                {InventoryFields.Time, 7}, {InventoryFields.Location, 8}, {InventoryFields.TransferredBy, 9},
                {InventoryFields.ReceivedBy, 10}, {InventoryFields.Color, 11}
            };

            foreach (var pair in entries)
            {
                if (colMap.ContainsKey(pair.Key))
                {
                    int idx = colMap[pair.Key];
                    string val = combined.ContainsKey(idx) ? combined[idx] : "";
                    if (pair.Value is DateTimePicker dtp)
                    {
                        if (pair.Key == InventoryFields.Date && DateTime.TryParseExact(val, "MM/dd/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                            dtp.Value = parsedDate;
                        else if (pair.Key == InventoryFields.Time && DateTime.TryParseExact(val, "hh:mm tt", null, System.Globalization.DateTimeStyles.None, out DateTime parsedTime))
                            dtp.Value = DateTime.Today.Add(parsedTime.TimeOfDay);
                    }
                    else if (pair.Value is TextBox tb)
                        tb.Text = val;
                    else if (pair.Value is ComboBox cb)
                        cb.Text = val;
                }
            }
        }

        private void ClearFields()
        {
            if (MessageBox.Show("Clear all fields?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                foreach (var ctrl in entries.Values)
                {
                    if (ctrl is DateTimePicker dtp)
                        dtp.Value = DateTime.Now;
                    else if (ctrl is TextBox tb)
                        tb.Text = "";
                    else if (ctrl is ComboBox cb)
                        cb.SelectedIndex = -1;
                }
            }
        }

        #endregion

        #region Grid and Data Loading

        private void LoadData()
        {
            var rows = db.ExecuteQuery(
                @"SELECT * FROM inventory 
                  ORDER BY STRFTIME('%Y-%m-%d %H:%M', 
                                   date || ' ' || 
                                   CASE 
                                       WHEN time LIKE '%PM' AND SUBSTR(time, 1, INSTR(time, ':') - 1) != '12' THEN 
                                           CAST(CAST(SUBSTR(time, 1, INSTR(time, ':') - 1) AS INTEGER) + 12 AS TEXT) 
                                       WHEN time LIKE '%AM' AND SUBSTR(time, 1, INSTR(time, ':') - 1) = '12' THEN 
                                           '00' 
                                       ELSE 
                                           SUBSTR(time, 1, INSTR(time, ':') - 1) 
                                   END || SUBSTR(time, INSTR(time, ':'), 6)) DESC 
                  LIMIT @pageSize OFFSET @offset",
                new Dictionary<string, object>
                {
                    {"@offset", (currentPage - 1) * PageSize},
                    {"@pageSize", PageSize}
                }
            );
            UpdateGrid(rows);
            UpdatePageInfo(Controls.Find("lblPageInfo", true).FirstOrDefault() as Label ?? new Label());
        }

        private void UpdateGrid(List<object[]> rows)
        {
            dgvInventory.Rows.Clear();
            foreach (var row in rows)
                dgvInventory.Rows.Add(row);

            dgvInventory.Sort(dgvInventory.Columns["Date"], System.ComponentModel.ListSortDirection.Descending);
        }

        private void UpdatePageInfo(Label lblPageInfo)
        {
            var totalRows = db.ExecuteQuery("SELECT COUNT(*) FROM inventory")[0][0];
            int totalPages = (int)Math.Ceiling((double)Convert.ToInt32(totalRows) / PageSize);
            lblPageInfo.Text = $"Page {currentPage} of {totalPages} ({totalRows} assets)";
        }

        private void SelectEntry()
        {
            if (dgvInventory.SelectedRows.Count == 0)
                return;

            DataGridViewRow selRow = dgvInventory.SelectedRows[0];
            var single = new Dictionary<int, string>();
            for (int i = 1; i <= 11; i++)
                single[i] = selRow.Cells[i].Value?.ToString().Trim() ?? "";
            FillFormWithCombined(single);
        }

        #endregion

        #region CSV Import/Export and Known Inventory

        private void ExportData()
        {
            var rows = db.ExecuteQuery("SELECT * FROM inventory");
            if (rows.Count == 0)
            {
                MessageBox.Show("No data available to export.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog { Filter = "CSV Files (*.csv)|*.csv", DefaultExt = "csv" };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(sfd.FileName, false, Encoding.UTF8))
                {
                    string[] headers = { "ID", "Asset Tag", "Description", "Serial Number", "Transfer Sheet", "Notes", "Date", "Time", "Location", "Transferred By", "Received By", "Color" };
                    sw.WriteLine(string.Join(",", headers));
                    foreach (var row in rows)
                    {
                        var fields = row.Select(f => f.ToString());
                        sw.WriteLine(string.Join(",", fields));
                    }
                }
                MessageBox.Show("Data exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ImportInventoryCSV()
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV Files (*.csv)|*.csv" };
            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string[] lines = File.ReadAllLines(ofd.FileName);
            if (lines.Length < 2)
            {
                MessageBox.Show("CSV file has no data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string[] expected = { "Asset Tag", "Description", "Serial Number", "Transfer Sheet", "Notes", "Date", "Time", "Location", "Transferred By", "Received By", "Color" };
            int insertedCount = 0;
            var entriesList = new List<Dictionary<string, string>>();

            string[] dateFormats = { "M/d/yyyy", "MM/dd/yyyy", "M/dd/yyyy", "MM/d/yyyy" };
            string[] timeFormats = { "h:mm tt", "hh:mm tt", "H:mm", "HH:mm" };

            for (int i = 1; i < lines.Length; i++)
            {
                string[] fields = lines[i].Split(',');
                var rowData = new Dictionary<string, string>();
                for (int j = 0; j < expected.Length; j++)
                    rowData[expected[j]] = (j < fields.Length) ? fields[j].Trim() : "";

                string dateStr = rowData["Date"].Trim();
                if (!DateTime.TryParseExact(dateStr, dateFormats, null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                {
                    MessageBox.Show($"Invalid date format at row {i + 1}: {dateStr}. Expected formats: 4/5/2025, 04/05/2025, etc.", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string timeStr = rowData["Time"].Trim();
                if (!DateTime.TryParseExact(timeStr, timeFormats, null, System.Globalization.DateTimeStyles.None, out DateTime parsedTime))
                {
                    MessageBox.Show($"Invalid time format at row {i + 1}: {timeStr}. Expected formats: 1:00 PM, 01:00 PM, 13:00, etc.", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var entry = new Dictionary<string, string>
                {
                    [InventoryFields.AssetTag] = rowData["Asset Tag"],
                    [InventoryFields.Description] = rowData["Description"],
                    [InventoryFields.SerialNumber] = rowData["Serial Number"],
                    [InventoryFields.TransferSheet] = rowData["Transfer Sheet"],
                    [InventoryFields.Notes] = rowData["Notes"],
                    [InventoryFields.Date] = parsedDate.ToString("MM/dd/yyyy"),
                    [InventoryFields.Time] = parsedTime.ToString("hh:mm tt"),
                    [InventoryFields.Location] = rowData["Location"],
                    [InventoryFields.TransferredBy] = rowData["Transferred By"],
                    [InventoryFields.ReceivedBy] = rowData["Received By"],
                    [InventoryFields.Color] = rowData["Color"]
                };
                entriesList.Add(entry);
            }

            try
            {
                db.BulkInsert(entriesList);
                insertedCount = entriesList.Count;
                currentPage = 1;
                LoadData();
                dgvInventory.Sort(dgvInventory.Columns["Date"], System.ComponentModel.ListSortDirection.Descending);
                MessageBox.Show($"Imported {insertedCount} rows from CSV.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                File.AppendAllText("error.log", $"{DateTime.Now}: CSV import failed - {ex.Message}\n");
                MessageBox.Show($"Error importing CSV: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadKnownInventory()
        {
            string directoryPath = @"C:\Temp\";
            var di = new DirectoryInfo(directoryPath);
            FileInfo[] excelFiles = di.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly);
            if (excelFiles.Length == 0)
            {
                MessageBox.Show("No Excel (.xlsx) files found in:\n" + directoryPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            FileInfo newestFile = excelFiles.OrderByDescending(f => f.LastWriteTime).First();

            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("SetNonCommercialPersonal");

                using (var package = new ExcelPackage(newestFile))
                {
                    var ws = package.Workbook.Worksheets["data"];
                    if (ws == null)
                    {
                        MessageBox.Show($"Worksheet 'data' not found in {newestFile.Name}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    int rowCount = ws.Dimension.Rows;
                    int insertedCount = 0;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string assetTag = ws.Cells[row, 1].Text.Trim();
                        string description = ws.Cells[row, 2].Text.Trim();
                        string serial = ws.Cells[row, 3].Text.Trim();

                        if (string.IsNullOrWhiteSpace(assetTag) && string.IsNullOrWhiteSpace(serial))
                            continue;

                        bool exists = false;
                        if (!string.IsNullOrEmpty(serial))
                        {
                            var count = db.ExecuteQuery(
                                $"SELECT COUNT(*) FROM known_inventory WHERE {InventoryFields.SerialNumber} = @serial",
                                new Dictionary<string, object> { { "@serial", serial } }
                            )[0][0];
                            exists = Convert.ToInt32(count) > 0;
                        }
                        else
                        {
                            var count = db.ExecuteQuery(
                                $"SELECT COUNT(*) FROM known_inventory WHERE {InventoryFields.AssetTag} = @asset_tag",
                                new Dictionary<string, object> { { "@asset_tag", assetTag } }
                            )[0][0];
                            exists = Convert.ToInt32(count) > 0;
                        }

                        if (exists)
                            continue;

                        try
                        {
                            db.ExecuteNonQuery(
                                $"INSERT INTO known_inventory ({InventoryFields.AssetTag}, {InventoryFields.Description}, {InventoryFields.SerialNumber}) VALUES (@asset_tag, @description, @serial_number)",
                                new Dictionary<string, object>
                                {
                                    { "@asset_tag", assetTag },
                                    { "@description", description },
                                    { "@serial_number", serial }
                                }
                            );
                            insertedCount++;
                        }
                        catch (SQLiteException)
                        {
                        }
                    }
                    MessageBox.Show($"Known Inventory loaded/updated from:\n{newestFile.Name}\n\nRows inserted: {insertedCount}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText("error.log", $"{DateTime.Now}: Excel import failed - {ex.Message}\n");
                MessageBox.Show("Error reading Excel file:\n" + ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Help Window

        private void ShowHelp()
        {
            string helpContent = "";
            using (var stream = new MemoryStream(Properties.Resources.helpme))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    helpContent = reader.ReadToEnd();
                }
            }
            Form helpForm = new Form
            {
                Text = "Help",
                Size = new Size(1600, 715)
            };
            TextBox txtHelp = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Both,
                Text = helpContent
            };
            helpForm.Controls.Add(txtHelp);
            helpForm.ShowDialog();
        }

        #endregion

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                db.ExecuteNonQuery("PRAGMA wal_checkpoint;", new Dictionary<string, object>());
            }
            catch (Exception ex)
            {
                File.AppendAllText("error.log", $"{DateTime.Now}: Error on close - {ex.Message}\n");
            }
            db.Dispose();
            base.OnFormClosing(e);
        }
    }
}