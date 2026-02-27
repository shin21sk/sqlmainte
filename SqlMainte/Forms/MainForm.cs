using SqlMainte.Models;
using SqlMainte.Services;

namespace SqlMainte.Forms;

public class MainForm : Form
{
    // ---- 行状態 ----
    private enum RowState { Unchanged, Modified, Added, DeletePending }

    // ---- 定数 ----
    private static readonly Color ColorUnchanged     = Color.White;
    private static readonly Color ColorModified      = Color.LightYellow;
    private static readonly Color ColorAdded         = Color.LightGreen;
    private static readonly Color ColorDeletePending = Color.FromArgb(255, 182, 182); // 薄い赤

    // ---- UI部品 ----
    private readonly ComboBox _cboTable         = new();
    private readonly ToolStripButton _btnReload = new() { Text = "再読込",        DisplayStyle = ToolStripItemDisplayStyle.Text };
    private readonly ToolStripButton _btnCancel = new() { Text = "キャンセル",    DisplayStyle = ToolStripItemDisplayStyle.Text };
    private readonly ToolStripButton _btnSave   = new() { Text = "　保存　",      DisplayStyle = ToolStripItemDisplayStyle.Text };
    private readonly ToolStripButton _btnExport = new() { Text = "Excel出力",     DisplayStyle = ToolStripItemDisplayStyle.Text };
    private readonly ToolStripButton _btnImport = new() { Text = "Excelインポート", DisplayStyle = ToolStripItemDisplayStyle.Text };
    private readonly Button _btnAddRow          = new() { Text = "行追加" };
    private readonly Button _btnDelRow          = new() { Text = "行削除" };
    private readonly DataGridView _grid  = new();
    private readonly StatusStrip _status = new();
    private readonly ToolStripStatusLabel _lblStatus = new();

    // ---- 状態 ----
    private List<ColumnInfo> _columns = [];
    private readonly Dictionary<DataGridViewRow, RowState> _rowStates = [];
    // 既存行の元PK値（UPDATE/DELETE の WHERE 用）
    private readonly Dictionary<DataGridViewRow, Dictionary<string, object?>> _originalKeys = [];
    // DB から読み込んだ元データ（インポート時の差分判定用）
    private List<Dictionary<string, object?>> _originalDbRows = [];

    private AppSettings _settings = null!;
    private DatabaseService _db = null!;
    private SchemaService _schema = null!;
    private TableConfig CurrentTable => (TableConfig)_cboTable.SelectedItem!;

    public MainForm()
    {
        InitializeLayout();
        LoadSettings();
        WireEvents();
        if (_cboTable.Items.Count > 0)
            _cboTable.SelectedIndex = 0;
    }

    // ================================================================
    //  初期化
    // ================================================================
    private void LoadSettings()
    {
        _settings = ConfigService.Load();
        _db = new DatabaseService(_settings.ConnectionString);
        _schema = new SchemaService(_settings.ConnectionString);

        _cboTable.Items.Clear();
        foreach (var t in _settings.Tables)
            _cboTable.Items.Add(t);
        _cboTable.DisplayMember = "DisplayName";
    }

    private void WireEvents()
    {
        _cboTable.SelectedIndexChanged += (_, _) => LoadTable();
        _btnReload.Click += (_, _) => ReloadWithConfirm();
        _btnCancel.Click += (_, _) => ReloadWithConfirm();
        _btnSave.Click  += (_, _) => SaveChanges();
        _btnAddRow.Click  += (_, _) => AddRow();
        _btnDelRow.Click  += (_, _) => DeleteSelectedRows();
        _btnExport.Click  += (_, _) => ExportToExcel();
        _btnImport.Click  += (_, _) => ImportFromExcel();

        _grid.CellValueChanged    += OnCellValueChanged;
        _grid.CellBeginEdit       += OnCellBeginEdit;
        _grid.DataError           += (_, e) => e.ThrowException = false;
    }

    // ================================================================
    //  テーブル読込
    // ================================================================
    private void LoadTable()
    {
        if (_cboTable.SelectedItem is null) return;

        try
        {
            SetStatus("読込中...");
            _rowStates.Clear();
            _originalKeys.Clear();

            var tbl = CurrentTable;
            _columns = _schema.GetColumns(tbl.Name, tbl.PrimaryKeys);

            BuildGridColumns();

            var rows = _db.FetchAll(tbl.Name, _columns);
            _originalDbRows = rows; // インポート時の差分判定用に保持

            _grid.SuspendLayout();
            _grid.Rows.Clear();
            foreach (var row in rows)
            {
                var values = _columns.Select(c => row.TryGetValue(c.ColumnName, out var v) ? v?.ToString() ?? "" : "").ToArray();
                int idx = _grid.Rows.Add(values);
                var dgvRow = _grid.Rows[idx];
                _rowStates[dgvRow] = RowState.Unchanged;
                _originalKeys[dgvRow] = ExtractKeys(dgvRow);
            }
            _grid.ResumeLayout();

            ApplyRowColors();
            SetStatus($"{rows.Count} 件読込完了");
        }
        catch (Exception ex)
        {
            SetStatus("エラー");
            MessageBox.Show($"読込エラー:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BuildGridColumns()
    {
        _grid.Columns.Clear();
        foreach (var col in _columns)
        {
            var dgvCol = new DataGridViewTextBoxColumn
            {
                Name = col.ColumnName,
                HeaderText = col.ColumnName,
                DataPropertyName = col.ColumnName,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };

            // バイナリ列はツールチップで説明
            if (col.IsBinary)
                dgvCol.ToolTipText = "カンマ区切りで入力（例: apple,banana）";

            _grid.Columns.Add(dgvCol);
        }
    }

    // ================================================================
    //  セルイベント
    // ================================================================
    private string? _cellValueBeforeEdit;

    private void OnCellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
    {
        var row = _grid.Rows[e.RowIndex];

        // 削除予定行は編集不可
        if (_rowStates.TryGetValue(row, out var state) && state == RowState.DeletePending)
        {
            e.Cancel = true;
            return;
        }

        // 既存行のPK列は読取専用
        var col = _columns[e.ColumnIndex];
        if (col.IsPrimaryKey && _rowStates.TryGetValue(row, out var rs) && rs != RowState.Added)
        {
            e.Cancel = true;
            return;
        }

        _cellValueBeforeEdit = row.Cells[e.ColumnIndex].Value?.ToString();
    }

    private void OnCellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        var row = _grid.Rows[e.RowIndex];

        if (!_rowStates.TryGetValue(row, out var state)) return;
        if (state == RowState.Unchanged)
        {
            var newVal = row.Cells[e.ColumnIndex].Value?.ToString();
            if (newVal != _cellValueBeforeEdit)
            {
                _rowStates[row] = RowState.Modified;
                ApplyRowColor(row);
            }
        }
    }

    // ================================================================
    //  行追加・削除
    // ================================================================
    private void AddRow()
    {
        int idx = _grid.Rows.Add();
        var row = _grid.Rows[idx];
        _rowStates[row] = RowState.Added;
        ApplyRowColor(row);
        _grid.CurrentCell = row.Cells[0];
    }

    private void DeleteSelectedRows()
    {
        foreach (DataGridViewRow row in _grid.SelectedRows)
        {
            if (row.IsNewRow) continue;

            if (_rowStates.TryGetValue(row, out var state) && state == RowState.Added)
            {
                // 新規行はその場で除去
                _rowStates.Remove(row);
                _originalKeys.Remove(row);
                _grid.Rows.Remove(row);
            }
            else
            {
                _rowStates[row] = RowState.DeletePending;
                ApplyRowColor(row);
            }
        }
    }

    // ================================================================
    //  保存
    // ================================================================
    private void SaveChanges()
    {
        // バリデーション
        var errors = ValidateInput();
        if (errors.Count > 0)
        {
            MessageBox.Show("入力エラーがあります:\n\n" + string.Join("\n", errors),
                "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (MessageBox.Show("保存します。よろしいですか？", "確認",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        try
        {
            SetStatus("保存中...");

            var toInsert = new List<Dictionary<string, object?>>();
            var toUpdate = new List<Dictionary<string, object?>>();
            var toDelete = new List<Dictionary<string, object?>>();

            foreach (DataGridViewRow row in _grid.Rows)
            {
                if (row.IsNewRow) continue;
                if (!_rowStates.TryGetValue(row, out var state)) continue;

                switch (state)
                {
                    case RowState.Added:
                        toInsert.Add(ExtractRow(row));
                        break;
                    case RowState.Modified:
                        toUpdate.Add(ExtractRowWithOriginalKeys(row));
                        break;
                    case RowState.DeletePending:
                        toDelete.Add(_originalKeys.TryGetValue(row, out var keys) ? keys : ExtractRow(row));
                        break;
                }
            }

            _db.SaveChanges(CurrentTable.Name, _columns, toInsert, toUpdate, toDelete);
            SetStatus($"保存完了（追加:{toInsert.Count} 更新:{toUpdate.Count} 削除:{toDelete.Count}）");
            LoadTable();
        }
        catch (Exception ex)
        {
            SetStatus("保存エラー");
            MessageBox.Show($"保存エラー:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private List<string> ValidateInput()
    {
        var errors = new List<string>();

        foreach (DataGridViewRow row in _grid.Rows)
        {
            if (row.IsNewRow) continue;
            if (!_rowStates.TryGetValue(row, out var state)) continue;
            if (state == RowState.DeletePending || state == RowState.Unchanged) continue;

            foreach (var col in _columns)
            {
                if (col.IsIdentity && state == RowState.Added) continue;

                var val = row.Cells[col.ColumnName].Value?.ToString() ?? "";

                if (!col.IsNullable && val == string.Empty && !col.IsBinary)
                    errors.Add($"行{row.Index + 1} [{col.ColumnName}] は必須項目です。");

                if (col.MaxLength.HasValue && col.MaxLength > 0 && val.Length > col.MaxLength)
                    errors.Add($"行{row.Index + 1} [{col.ColumnName}] は {col.MaxLength} 文字以内で入力してください。");
            }
        }

        return errors;
    }

    // ================================================================
    //  ユーティリティ
    // ================================================================
    private void ReloadWithConfirm()
    {
        bool hasChanges = _rowStates.Values.Any(s => s != RowState.Unchanged);
        if (hasChanges)
        {
            if (MessageBox.Show("未保存の変更があります。破棄して再読込しますか？", "確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;
        }
        LoadTable();
    }

    private Dictionary<string, object?> ExtractRow(DataGridViewRow row)
    {
        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var col in _columns)
            dict[col.ColumnName] = row.Cells[col.ColumnName].Value?.ToString();
        return dict;
    }

    private Dictionary<string, object?> ExtractRowWithOriginalKeys(DataGridViewRow row)
    {
        var dict = ExtractRow(row);
        // UPDATE の WHERE 用に元PKを上書き
        if (_originalKeys.TryGetValue(row, out var origKeys))
            foreach (var (k, v) in origKeys)
                dict[$"__orig_{k}"] = v;
        return dict;
    }

    private Dictionary<string, object?> ExtractKeys(DataGridViewRow row)
    {
        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var col in _columns.Where(c => c.IsPrimaryKey))
            dict[col.ColumnName] = row.Cells[col.ColumnName].Value?.ToString();
        return dict;
    }

    private void ApplyRowColors()
    {
        foreach (DataGridViewRow row in _grid.Rows)
            ApplyRowColor(row);
    }

    private void ApplyRowColor(DataGridViewRow row)
    {
        if (!_rowStates.TryGetValue(row, out var state)) return;
        row.DefaultCellStyle.BackColor = state switch
        {
            RowState.Modified     => ColorModified,
            RowState.Added        => ColorAdded,
            RowState.DeletePending => ColorDeletePending,
            _                    => ColorUnchanged
        };
    }

    private void SetStatus(string message)
    {
        _lblStatus.Text = message;
        _status.Refresh();
    }

    // ================================================================
    //  Excel エクスポート
    // ================================================================
    private void ExportToExcel()
    {
        if (_columns.Count == 0)
        {
            MessageBox.Show("テーブルを選択してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dlg = new SaveFileDialog
        {
            Title = "Excel 出力先を選択",
            Filter = "Excel ファイル (*.xlsx)|*.xlsx",
            FileName = $"{CurrentTable.Name}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
        };

        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            SetStatus("Excel 出力中...");

            // DeletePending 行を除いた現在のグリッドデータを収集
            var exportRows = new List<IReadOnlyDictionary<string, string>>();
            foreach (DataGridViewRow row in _grid.Rows)
            {
                if (row.IsNewRow) continue;
                if (_rowStates.TryGetValue(row, out var state) && state == RowState.DeletePending)
                    continue;

                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var col in _columns)
                    dict[col.ColumnName] = row.Cells[col.ColumnName].Value?.ToString() ?? string.Empty;

                exportRows.Add(dict);
            }

            new ExcelService().Export(dlg.FileName, _columns, exportRows);
            SetStatus($"Excel 出力完了（{exportRows.Count} 件）");

            // 出力後にファイルを開くか確認
            if (MessageBox.Show("ファイルを開きますか？", "Excel 出力完了",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = dlg.FileName,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            SetStatus("Excel 出力エラー");
            MessageBox.Show($"出力エラー:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ================================================================
    //  Excel インポート
    // ================================================================
    private void ImportFromExcel()
    {
        if (_columns.Count == 0)
        {
            MessageBox.Show("テーブルを選択してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dlg = new OpenFileDialog
        {
            Title = "インポートする Excel ファイルを選択",
            Filter = "Excel ファイル (*.xlsx)|*.xlsx"
        };

        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            SetStatus("Excel 読込中...");

            var svc = new ExcelService();
            var (importedRows, warnings) = svc.Import(dlg.FileName, _columns);

            // 警告があれば表示（中断はしない）
            if (warnings.Count > 0)
            {
                MessageBox.Show(
                    "以下の警告があります。続行しますか？\n\n" + string.Join("\n", warnings),
                    "インポート警告",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            if (importedRows.Count == 0)
            {
                MessageBox.Show("インポートするデータがありませんでした。", "情報",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetStatus("インポート完了（0件）");
                return;
            }

            ApplyImportedRows(importedRows);
            SetStatus($"インポート完了（{importedRows.Count} 件）");
        }
        catch (Exception ex)
        {
            SetStatus("Excel インポートエラー");
            MessageBox.Show($"インポートエラー:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// インポートした行データをグリッドに反映し、差分に応じて行状態を設定する。
    /// ・Excel にある行 → DB と照合して Added / Modified / Unchanged
    /// ・DB にあるが Excel にない行 → DeletePending
    /// </summary>
    private void ApplyImportedRows(List<Dictionary<string, string>> importedRows)
    {
        // 元DBデータを PK キーで引ける辞書にする
        var originalByPk = _originalDbRows.ToDictionary(
            r => GetPkKey(r),
            r => r,
            StringComparer.OrdinalIgnoreCase);

        _rowStates.Clear();
        _originalKeys.Clear();
        _grid.SuspendLayout();
        _grid.Rows.Clear();

        var importedPkKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var importRow in importedRows)
        {
            var pkKey = GetPkKeyFromStrings(importRow);
            importedPkKeys.Add(pkKey);

            var values = _columns
                .Select(c => importRow.TryGetValue(c.ColumnName, out var v) ? v : string.Empty)
                .ToArray<object?>();

            int idx = _grid.Rows.Add(values);
            var dgvRow = _grid.Rows[idx];

            if (originalByPk.TryGetValue(pkKey, out var origRow))
            {
                // DB に存在する行：値が変わっているか比較
                bool changed = _columns.Any(c =>
                {
                    var origVal = origRow.TryGetValue(c.ColumnName, out var ov) ? ov?.ToString() ?? "" : "";
                    var newVal  = importRow.TryGetValue(c.ColumnName, out var nv) ? nv : "";
                    return !string.Equals(origVal, newVal, StringComparison.Ordinal);
                });

                _rowStates[dgvRow]  = changed ? RowState.Modified : RowState.Unchanged;
                _originalKeys[dgvRow] = ExtractKeys(dgvRow);
            }
            else
            {
                // DB に存在しない行 → 新規
                _rowStates[dgvRow] = RowState.Added;
            }
        }

        // DB にあって Excel にない行 → DeletePending として追加表示
        foreach (var origRow in _originalDbRows)
        {
            var pkKey = GetPkKey(origRow);
            if (importedPkKeys.Contains(pkKey)) continue;

            var values = _columns
                .Select(c => origRow.TryGetValue(c.ColumnName, out var v) ? v?.ToString() ?? "" : "")
                .ToArray<object?>();

            int idx = _grid.Rows.Add(values);
            var dgvRow = _grid.Rows[idx];
            _rowStates[dgvRow]  = RowState.DeletePending;
            _originalKeys[dgvRow] = ExtractKeys(dgvRow);
        }

        _grid.ResumeLayout();
        ApplyRowColors();
    }

    /// <summary>DB 取得行（object? 値）から PK 複合キー文字列を生成する</summary>
    private string GetPkKey(Dictionary<string, object?> row)
        => string.Join("\0", _columns
            .Where(c => c.IsPrimaryKey)
            .Select(c => row.TryGetValue(c.ColumnName, out var v) ? v?.ToString() ?? "" : ""));

    /// <summary>インポート行（string 値）から PK 複合キー文字列を生成する</summary>
    private string GetPkKeyFromStrings(Dictionary<string, string> row)
        => string.Join("\0", _columns
            .Where(c => c.IsPrimaryKey)
            .Select(c => row.TryGetValue(c.ColumnName, out var v) ? v : ""));

    // ================================================================
    //  レイアウト
    // ================================================================
    private void InitializeLayout()
    {
        Text = "SQL テーブルメンテナンス";
        Size = new Size(1100, 700);
        MinimumSize = new Size(800, 500);
        StartPosition = FormStartPosition.CenterScreen;
        WindowState = FormWindowState.Maximized;
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

        // ツールバー（上部）
        var toolbar = new ToolStrip
        {
            GripStyle = ToolStripGripStyle.Hidden,
            Padding = new Padding(4, 2, 4, 2),
            AutoSize = true
        };

        // テーブル選択コンボ
        toolbar.Items.Add(new ToolStripLabel("テーブル:"));
        _cboTable.DropDownStyle = ComboBoxStyle.DropDownList;
        var cboHost = new ToolStripControlHost(_cboTable) { AutoSize = false, Width = 200 };
        toolbar.Items.Add(cboHost);

        toolbar.Items.Add(new ToolStripSeparator());

        // 再読込・キャンセル
        toolbar.Items.Add(_btnReload);
        toolbar.Items.Add(_btnCancel);

        toolbar.Items.Add(new ToolStripSeparator());

        // 保存（太字・強調色）
        _btnSave.Font = new Font(toolbar.Font, FontStyle.Bold);
        _btnSave.BackColor = Color.SteelBlue;
        _btnSave.ForeColor = Color.White;
        toolbar.Items.Add(_btnSave);

        toolbar.Items.Add(new ToolStripSeparator());

        // Excel 操作
        toolbar.Items.Add(_btnExport);
        toolbar.Items.Add(_btnImport);

        // グリッド
        _grid.Dock = DockStyle.Fill;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.MultiSelect = true;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _grid.RowHeadersVisible = false;
        _grid.BorderStyle = BorderStyle.None;
        _grid.Font = new Font("Meiryo UI", 9f);

        // 下部パネル（行操作ボタン）
        _btnAddRow.AutoSize = true;
        _btnDelRow.AutoSize = true;
        var bottomPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 36,
            Padding = new Padding(4, 4, 0, 0),
            WrapContents = false
        };
        bottomPanel.Controls.Add(_btnAddRow);
        bottomPanel.Controls.Add(_btnDelRow);

        // ステータスバー
        _status.Items.Add(_lblStatus);

        Controls.Add(_grid);
        Controls.Add(bottomPanel);
        Controls.Add(toolbar);
        Controls.Add(_status);
    }
}
