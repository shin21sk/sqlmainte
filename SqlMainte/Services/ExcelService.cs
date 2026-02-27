using ClosedXML.Excel;
using SqlMainte.Models;

namespace SqlMainte.Services;

public class ExcelService
{
    /// <summary>
    /// グリッドのデータを Excel ファイルにエクスポートする。
    /// DeletePending の行は除外し、現在の表示値をそのまま出力する。
    /// </summary>
    public void Export(
        string filePath,
        List<ColumnInfo> columns,
        IEnumerable<IReadOnlyDictionary<string, string>> rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // ヘッダー行
        for (int c = 0; c < columns.Count; c++)
        {
            var cell = ws.Cell(1, c + 1);
            cell.Value = columns[c].ColumnName;
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0x1F, 0x49, 0x7D); // 濃い青
            cell.Style.Font.FontColor = XLColor.White;
        }

        // データ行
        int rowNum = 2;
        foreach (var row in rows)
        {
            for (int c = 0; c < columns.Count; c++)
            {
                var colName = columns[c].ColumnName;
                var val = row.TryGetValue(colName, out var v) ? v : string.Empty;
                ws.Cell(rowNum, c + 1).Value = val;
            }
            rowNum++;
        }

        // 列幅自動調整
        ws.Columns().AdjustToContents();

        wb.SaveAs(filePath);
    }

    /// <summary>
    /// Excel ファイルを読み込み、列名をキーとする行データのリストを返す。
    /// </summary>
    /// <param name="filePath">読み込む .xlsx ファイルのパス</param>
    /// <param name="requiredColumns">存在確認する列名リスト（警告用）</param>
    /// <returns>行データのリスト。列名は元の大文字小文字を保持。</returns>
    public (List<Dictionary<string, string>> Rows, List<string> Warnings) Import(
        string filePath,
        List<ColumnInfo> requiredColumns)
    {
        var warnings = new List<string>();

        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheets.First();

        // ヘッダー行を取得
        var headerRow = ws.Row(1);
        var headers = new Dictionary<int, string>(); // 列番号 → 列名

        int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (int c = 1; c <= lastCol; c++)
        {
            var h = headerRow.Cell(c).GetString().Trim();
            if (!string.IsNullOrEmpty(h))
                headers[c] = h;
        }

        // 必須列の存在確認
        foreach (var col in requiredColumns)
        {
            if (!headers.Values.Any(h => h.Equals(col.ColumnName, StringComparison.OrdinalIgnoreCase)))
                warnings.Add($"列 [{col.ColumnName}] が Excel に見つかりません。空欄として扱われます。");
        }

        // データ行を読み込み
        var rows = new List<Dictionary<string, string>>();
        int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

        for (int r = 2; r <= lastRow; r++)
        {
            var wsRow = ws.Row(r);

            // 行全体が空なら読み飛ばす
            if (headers.Keys.All(c => string.IsNullOrWhiteSpace(wsRow.Cell(c).GetString())))
                continue;

            var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (colNum, colName) in headers)
                rowData[colName] = wsRow.Cell(colNum).GetString().Trim();

            rows.Add(rowData);
        }

        return (rows, warnings);
    }
}
