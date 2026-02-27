# CLAUDE.md

## プロジェクト概要

SQL Server のマスタテーブルを Excel 風グリッドで編集・保存する Windows デスクトップツール。

- **技術**: .NET 8 / C# / Windows Forms
- **プロジェクトパス**: `SqlMainte/`

## よく使うコマンド

```bash
# ビルド
cd SqlMainte && dotnet build

# リリースビルド
cd SqlMainte && dotnet build -c Release

# 実行
cd SqlMainte && dotnet run
```

## プロジェクト構成

```
SqlMainte/
├── appsettings.json              # 接続文字列・対象テーブル定義
├── app.ico                       # アプリアイコン（PowerShell で生成）
├── Program.cs                    # エントリポイント・グローバル例外ハンドラ
├── Models/
│   ├── AppSettings.cs            # AppSettings, TableConfig モデル
│   └── ColumnInfo.cs             # カラム情報（IsBinary プロパティあり）
├── Services/
│   ├── BinaryColumnSerializer.cs # ★ バイナリ列変換の一元管理（後述）
│   ├── ConfigService.cs          # appsettings.json 読込
│   ├── SchemaService.cs          # INFORMATION_SCHEMA からカラム定義取得
│   ├── DatabaseService.cs        # INSERT / UPDATE / DELETE
│   └── ExcelService.cs           # Excel エクスポート・インポート（ClosedXML）
└── Forms/
    └── MainForm.cs               # メイン画面（グリッド編集 UI）
```

## appsettings.json の構造

```json
{
  "ConnectionString": "接続文字列",
  "Tables": [
    {
      "Name": "実テーブル名",
      "DisplayName": "画面表示名",
      "PrimaryKeys": ["PK列1", "PK列2"]
    }
  ]
}
```

- `PrimaryKeys` は配列で複合主キーに対応
- `appsettings.json` は `bin/` にコピーされるため、開発中は `SqlMainte/appsettings.json` を編集する

## 主要な設計方針

### 行状態管理（MainForm.cs）

`Dictionary<DataGridViewRow, RowState>` で各行の状態を追跡する。

| 状態 | 色 | 保存時 |
|------|-----|--------|
| Unchanged | 白 | スキップ |
| Modified | 黄 | UPDATE |
| Added | 緑 | INSERT |
| DeletePending | 赤 | DELETE |

- 既存行の元 PK 値は `_originalKeys` に保持し UPDATE / DELETE の WHERE 句に使用
- DB から読んだ元データは `_originalDbRows` に保持し Excel インポート時の差分判定に使用

### バイナリ列の変換（BinaryColumnSerializer.cs）

`varbinary` / `binary` 型の列はカンマ区切りテキストで表示・入力し、保存時に変換する。

- **現在のフォーマット**: `string[]` を JSON UTF-8 でシリアライズ
- **フォーマットを変更する場合は `BinaryColumnSerializer.cs` の `Serialize` / `Deserialize` のみ修正すればよい**
- 空欄は空リスト `[]` をシリアライズして保存（NULL にはしない）

### 保存処理（DatabaseService.cs）

- INSERT / UPDATE / DELETE を 1 トランザクションで実行
- 1 件でも失敗したら全件ロールバック
- IDENTITY 列は INSERT 時に列リストから除外（SQL Server が自動採番）

### Excel インポートの差分判定（MainForm.cs `ApplyImportedRows`）

PK をキーに `_originalDbRows` と照合して行状態を自動設定する。
Excel にあって DB にない行 → Added、DB にあって Excel にない行 → DeletePending。

## 注意事項

- `Validate()` は `ContainerControl.Validate()` と名前が衝突するため `ValidateInput()` という名前にしている
- ツールバーのボタンは `ToolStripButton`、下部パネルのボタンは `Button + AutoSize=true + FlowLayoutPanel` を使用
- `dotnet new winforms` で `--framework net8.0-windows` は指定不可（SDK が .NET 10 のため）。`net8.0` で作成後に自動で `net8.0-windows` になる
