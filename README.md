# SQL テーブルメンテナンスツール

SQL Server のマスタテーブルを Excel のようなグリッド形式で編集・保存できる Windows デスクトップツールです。

---

## 動作環境

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11 |
| ランタイム | .NET 8 以上（Windows） |
| DB | SQL Server（Express / Standard / Developer 等） |

---

## セットアップ

### 1. ビルド

```bash
cd SqlMainte
dotnet build -c Release
```

実行ファイルは `bin/Release/net8.0-windows/SqlMainte.exe` に生成されます。

### 2. 接続先・テーブルの設定

`SqlMainte.exe` と同じフォルダにある **`appsettings.json`** を編集します。

```json
{
  "ConnectionString": "Server=サーバー名;Database=DB名;Trusted_Connection=True;TrustServerCertificate=True;",
  "Tables": [
    {
      "Name": "M_Category",
      "DisplayName": "カテゴリマスタ",
      "PrimaryKeys": [ "CategoryId" ]
    },
    {
      "Name": "M_Product",
      "DisplayName": "商品マスタ",
      "PrimaryKeys": [ "CompanyId", "ProductCode" ]
    }
  ]
}
```

#### 設定項目

| キー | 説明 |
|------|------|
| `ConnectionString` | SQL Server への接続文字列 |
| `Tables[].Name` | SQL Server 上の実際のテーブル名 |
| `Tables[].DisplayName` | 画面上のコンボボックスに表示する名前 |
| `Tables[].PrimaryKeys` | 主キー列名の配列（複合主キーも対応） |

#### 接続文字列の例

```
# Windows 認証
Server=localhost;Database=MyDB;Trusted_Connection=True;TrustServerCertificate=True;

# SQL Server 認証
Server=192.168.1.10,1433;Database=MyDB;User Id=sa;Password=yourPassword;TrustServerCertificate=True;

# 名前付きインスタンス
Server=MYSERVER\SQLEXPRESS;Database=MyDB;Trusted_Connection=True;TrustServerCertificate=True;
```

---

## 画面の見方

```
+--[ テーブル: [カテゴリマスタ ▼] ]--[ 再読込 ]--[ キャンセル ]--[ 保存 ]--[ Excel出力 ]--[ Excelインポート ]--+
|                                                                                                          |
|  ID  |  コード  |  名称        |  備考                                                                    |
|   1  |  A001    |  カテゴリA   |                   ← 白：未変更                                           |
|   2  |  A002    |  カテゴリB改  |                  ← 黄：編集済み                                          |
|   3  |  A003    |  新規        |                   ← 緑：新規追加                                          |
|   4  |  A004    |  削除予定    |                   ← 赤：削除予定（保存時に削除）                             |
|                                                                                                          |
|  [ 行追加 ]  [ 行削除 ]                                                                                   |
+--[ 4 件読込完了 ]--------------------------------------------------------------------------------+
```

### 行の色の意味

| 色 | 意味 | 保存時の動作 |
|----|------|------------|
| 白 | 未変更 | 何もしない |
| 黄 | 編集済み | UPDATE |
| 緑 | 新規追加行 | INSERT |
| 赤 | 削除予定行（編集不可） | DELETE |

---

## 操作方法

### データの編集

1. 画面上部のコンボボックスでテーブルを選択するとデータが一覧表示されます
2. セルを直接クリックして編集できます（Excel と同様）
3. **主キー列は既存行では編集不可**です（誤変更防止）

### 行の追加

- **`行追加`** ボタンを押すと末尾に空行（緑）が追加されます
- 各セルに値を入力してください
- IDENTITY（自動採番）列は空欄のまま保存すると SQL Server が自動採番します

### 行の削除

1. 削除したい行をクリックして選択（複数行は Ctrl / Shift 併用）
2. **`行削除`** ボタンを押すと赤色（削除予定）になります
3. 新規追加行（緑）を削除すると即座に行が消えます
4. 削除を取り消すには **`キャンセル`** または **`再読込`** で元に戻ります

### 保存

- **`保存`** ボタンを押すと確認ダイアログが表示されます
- OK すると追加・変更・削除をまとめて **1トランザクション** で実行します
- 1件でもエラーになった場合は全件ロールバックされます

### 再読込 / キャンセル

- **`再読込`** / **`キャンセル`** ボタン：未保存の変更を破棄して DB から再取得します
- 未保存の変更がある場合は確認ダイアログが表示されます

---

## Excel エクスポート / インポート

### エクスポート

1. **`Excel出力`** ボタンを押してファイルの保存先を選択します
2. グリッドに表示中のデータ（削除予定行を除く）が `.xlsx` 形式で出力されます
3. 出力後に「ファイルを開きますか？」と確認されます

### インポート

1. **`Excelインポート`** ボタンを押してファイルを選択します
2. Excel の **1行目をヘッダー行**として読み込み、列名でグリッドの列と照合します
3. 読み込み後、DB の現在のデータと自動比較してグリッドに反映します

| 条件 | 行の色 |
|------|--------|
| Excel の行と DB の値が同じ | 白（未変更） |
| Excel の行と DB の値が違う | 黄（変更あり） |
| Excel にあって DB にない行 | 緑（新規追加） |
| DB にあって Excel にない行 | 赤（削除予定） |

4. グリッドで差分を確認後、**`保存`** ボタンで確定します

#### Excel ファイルの注意事項

- **1行目は列名（ヘッダー）**として扱われます。データは2行目以降に記載してください
- 列名は大文字・小文字を区別しません
- Excel に存在しない列は空欄として扱われます
- 空行は自動的にスキップされます

#### 典型的なワークフロー

```
[Excel出力] → Excel で編集・追記 → [Excelインポート] → グリッドで差分確認 → [保存]
```

---

## バイナリ列（varbinary / binary）について

`varbinary` や `binary` 型の列は、内部的に文字列の配列を JSON UTF-8 でシリアライズした値として扱います。

- **表示**：バイナリデータをデシリアライズして **カンマ区切りのテキスト** で表示します
  - 例：`apple,banana,orange`
- **入力**：カンマ区切りで文字列を入力します
  - 例：`東京,大阪,名古屋`
- **空欄**：空リスト `[]` をシリアライズして保存します（NULL にはなりません）

---

## バリデーション

保存時に以下のチェックが実行されます。

| チェック内容 | 対象 |
|------------|------|
| 必須チェック（NOT NULL） | 追加・変更行の NULL 非許容列 |
| 文字数チェック | 最大桁数を超えた入力 |

エラーがある場合はメッセージが表示され、保存は実行されません。

---

## 依存ライブラリ

| ライブラリ | バージョン | 用途 |
|-----------|-----------|------|
| Microsoft.Data.SqlClient | 6.1.4 | SQL Server 接続・クエリ実行 |
| ClosedXML | 0.105.0 | Excel 読み書き |
| Microsoft.Extensions.Configuration.Json | - | appsettings.json 読込 |
| Microsoft.Extensions.Configuration.Binder | - | 設定オブジェクトへのバインド |
