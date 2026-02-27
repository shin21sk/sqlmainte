using System.Text;
using System.Text.Json;

namespace SqlMainte.Services;

/// <summary>
/// バイナリ列（varbinary等）とUI表示文字列（カンマ区切り）の相互変換。
/// フォーマットを変更する場合はこのクラスのみ修正すればよい。
/// 現在のフォーマット: string[] を JSON UTF-8 でシリアライズ
/// </summary>
public static class BinaryColumnSerializer
{
    /// <summary>
    /// byte[] → カンマ区切り表示文字列
    /// </summary>
    public static string ToDisplayString(byte[]? data)
    {
        if (data is null || data.Length == 0)
            return string.Empty;

        var values = Deserialize(data);
        return string.Join(",", values);
    }

    /// <summary>
    /// カンマ区切り表示文字列 → byte[]
    /// 空文字列・null は空リスト [] をシリアライズして返す
    /// </summary>
    public static byte[] FromDisplayString(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return Serialize([]);

        var values = text.Split(',');
        return Serialize(values);
    }

    // ---- フォーマット切替はここだけ ----

    private static byte[] Serialize(string[] values)
        => JsonSerializer.SerializeToUtf8Bytes(values);

    private static string[] Deserialize(byte[] data)
    {
        try
        {
            return JsonSerializer.Deserialize<string[]>(data) ?? [];
        }
        catch
        {
            // 読めなかった場合は生バイト列をカンマ区切りで返す（フォーマット変更前のデータ対策）
            return data.Select(b => b.ToString()).ToArray();
        }
    }
}
