using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PDMTools.Models
{
    /// <summary>
    /// 使用者選擇要顯示的資料卡變數設定，持久化至 columns.json。
    /// </summary>
    public sealed class ColumnSettings
    {
        public static readonly string FilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "columns.json");

        /// <summary>使用者勾選的資料卡變數名稱（與 PDM Vault 定義一致）。</summary>
        public List<string> SelectedVariables { get; set; } = new List<string>();

        public bool IsEmpty => SelectedVariables == null || SelectedVariables.Count == 0;

        public static bool FileExists => File.Exists(FilePath);

        // ── 讀寫 ──────────────────────────────────────────────────────────

        public static ColumnSettings Load()
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    var json = File.ReadAllText(FilePath, Encoding.UTF8);
                    var names = ParseJsonStringArray(json);
                    if (names != null && names.Count > 0)
                        return new ColumnSettings { SelectedVariables = names };
                }
            }
            catch { }

            return new ColumnSettings();
        }

        public void Save()
        {
            try
            {
                File.WriteAllText(FilePath, BuildJsonStringArray(SelectedVariables), Encoding.UTF8);
            }
            catch { }
        }

        // ── 輕量 JSON 序列化（字串陣列，不依賴外部套件）────────────────────

        private static string BuildJsonStringArray(IEnumerable<string> values)
        {
            var sb = new StringBuilder("[");
            var first = true;
            foreach (var v in values ?? new List<string>())
            {
                if (!first) sb.Append(',');
                first = false;
                sb.Append('"');
                sb.Append((v ?? string.Empty)
                    .Replace("\\", "\\\\")
                    .Replace("\"", "\\\""));
                sb.Append('"');
            }
            sb.Append(']');
            return sb.ToString();
        }

        private static List<string> ParseJsonStringArray(string json)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(json)) return result;

            json = json.Trim();
            if (!json.StartsWith("[") || !json.EndsWith("]")) return result;

            json = json.Substring(1, json.Length - 2);

            var inString = false;
            var escape = false;
            var current = new StringBuilder();

            foreach (var c in json)
            {
                if (escape)
                {
                    current.Append(c);
                    escape = false;
                }
                else if (c == '\\')
                {
                    escape = true;
                }
                else if (c == '"')
                {
                    if (inString)
                    {
                        result.Add(current.ToString());
                        current.Clear();
                        inString = false;
                    }
                    else
                    {
                        inString = true;
                    }
                }
                else if (inString)
                {
                    current.Append(c);
                }
            }

            return result;
        }
    }
}
