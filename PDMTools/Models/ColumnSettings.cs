using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PDMTools.Models
{
    /// <summary>
    /// 使用者選擇要顯示的資料卡變數設定，持久化至 columns.json。
    /// JSON 格式：{"selected":[...],"known":[...]}
    /// 舊版（純陣列格式）可自動升級。
    /// </summary>
    public sealed class ColumnSettings
    {
        public static readonly string FilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "columns.json");

        /// <summary>使用者勾選的資料卡變數（顯示於表格 / Excel）。</summary>
        public List<string> SelectedVariables { get; set; } = new List<string>();

        /// <summary>
        /// 使用者勾選要顯示的固定欄位名稱清單。
        /// null 或空清單表示「尚未設定」，程式應預設全部顯示。
        /// </summary>
        public List<string> SelectedFixedColumns { get; set; } = null;

        /// <summary>
        /// 上次開啟設定視窗時 Vault 提供的完整變數清單。
        /// 用於啟動時偵測是否有新變數加入。
        /// </summary>
        public List<string> KnownVariables { get; set; } = new List<string>();

        /// <summary>使用者儲存的欄位組合清單。</summary>
        public List<ColumnPresetProfile> ColumnProfiles { get; set; } = new List<ColumnPresetProfile>();

        /// <summary>上次使用（套用）的欄位組合名稱。</summary>
        public string LastUsedProfileName { get; set; } = string.Empty;

        public bool IsEmpty => SelectedVariables == null || SelectedVariables.Count == 0;

        /// <summary>
        /// 固定欄位是否已設定。false 表示尚未設定，應預設全部顯示。
        /// </summary>
        public bool HasFixedColumnsSetting =>
            SelectedFixedColumns != null && SelectedFixedColumns.Count > 0;

        public static bool FileExists => File.Exists(FilePath);

        // ── 讀 ─────────────────────────────────────────────────────────────

        public static ColumnSettings Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new ColumnSettings();
                var json = File.ReadAllText(FilePath, Encoding.UTF8).Trim();
                if (string.IsNullOrEmpty(json)) return new ColumnSettings();

                // 舊格式：純陣列 ["a","b"]
                if (json.StartsWith("["))
                {
                    var legacy = ParseJsonStringArray(json);
                    return new ColumnSettings
                    {
                        SelectedVariables = legacy ?? new List<string>(),
                        KnownVariables    = new List<string>()   // 舊版沒有 known，視為未知
                    };
                }

                // 新格式：物件 {"selected":[...],"known":[...],"fixedColumns":[...]}
                var selected     = ExtractArray(json, "selected")     ?? new List<string>();
                var known        = ExtractArray(json, "known")        ?? new List<string>();
                var fixedColumns = ExtractArray(json, "fixedColumns"); // null = 尚未設定
                var profileNames = ExtractArray(json, "profileNames") ?? new List<string>();
                var profileData  = ExtractArray(json, "profileSelected") ?? new List<string>();
                var lastUsed     = ExtractString(json, "lastUsedProfile") ?? string.Empty;
                var profiles     = BuildProfiles(profileNames, profileData);
                return new ColumnSettings
                {
                    SelectedVariables  = selected,
                    KnownVariables     = known,
                    SelectedFixedColumns = fixedColumns,   // null 保留，讓程式預設全顯示
                    ColumnProfiles = profiles,
                    LastUsedProfileName = lastUsed
                };
            }
            catch
            {
                return new ColumnSettings();
            }
        }

        // ── 寫 ─────────────────────────────────────────────────────────────

        public void Save()
        {
            try
            {
                var json = "{\"selected\":"
                    + BuildJsonStringArray(SelectedVariables)
                    + ",\"known\":"
                    + BuildJsonStringArray(KnownVariables)
                    + ",\"fixedColumns\":"
                    + BuildJsonStringArray(SelectedFixedColumns ?? new List<string>())
                    + ",\"profileNames\":"
                    + BuildJsonStringArray((ColumnProfiles ?? new List<ColumnPresetProfile>()).Select(p => p?.Name ?? string.Empty))
                    + ",\"profileSelected\":"
                    + BuildJsonStringArray((ColumnProfiles ?? new List<ColumnPresetProfile>()).Select(p => JoinProfileColumns(p?.SelectedItems)))
                    + ",\"lastUsedProfile\":"
                    + BuildJsonString(LastUsedProfileName)
                    + "}";
                File.WriteAllText(FilePath, json, Encoding.UTF8);
            }
            catch { }
        }

        // ── 輕量 JSON 工具 ──────────────────────────────────────────────────

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

        private static string BuildJsonString(string value)
        {
            var safe = (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"");
            return "\"" + safe + "\"";
        }

        /// <summary>從 JSON 物件字串中找到指定 key 對應的字串陣列。</summary>
        private static List<string> ExtractArray(string json, string key)
        {
            var search  = "\"" + key + "\"";
            var keyIdx  = json.IndexOf(search, StringComparison.OrdinalIgnoreCase);
            if (keyIdx < 0) return null;

            var colonIdx = json.IndexOf(':', keyIdx + search.Length);
            if (colonIdx < 0) return null;

            var startIdx = json.IndexOf('[', colonIdx);
            if (startIdx < 0) return null;

            // 找對應的 ] (考慮巢狀——雖然我們的資料不會有，但保險起見)
            var depth  = 0;
            var endIdx = -1;
            for (var i = startIdx; i < json.Length; i++)
            {
                if      (json[i] == '[') depth++;
                else if (json[i] == ']') { depth--; if (depth == 0) { endIdx = i; break; } }
            }
            if (endIdx < 0) return null;

            return ParseJsonStringArray(json.Substring(startIdx, endIdx - startIdx + 1));
        }

        private static string ExtractString(string json, string key)
        {
            var search = "\"" + key + "\"";
            var keyIdx = json.IndexOf(search, StringComparison.OrdinalIgnoreCase);
            if (keyIdx < 0) return null;

            var colonIdx = json.IndexOf(':', keyIdx + search.Length);
            if (colonIdx < 0) return null;

            var quoteStart = json.IndexOf('"', colonIdx + 1);
            if (quoteStart < 0) return null;

            var sb = new StringBuilder();
            var escaped = false;
            for (var i = quoteStart + 1; i < json.Length; i++)
            {
                var c = json[i];
                if (escaped)
                {
                    sb.Append(c);
                    escaped = false;
                    continue;
                }

                if (c == '\\')
                {
                    escaped = true;
                    continue;
                }

                if (c == '"')
                {
                    return sb.ToString();
                }

                sb.Append(c);
            }

            return null;
        }

        private static List<string> ParseJsonStringArray(string json)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(json)) return result;

            json = json.Trim();
            if (!json.StartsWith("[") || !json.EndsWith("]")) return result;
            json = json.Substring(1, json.Length - 2);

            var inString = false;
            var escape   = false;
            var current  = new StringBuilder();

            foreach (var c in json)
            {
                if (escape)          { current.Append(c); escape = false; }
                else if (c == '\\')  { escape = true; }
                else if (c == '"')
                {
                    if (inString) { result.Add(current.ToString()); current.Clear(); inString = false; }
                    else          { inString = true; }
                }
                else if (inString)   { current.Append(c); }
            }
            return result;
        }

        private const char ProfileColumnSeparator = '\u001F';

        private static string JoinProfileColumns(IEnumerable<string> values)
        {
            return string.Join(ProfileColumnSeparator.ToString(), values ?? Enumerable.Empty<string>());
        }

        private static List<string> SplitProfileColumns(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return new List<string>();
            return raw.Split(new[] { ProfileColumnSeparator }, StringSplitOptions.None)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();
        }

        private static List<ColumnPresetProfile> BuildProfiles(IReadOnlyList<string> names, IReadOnlyList<string> selectedData)
        {
            var profiles = new List<ColumnPresetProfile>();
            var count = Math.Min(names?.Count ?? 0, selectedData?.Count ?? 0);
            for (var i = 0; i < count; i++)
            {
                var name = names[i]?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(name)) continue;

                profiles.Add(new ColumnPresetProfile
                {
                    Name = name,
                    SelectedItems = SplitProfileColumns(selectedData[i] ?? string.Empty)
                });
            }

            return profiles;
        }
    }
}
