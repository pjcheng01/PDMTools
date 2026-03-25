using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using PDMTools.Models;

namespace PDMTools.Services
{
    public static class PdmProjectSnapshotSerializer
    {
        private static readonly JsonSerializerOptions Options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        public static List<BomItem> CloneBomItems(IEnumerable<BomItem> items)
        {
            if (items == null)
            {
                return new List<BomItem>();
            }

            var list = new List<BomItem>();
            foreach (var x in items)
            {
                var c = CloneByJson(x);
                if (c != null)
                {
                    list.Add(c);
                }
            }

            return list;
        }

        public static List<ReferenceAuditRow> CloneReferenceAuditRows(IEnumerable<ReferenceAuditRow> rows)
        {
            if (rows == null)
            {
                return new List<ReferenceAuditRow>();
            }

            return rows.Select(r => CloneByJson(r)).Where(r => r != null).ToList();
        }

        private static T CloneByJson<T>(T obj)
            where T : class
        {
            if (obj == null)
            {
                return null;
            }

            return JsonSerializer.Deserialize<T>(JsonSerializer.Serialize(obj, Options), Options);
        }

        public static void Save(PdmProjectSnapshotDocument doc, string path)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            var json = JsonSerializer.Serialize(doc, Options);
            File.WriteAllText(path, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        public static PdmProjectSnapshotDocument Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("路徑不可為空。", nameof(path));
            }

            if (!File.Exists(path))
            {
                throw new FileNotFoundException("找不到快照檔。", path);
            }

            var json = File.ReadAllText(path, Encoding.UTF8);
            var doc = JsonSerializer.Deserialize<PdmProjectSnapshotDocument>(json, Options);
            if (doc == null)
            {
                throw new InvalidOperationException("無法解析快照檔案。");
            }

            if (doc.FormatVersion < 1)
            {
                throw new InvalidOperationException($"不支援的快照格式版本：{doc.FormatVersion}");
            }

            if (doc.SavedAtLocal == default && doc.SavedAtUtcLegacy.HasValue)
            {
                doc.SavedAtLocal = DateTime.SpecifyKind(doc.SavedAtUtcLegacy.Value, DateTimeKind.Utc).ToLocalTime();
            }

            return doc;
        }
    }
}
