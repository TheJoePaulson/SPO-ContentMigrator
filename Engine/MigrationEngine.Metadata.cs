using System.Collections;
using System.Globalization;
using Microsoft.SharePoint.Client;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;

namespace ContentMigrator.Engine;

/// <summary>A person value waiting to be resolved on the destination web (Web.EnsureUser).</summary>
internal sealed class UserReference
{
    public UserReference(string[] identities, bool multi)
    {
        Identities = identities;
        Multi = multi;
    }
    public string[] Identities { get; }
    public bool Multi { get; }
    public override string ToString() => string.Join(";", Identities);
}

public sealed partial class MigrationEngine
{
    private readonly Dictionary<string, List<FieldMapEntry>> _fieldMapCache = new(StringComparer.OrdinalIgnoreCase);

    // ======================================================================= field map

    /// <summary>
    /// Source fields that also exist (by InternalName) on the destination,
    /// excluding hidden, read-only, system and "_" fields. Computed once per list pair per run
    /// (the script recomputed it for every file).
    /// </summary>
    private List<FieldMapEntry> GetWritableFieldMap(SpConnection src, SpListInfo srcList, SpConnection dst, SpListInfo dstList)
    {
        var key = $"{src.SiteUrl}|{srcList.Id}|{dst.SiteUrl}|{dstList.Id}";
        if (_fieldMapCache.TryGetValue(key, out var cached)) return cached;

        var sourceFields = GetFields(src, srcList.List);
        var destByName = GetFields(dst, dstList.List)
            .GroupBy(f => f.InternalName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var map = new List<FieldMapEntry>();
        foreach (var sf in sourceFields)
        {
            if (sf.Hidden || sf.ReadOnlyField) continue;
            if (_systemFields.Contains(sf.InternalName)) continue;
            if (sf.InternalName.StartsWith("_")) continue;
            if (!destByName.TryGetValue(sf.InternalName, out var df)) continue;
            if (df.Hidden || df.ReadOnlyField) continue;
            map.Add(new FieldMapEntry
            {
                InternalName = sf.InternalName,
                DisplayName = sf.Title,
                SourceTypeAsString = sf.TypeAsString,
                DestTypeAsString = df.TypeAsString
            });
        }
        _fieldMapCache[key] = map;
        return map;
    }

    // ======================================================================= compatibility report

    /// <summary>Field compatibility report -> Reports\MetadataCompatibility_&lt;context&gt;.csv</summary>
    private void WriteFieldCompatibilityReport(SpConnection src, SpListInfo srcList, SpConnection dst, SpListInfo dstList,
        string contextLabel)
    {
        if (!_s.Reports.EnableMetadataCompatibilityReport) return;
        Log($"Generating metadata compatibility report for '{srcList.Title}' -> '{dstList.Title}'...");

        var sourceFields = GetFields(src, srcList.List);
        var destByName = GetFields(dst, dstList.List)
            .GroupBy(f => f.InternalName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var report = new List<FieldCompatibilityRecord>();
        foreach (var sf in sourceFields)
        {
            string status = "Supported", notes = "";
            destByName.TryGetValue(sf.InternalName, out var df);
            if (sf.Hidden) { status = "Skipped"; notes = "Source field is hidden."; }
            else if (sf.ReadOnlyField) { status = "Skipped"; notes = "Source field is read-only."; }
            else if (_systemFields.Contains(sf.InternalName)) { status = "Skipped"; notes = "System field handled separately or intentionally excluded."; }
            else if (sf.InternalName.StartsWith("_")) { status = "Skipped"; notes = "Internal/system-style field excluded."; }
            else if (df == null) { status = "MissingInDestination"; notes = "Destination does not contain a matching internal field name."; }
            else if (df.Hidden || df.ReadOnlyField) { status = "DestinationNotWritable"; notes = "Destination field exists but is hidden or read-only."; }
            else if (!string.Equals(sf.TypeAsString, df.TypeAsString, StringComparison.Ordinal))
            {
                status = "TypeMismatch";
                notes = $"Source type '{sf.TypeAsString}' does not match destination type '{df.TypeAsString}'.";
            }
            else if (!_supportedTypes.Contains(sf.TypeAsString)) { status = "UnsupportedInV1"; notes = "Field type is not currently included in metadata replay."; }

            report.Add(new FieldCompatibilityRecord
            {
                ContextLabel = contextLabel,
                FieldInternalName = sf.InternalName,
                FieldDisplayName = sf.Title,
                SourceTypeAsString = sf.TypeAsString,
                DestinationTypeAsString = df?.TypeAsString ?? "",
                Status = status,
                Notes = notes
            });
        }

        var path = _run.ReportPath($"MetadataCompatibility_{RunContext.SafeFileName(contextLabel)}.csv");
        Csv.Write(path, report);
        Log($"Metadata compatibility report exported: {path}", MigratorLogLevel.Success);
        Log($"Supported: {report.Count(r => r.Status == "Supported")} | Missing: {report.Count(r => r.Status == "MissingInDestination")} | " +
            $"UnsupportedInV1: {report.Count(r => r.Status == "UnsupportedInV1")} | TypeMismatch: {report.Count(r => r.Status == "TypeMismatch")} | " +
            $"Skipped: {report.Count(r => r.Status == "Skipped")}");
    }

    // ======================================================================= value conversion

    /// <summary>Converts a field value for copying. Person values become <see cref="UserReference"/>s.</summary>
    private object ConvertFieldValue(object value, string fieldType)
    {
        if (value == null) return null;
        switch (fieldType)
        {
            case "Text":
            case "Note":
            case "Choice":
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            case "MultiChoice":
                if (value is string[] arr) return arr;
                if (value is IEnumerable e and not string)
                    return e.Cast<object>().Select(o => Convert.ToString(o, CultureInfo.InvariantCulture))
                        .Where(s => !string.IsNullOrEmpty(s)).ToArray();
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            case "Number":
            case "Integer": // v2.2.8: Integer is treated exactly like Number
            case "Currency":
            case "DateTime":
                return value;
            case "Boolean":
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            case "URL":
                return value is FieldUrlValue u ? new FieldUrlValue { Url = u.Url, Description = u.Description } : value;
            case "User":
            {
                var id = UserIdentity(value);
                return string.IsNullOrWhiteSpace(id) ? null : new UserReference(new[] { id }, false);
            }
            case "UserMulti":
            {
                var ids = (value is IEnumerable many and not string ? many.Cast<object>() : new[] { value })
                    .Select(UserIdentity).Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();
                return ids.Length == 0 ? null : new UserReference(ids, true);
            }
            default:
                return null;
        }
    }

    private static string UserIdentity(object value)
    {
        if (value is FieldUserValue fu)
        {
            if (!string.IsNullOrWhiteSpace(fu.Email)) return fu.Email;
            return fu.LookupValue;
        }
        return value?.ToString();
    }

    private static bool IsEmptyValue(object v) =>
        v == null || (v is string s && s.Length == 0) || (v is Array a && a.Length == 0);

    /// <summary>Readable form of a field value for MetadataLog.csv.</summary>
    private static string Describe(object value)
    {
        switch (value)
        {
            case null: return "";
            case string s: return s;
            case FieldUserValue u: return !string.IsNullOrWhiteSpace(u.Email) ? u.Email : u.LookupValue;
            case FieldLookupValue l: return l.LookupValue;
            case FieldUrlValue url: return url.Url;
            case DateTime dt: return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            case IEnumerable e: return string.Join(";", e.Cast<object>().Select(Describe));
            case IFormattable f: return f.ToString(null, CultureInfo.InvariantCulture);
            default: return value.ToString();
        }
    }

    /// <summary>
    /// Resolves every <see cref="UserReference"/> against the destination web BEFORE any item
    /// write is queued, so EnsureUser round trips never interleave with a pending update.
    /// Unresolvable people are logged to UnresolvedUsers.csv and left out.
    /// </summary>
    private Dictionary<string, object> ResolveValuesForWrite(SpConnection dst, Dictionary<string, object> values, string sourceRef)
    {
        var resolved = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in values)
        {
            if (kv.Value is not UserReference ur)
            {
                resolved[kv.Key] = kv.Value;
                continue;
            }

            var users = new List<FieldUserValue>();
            foreach (var identity in ur.Identities)
            {
                var id = dst.ResolveUserId(identity, out var error);
                if (id.HasValue) users.Add(new FieldUserValue { LookupId = id.Value });
                else _run.AddUnresolvedUserLog(sourceRef, kv.Key, identity, error);
            }
            if (users.Count == 0) continue;
            resolved[kv.Key] = ur.Multi ? users.ToArray() : users[0];
        }
        return resolved;
    }

    /// <summary>Converted values for one source item, with MetadataLog rows.</summary>
    private (Dictionary<string, object> Values, int Copied, int Skipped) BuildMetadataValues(ListItem sourceItem,
        List<FieldMapEntry> fieldMap, string sourceRef, string destinationRef, bool logCopied, bool logSkipped)
    {
        var values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        int copied = 0, skipped = 0;
        var fv = sourceItem.FieldValues;
        foreach (var field in fieldMap)
        {
            if (!_supportedTypes.Contains(field.SourceTypeAsString))
            {
                skipped++;
                if (logSkipped)
                {
                    fv.TryGetValue(field.InternalName, out var raw);
                    _run.AddMetadataLog(sourceRef, destinationRef, field.InternalName, field.DisplayName,
                        field.SourceTypeAsString, Describe(raw), "", "SkippedUnsupported", "Skipped",
                        "Unsupported field type. Value logged but not applied.");
                }
                continue;
            }
            if (!fv.TryGetValue(field.InternalName, out var sourceValue)) continue;

            var converted = ConvertFieldValue(sourceValue, field.SourceTypeAsString);
            if (IsEmptyValue(converted)) continue;

            values[field.InternalName] = converted;
            copied++;
            if (logCopied)
            {
                _run.AddMetadataLog(sourceRef, destinationRef, field.InternalName, field.DisplayName,
                    field.SourceTypeAsString, Describe(sourceValue), Describe(converted), "Copied", "Success", "");
            }
        }
        return (values, copied, skipped);
    }

    // ======================================================================= system metadata

    /// <summary>
    /// Replays Created, Modified, Author, Editor with UpdateOverwriteVersion
    /// (no new version). Returns Applied | NoSystemValues | Skipped | Error.
    /// </summary>
    private string ApplySystemMetadata(SpConnection dst, ListItem destinationItem, ListItem sourceItem,
        string sourceRef, string destinationRef)
    {
        if (!_opt.PreserveSystemMetadata) return "Skipped";
        try
        {
            var fv = sourceItem.FieldValues;
            var values = new Dictionary<string, object>();
            if (fv.TryGetValue("Created", out var created) && created != null) values["Created"] = created;
            if (fv.TryGetValue("Modified", out var modified) && modified != null) values["Modified"] = modified;
            foreach (var name in new[] { "Author", "Editor" })
            {
                if (fv.TryGetValue(name, out var v) && v is FieldUserValue u && !string.IsNullOrWhiteSpace(u.Email))
                {
                    values[name] = new UserReference(new[] { u.Email }, false);
                }
            }
            if (values.Count == 0) return "NoSystemValues";

            var resolved = ResolveValuesForWrite(dst, values, sourceRef);
            foreach (var kv in resolved) destinationItem[kv.Key] = kv.Value;
            destinationItem.UpdateOverwriteVersion();
            dst.ExecuteQuery();

            _run.AddMetadataLog(sourceRef, destinationRef, "Created/Modified/Author/Editor", "System Metadata", "System",
                "System metadata from source item", "Applied to destination item", "SystemMetadataApplied", "Success", "");
            return "Applied";
        }
        catch (Exception ex)
        {
            _run.AddMetadataLog(sourceRef, destinationRef, "Created/Modified/Author/Editor", "System Metadata", "System",
                "System metadata from source item", "", "SystemMetadataApplyFailed", "Error", ex.Message);
            _run.AddErrorLog("Metadata", "SetSystemMetadata", sourceRef, destinationRef, ex);
            return "Error";
        }
    }

    // ======================================================================= file metadata

    /// <summary>Column values, then system metadata, onto a copied/moved file.</summary>
    private string ApplyFileMetadata(SpConnection src, SpListInfo srcLib, SpConnection dst, SpListInfo dstLib,
        string sourceUrl, string destinationUrl, ListItem preloadedSourceItem)
    {
        try
        {
            var sourceItem = preloadedSourceItem;
            if (sourceItem == null)
            {
                sourceItem = src.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(sourceUrl)).ListItemAllFields;
                src.Context.Load(sourceItem);
                src.ExecuteQuery();
            }

            var fieldMap = GetWritableFieldMap(src, srcLib, dst, dstLib);
            var (values, copied, skipped) = BuildMetadataValues(sourceItem, fieldMap, sourceUrl, destinationUrl,
                logCopied: true, logSkipped: true);

            var resolved = values.Count > 0 ? ResolveValuesForWrite(dst, values, sourceUrl) : values;

            var destinationItem = dst.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(destinationUrl)).ListItemAllFields;
            if (resolved.Count > 0)
            {
                foreach (var kv in resolved) destinationItem[kv.Key] = kv.Value;
                destinationItem.UpdateOverwriteVersion();
                dst.ExecuteQuery();
            }

            var systemStatus = ApplySystemMetadata(dst, destinationItem, sourceItem, sourceUrl, destinationUrl);
            return $"MetadataCopied:{copied};Skipped:{skipped};System:{systemStatus}";
        }
        catch (Exception ex)
        {
            _run.AddErrorLog("Metadata", "ApplyFileMetadata", sourceUrl, destinationUrl, ex);
            return "MetadataError";
        }
    }
}
