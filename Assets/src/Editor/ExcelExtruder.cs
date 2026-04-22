using UnityEngine;
using Sylvan.Data;
using Sylvan.Data.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using MemoryPack;

namespace ExcelExtruder
{
    public class ExcelSerialize
    {
        protected virtual string EXCEL_SKIP => "#";
        protected virtual string EXCEL_FIELDNAME => "@";
        protected virtual string EXCELRES_PATH => "./Documents/Excel/";
        protected virtual string CSV_PATH => "./Documents/CSV/";
        protected virtual string BIN_PATH => Application.dataPath + "/Resources/StaticData/";
        protected virtual string CONFIG_PATH => "./excelconfig";

        private TypeConvert m_typeConvert;
        private DataValidator m_validator;
        private Action<string, float, string, string> _EVENT_PROGRESS;
        private Action _EVENT_END_PROGRESS;
        private Action<string> _EVENT_LOG;
        private Action<string> _EVENT_ERROR_LOG;

        private void Progress(float progress, string action, string name) => _EVENT_PROGRESS?.Invoke("Excels 序列化", progress, action, name);
        private void EndProgress() => _EVENT_END_PROGRESS?.Invoke();
        private void LogError(string error) => _EVENT_ERROR_LOG?.Invoke(error);
        private void Log(string log) => _EVENT_LOG?.Invoke(log);

        public void Init(TypeConvert typeConvert, Action EVENT_END_PROGRESS,
            Action<string, float, string, string> EVENT_PROGRESS,
            Action<string> EVENT_ERROR_LOG,
            Action<string> EVENT_LOG)
        {
            Assembly assembly = null;
            var asmPath = Path.Combine(".", "Library", "ScriptAssemblies", "Assembly-CSharp.dll");
            if (File.Exists(asmPath))
                assembly = Assembly.LoadFile(Path.GetFullPath(asmPath));

            m_typeConvert = typeConvert;
            m_typeConvert.Init(assembly);
            m_validator = new DataValidator();

            if (!Directory.Exists(EXCELRES_PATH)) Directory.CreateDirectory(EXCELRES_PATH);
            if (!Directory.Exists(CSV_PATH)) Directory.CreateDirectory(CSV_PATH);
            if (!Directory.Exists(BIN_PATH)) Directory.CreateDirectory(BIN_PATH);

            _EVENT_PROGRESS = EVENT_PROGRESS;
            _EVENT_END_PROGRESS = EVENT_END_PROGRESS;
            _EVENT_LOG = EVENT_LOG;
            _EVENT_ERROR_LOG = EVENT_ERROR_LOG;
        }

        /// <summary>
        /// 加载已有的 Excel 配置（含哈希信息）
        /// </summary>
        private ExcelConfig LoadExcelConfig()
        {
            if (!File.Exists(CONFIG_PATH))
                return new ExcelConfig();

            try
            {
                var bin = File.ReadAllBytes(CONFIG_PATH);
                var config = MemoryPackSerializer.Deserialize<ExcelConfig>(bin);
                return NormalizeConfig(config);
            }
            catch
            {
                Log("配置文件格式不兼容，将执行全量重建");
                return new ExcelConfig();
            }
        }

        /// <summary>
        /// 保存 Excel 配置
        /// </summary>
        private void SaveExcelConfig(ExcelConfig config)
        {
            var bin = MemoryPackSerializer.Serialize(config);
            File.WriteAllBytes(CONFIG_PATH, bin);
        }

        /// <summary>
        /// 计算文件的 MD5 哈希值
        /// </summary>
        private static string ComputeFileHash(string filePath)
        {
            using var md5 = MD5.Create();
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var hash = md5.ComputeHash(stream);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }

        /// <summary>
        /// 序列化所有 Excel 文件（支持增量模式）
        /// </summary>
        /// <param name="forceAll">true 时强制全量重建，忽略哈希缓存</param>
        public void SerializeAllExcel(bool forceAll = false)
        {
            try
            {
                Progress(0, "SerializeExcels", "Start");

                var existingConfig = forceAll ? new ExcelConfig() : LoadExcelConfig();
                var nextConfig = new ExcelConfig();
                var pendingCommits = new List<ExcelConvertResult>();
                var discoveredFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var files = GetExcelFiles();
                var processedCount = 0;
                var skippedCount = 0;
                var failedCount = 0;
                var count = files.Length;

                for (int index = 0; index < count; index++)
                {
                    var file = files[index];
                    if (file.Name.StartsWith("~$"))
                        continue;

                    var relativePath = GetWorkbookRelativePath(file);
                    discoveredFiles.Add(relativePath);
                    Progress((float)index / Math.Max(1, count), "SerializeExcel", relativePath);

                    var currentHash = ComputeFileHash(file.FullName);
                    existingConfig.Files.TryGetValue(relativePath, out var existingInfo);

                    if (!forceAll
                        && existingInfo != null
                        && string.Equals(existingInfo.Hash, currentHash, StringComparison.OrdinalIgnoreCase)
                        && OutputsExist(existingInfo))
                    {
                        nextConfig.Files[relativePath] = CloneInfo(existingInfo);
                        skippedCount++;
                        Log($"[增量跳过] {relativePath} 未变更");
                        continue;
                    }

                    var result = ConvertOneExcel(relativePath);
                    if (result.Success)
                    {
                        pendingCommits.Add(result);
                        nextConfig.Files[relativePath] = CreateExcelFileInfo(result, currentHash);
                        processedCount++;
                    }
                    else
                    {
                        failedCount++;
                        if (existingInfo != null)
                        {
                            nextConfig.Files[relativePath] = CloneInfo(existingInfo);
                            Log($"[保留旧产物] {relativePath} 导出失败，继续使用上一版 bytes");
                        }
                    }
                }

                var duplicateErrors = ValidateDuplicateSheets(nextConfig);
                if (duplicateErrors.Count > 0)
                {
                    foreach (var error in duplicateErrors)
                        LogError(error);

                    LogError("检测到重复 Sheet 名称，已取消本次提交，配置和产物保持不变");
                    EndProgress();
                    return;
                }

                CleanupRemovedOutputs(existingConfig, nextConfig, discoveredFiles);
                CommitOutputs(pendingCommits);
                SaveExcelConfig(nextConfig);

                Log($"序列化完成: 成功处理 {processedCount} 个文件, 跳过 {skippedCount} 个文件, 失败 {failedCount} 个文件");
                Progress(1, "SerializeExcels", "End");
            }
            catch (Exception e)
            {
                EndProgress();
                Debug.LogException(e);
            }
        }

        private FileInfo[] GetExcelFiles()
        {
            var folder = new DirectoryInfo(EXCELRES_PATH);
            if (!folder.Exists)
                return Array.Empty<FileInfo>();

            var files = folder.GetFiles("*.xlsx", SearchOption.AllDirectories);
            Array.Sort(files, (left, right) => string.CompareOrdinal(left.FullName, right.FullName));
            return files;
        }

        private ExcelConfig NormalizeConfig(ExcelConfig config)
        {
            var normalized = new ExcelConfig();
            if (config?.Files == null)
                return normalized;

            foreach (var pair in config.Files)
            {
                if (string.IsNullOrWhiteSpace(pair.Key) || pair.Value == null)
                    continue;

                normalized.Files[NormalizeRelativePath(pair.Key)] = CloneInfo(pair.Value);
            }

            return normalized;
        }

        private ExcelFileInfo CloneInfo(ExcelFileInfo info)
        {
            var clone = new ExcelFileInfo
            {
                Hash = info?.Hash,
                Sheets = new List<ExcelSheetInfo>()
            };

            if (info?.Sheets == null)
                return clone;

            foreach (var sheet in info.Sheets)
            {
                if (sheet == null || string.IsNullOrWhiteSpace(sheet.SheetName))
                    continue;

                clone.Sheets.Add(new ExcelSheetInfo
                {
                    SheetName = sheet.SheetName,
                    TypeName = string.IsNullOrWhiteSpace(sheet.TypeName) ? sheet.SheetName : sheet.TypeName
                });
            }

            return clone;
        }

        private string GetWorkbookRelativePath(FileInfo file)
        {
            var root = Path.GetFullPath(EXCELRES_PATH);
            return NormalizeRelativePath(Path.GetRelativePath(root, file.FullName));
        }

        private static string NormalizeRelativePath(string path)
        {
            return path.Replace('\\', '/');
        }

        private string GetWorkbookFullPath(string relativePath)
        {
            return Path.GetFullPath(Path.Combine(EXCELRES_PATH, relativePath.Replace('/', Path.DirectorySeparatorChar)));
        }

        private string GetCsvOutputPath(string sheetName)
        {
            return Path.Combine(CSV_PATH, sheetName + ".csv");
        }

        private string GetBinOutputPath(string sheetName)
        {
            return Path.Combine(BIN_PATH, sheetName + ".bytes");
        }

        private bool OutputsExist(ExcelFileInfo info)
        {
            if (info?.Sheets == null || info.Sheets.Count == 0)
                return false;

            foreach (var sheet in info.Sheets)
            {
                if (sheet == null || string.IsNullOrWhiteSpace(sheet.SheetName))
                    return false;

                if (!File.Exists(GetCsvOutputPath(sheet.SheetName)) || !File.Exists(GetBinOutputPath(sheet.SheetName)))
                    return false;
            }

            return true;
        }

        private ExcelFileInfo CreateExcelFileInfo(ExcelConvertResult result, string hash)
        {
            var info = new ExcelFileInfo
            {
                Hash = hash,
                Sheets = new List<ExcelSheetInfo>()
            };

            foreach (var sheet in result.Sheets)
            {
                info.Sheets.Add(new ExcelSheetInfo
                {
                    SheetName = sheet.SheetName,
                    TypeName = sheet.TypeName
                });
            }

            return info;
        }

        private List<string> ValidateDuplicateSheets(ExcelConfig config)
        {
            var errors = new List<string>();
            var owners = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (var file in config.Files)
            {
                if (file.Value?.Sheets == null)
                    continue;

                foreach (var sheet in file.Value.Sheets)
                {
                    if (sheet == null || string.IsNullOrWhiteSpace(sheet.SheetName))
                        continue;

                    if (!owners.TryAdd(sheet.SheetName, file.Key))
                    {
                        errors.Add($"[重复Sheet] {sheet.SheetName} 同时存在于 {owners[sheet.SheetName]} 和 {file.Key}");
                    }
                }
            }

            return errors;
        }

        private void CleanupRemovedOutputs(ExcelConfig previousConfig, ExcelConfig nextConfig, HashSet<string> discoveredFiles)
        {
            foreach (var pair in previousConfig.Files)
            {
                if (!discoveredFiles.Contains(pair.Key))
                {
                    DeleteOutputs(pair.Value);
                    continue;
                }

                if (!nextConfig.Files.TryGetValue(pair.Key, out var nextInfo))
                    continue;

                DeleteRemovedSheetOutputs(pair.Value, nextInfo);
            }
        }

        private void CommitOutputs(List<ExcelConvertResult> results)
        {
            foreach (var result in results)
            {
                foreach (var sheet in result.Sheets)
                {
                    File.WriteAllText(GetCsvOutputPath(sheet.SheetName), sheet.CsvContent);
                    File.WriteAllBytes(GetBinOutputPath(sheet.SheetName), sheet.BinaryData);
                }
            }
        }

        private void DeleteOutputs(ExcelFileInfo info)
        {
            if (info?.Sheets == null)
                return;

            foreach (var sheet in info.Sheets)
                DeleteSheetOutputs(sheet);
        }

        private void DeleteRemovedSheetOutputs(ExcelFileInfo previousInfo, ExcelFileInfo nextInfo)
        {
            if (previousInfo?.Sheets == null)
                return;

            var nextNames = new HashSet<string>(StringComparer.Ordinal);
            if (nextInfo?.Sheets != null)
            {
                foreach (var sheet in nextInfo.Sheets)
                    nextNames.Add(sheet.SheetName);
            }

            foreach (var sheet in previousInfo.Sheets)
            {
                if (!nextNames.Contains(sheet.SheetName))
                    DeleteSheetOutputs(sheet);
            }
        }

        private void DeleteSheetOutputs(ExcelSheetInfo sheet)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(sheet.SheetName))
                return;

            DeleteIfExists(GetCsvOutputPath(sheet.SheetName));
            DeleteIfExists(GetBinOutputPath(sheet.SheetName));
        }

        private static void DeleteIfExists(string path)
        {
            if (File.Exists(path))
                File.Delete(path);
        }

        /// <summary>
        /// 单个 Excel 文件的序列化结果
        /// </summary>
        public sealed class ExcelConvertResult
        {
            public string WorkbookPath;
            public bool Success = true;
            public List<SheetConvertResult> Sheets = new List<SheetConvertResult>();
        }

        public sealed class SheetConvertResult
        {
            public string SheetName;
            public string TypeName;
            public string CsvContent;
            public byte[] BinaryData;
        }

        private sealed class ParsedRowData
        {
            public int RowIndex;
            public object Data;
        }

        public ExcelConvertResult ConvertOneExcel(string workbookRelativePath)
        {
            var workbookPath = GetWorkbookFullPath(workbookRelativePath);
            var result = new ExcelConvertResult { WorkbookPath = workbookRelativePath };
            var workbookType = ExcelDataReader.GetWorkbookType(workbookPath);

            using var fs = new FileInfo(workbookPath).Open(FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var edr = ExcelDataReader.Create(fs, workbookType, new ExcelDataReaderOptions
            {
                Schema = ExcelSchema.NoHeaders,
                Culture = CultureInfo.InvariantCulture,
            });

            do
            {
                var sheetName = edr.WorksheetName;
                m_validator.ClearContext();

                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    LogError($"[{workbookRelativePath}] 存在未命名 Sheet，已取消该文件导出");
                    result.Success = false;
                    continue;
                }

                var classType = m_typeConvert.TryGetType(sheetName);
                if (classType == null)
                {
                    LogError($"[{workbookRelativePath}/{sheetName}] 找不到对应的数据类型，Sheet 名需要与类名一致或唯一匹配");
                    result.Success = false;
                    continue;
                }

                var genericList = TypeConvert.CreateGeneric(typeof(List<>), classType);
                var list = (IList)genericList;
                var heads = new List<int>();
                var headNames = new List<string>();
                var csvBuilder = new StringBuilder();
                var validationErrors = new List<ValidationError>();
                var diffWarnings = new List<ValidationError>();
                var parseErrors = new List<string>();
                Dictionary<string, object> oldDataMap = null;
                var newDataMap = new Dictionary<string, ParsedRowData>(StringComparer.Ordinal);
                FieldInfo primaryKeyField = null;
                var headerFound = false;
                var dataRowIndex = 0;
                var sheetHasFatalError = false;

                while (edr.Read())
                {
                    var firstCell = GetCellString(edr, 0);
                    if (firstCell == EXCEL_SKIP)
                        continue;

                    if (firstCell == EXCEL_FIELDNAME)
                    {
                        if (headerFound)
                        {
                            parseErrors.Add($"[{workbookRelativePath}/{sheetName}] 出现了重复的表头定义行 '@'");
                            sheetHasFatalError = true;
                            continue;
                        }

                        headerFound = true;
                        if (!TryReadHeader(classType, edr, workbookRelativePath, sheetName, heads, headNames))
                            sheetHasFatalError = true;

                        var headerRow = edr.Select(heads.ToArray());
                        AppendCsvRow(csvBuilder, headerRow);

                        if (headNames.Count == 0)
                            continue;

                        CheckNoFormulaFields(workbookPath, sheetName, classType, heads, headNames, validationErrors);

                        if (!TryResolvePrimaryKeyField(sheetName, classType, headNames, out primaryKeyField))
                            sheetHasFatalError = true;

                        if (primaryKeyField != null)
                            oldDataMap = LoadBaselineFromBin(sheetName, classType, primaryKeyField);

                        continue;
                    }

                    if (!headerFound)
                        continue;

                    var row = edr.Select(heads.ToArray());
                    AppendCsvRow(csvBuilder, row);
                    dataRowIndex++;

                    var rawValues = new Dictionary<string, string>(StringComparer.Ordinal);
                    for (int i = 0; i < row.FieldCount && i < headNames.Count; i++)
                        rawValues[headNames[i]] = GetCellString(row, i);

                    if (!TryConvertToObject(classType, row, headNames, sheetName, dataRowIndex, out var obj))
                    {
                        sheetHasFatalError = true;
                        continue;
                    }

                    validationErrors.AddRange(m_validator.Validate(obj, classType, sheetName, dataRowIndex, rawValues));

                    if (primaryKeyField != null)
                    {
                        var pk = GetPrimaryKeyValue(primaryKeyField, obj);
                        if (!string.IsNullOrEmpty(pk))
                        {
                            if (newDataMap.ContainsKey(pk))
                            {
                                parseErrors.Add($"[{workbookRelativePath}/{sheetName}] 第{dataRowIndex}行主键重复: {primaryKeyField.Name} = {pk}");
                                sheetHasFatalError = true;
                            }
                            else
                            {
                                newDataMap[pk] = new ParsedRowData
                                {
                                    RowIndex = dataRowIndex,
                                    Data = obj
                                };

                                if (oldDataMap != null && !oldDataMap.ContainsKey(pk))
                                    validationErrors.AddRange(m_validator.ValidateNewRow(classType, sheetName, dataRowIndex, obj, oldDataMap));
                            }
                        }
                    }

                    list.Add(obj);
                }

                if (!headerFound)
                {
                    parseErrors.Add($"[{workbookRelativePath}/{sheetName}] 未找到表头定义行 '@'");
                    sheetHasFatalError = true;
                }

                diffWarnings.AddRange(PerformDiffValidation(sheetName, classType, oldDataMap, newDataMap));

                // foreach (var parseError in parseErrors)
                //     LogError(parseError);

                if (validationErrors.Count > 0)
                {
                    LogError($"[校验报告] {sheetName} 共发现 {validationErrors.Count} 个阻断错误:");
                    foreach (var error in validationErrors)
                        LogError(error.ToString());

                    sheetHasFatalError = true;
                }

                if (diffWarnings.Count > 0)
                {
                    Log($"[Diff提示] {sheetName} 共发现 {diffWarnings.Count} 个变更提示:");
                    foreach (var warning in diffWarnings)
                        Log(warning.ToString());
                }

                if (sheetHasFatalError)
                {
                    result.Success = false;
                    continue;
                }

                result.Sheets.Add(new SheetConvertResult
                {
                    SheetName = sheetName,
                    TypeName = classType.FullName ?? classType.Name,
                    CsvContent = csvBuilder.ToString(),
                    BinaryData = SerializeToBytes(genericList)
                });
            } while (edr.NextResult());

            return result;
        }

        private bool TryReadHeader(Type classType, DbDataReader row, string workbookPath, string sheetName, List<int> heads, List<string> headNames)
        {
            heads.Clear();
            headNames.Clear();

            var duplicates = new HashSet<string>(StringComparer.Ordinal);
            var success = true;

            for (int i = 0; i < row.FieldCount; i++)
            {
                var fieldName = GetCellString(row, i);
                if (string.IsNullOrEmpty(fieldName)
                    || fieldName == EXCEL_SKIP
                    || fieldName == EXCEL_FIELDNAME)
                    continue;

                if (!duplicates.Add(fieldName))
                {
                    LogError($"[{workbookPath}/{sheetName}] 表头字段重复: {fieldName}");
                    success = false;
                    continue;
                }

                if (GetPublicInstanceField(classType, fieldName) == null)
                {
                    LogError($"[{workbookPath}/{sheetName}] 找不到字段: {fieldName}，对应类型 {classType.FullName}");
                    success = false;
                    continue;
                }

                heads.Add(i);
                headNames.Add(fieldName);
            }

            if (headNames.Count == 0)
            {
                LogError($"[{workbookPath}/{sheetName}] 未解析到任何可导出字段");
                success = false;
            }

            return success;
        }

        private bool TryResolvePrimaryKeyField(string sheetName, Type classType, List<string> headNames, out FieldInfo primaryKeyField)
        {
            primaryKeyField = null;
            if (headNames == null || headNames.Count == 0)
            {
                LogError($"[{sheetName}] 表头为空，无法解析主键字段");
                return false;
            }

            var markedFields = new List<FieldInfo>();
            var uniqueFields = new List<FieldInfo>();

            foreach (var headName in headNames)
            {
                var field = GetPublicInstanceField(classType, headName);
                if (field == null)
                    continue;

                if (field.GetCustomAttribute<PrimaryKeyAttribute>() != null)
                    markedFields.Add(field);

                if (field.GetCustomAttribute<UniqueAttribute>() != null)
                    uniqueFields.Add(field);
            }

            if (markedFields.Count > 1)
            {
                LogError($"[{sheetName}] 存在多个 [PrimaryKey] 字段，请只保留一个");
                return false;
            }

            if (markedFields.Count == 1)
            {
                primaryKeyField = markedFields[0];
                return true;
            }

            var firstField = GetPublicInstanceField(classType, headNames[0]);
            if (firstField != null && firstField.GetCustomAttribute<UniqueAttribute>() != null)
            {
                primaryKeyField = firstField;
                Log($"[{sheetName}] 未标记 [PrimaryKey]，回退使用第一列唯一字段 {primaryKeyField.Name}");
                return true;
            }

            if (uniqueFields.Count == 1)
            {
                primaryKeyField = uniqueFields[0];
                Log($"[{sheetName}] 未标记 [PrimaryKey]，回退使用唯一字段 {primaryKeyField.Name}");
                return true;
            }

            if (uniqueFields.Count > 1)
                Log($"[{sheetName}] 存在多个 [Unique] 字段且未标记 [PrimaryKey]，跳过 Diff/新增校验");
            else
                Log($"[{sheetName}] 未找到主键字段，跳过 Diff/新增校验");

            return true;
        }

        private Dictionary<string, object> LoadBaselineFromBin(string sheetName, Type classType, FieldInfo pkField)
        {
            var dataMap = new Dictionary<string, object>(StringComparer.Ordinal);
            var binPath = GetBinOutputPath(sheetName);
            if (!File.Exists(binPath))
                return dataMap;

            try
            {
                var bin = File.ReadAllBytes(binPath);
                var listType = typeof(List<>).MakeGenericType(classType);
                var oldList = MemoryPackSerializer.Deserialize(listType, bin) as IList;
                if (oldList == null)
                    return dataMap;

                foreach (var item in oldList)
                {
                    if (item == null)
                        continue;

                    var pkValue = GetPrimaryKeyValue(pkField, item);
                    if (!string.IsNullOrEmpty(pkValue))
                        dataMap[pkValue] = item;
                }
            }
            catch (Exception e)
            {
                Log($"加载基线数据 {sheetName}.bytes 失败: {e.Message}");
            }

            return dataMap;
        }

        private List<ValidationError> PerformDiffValidation(
            string sheetName,
            Type classType,
            Dictionary<string, object> oldDataMap,
            Dictionary<string, ParsedRowData> newDataMap)
        {
            var warnings = new List<ValidationError>();
            if (oldDataMap == null || oldDataMap.Count == 0 || newDataMap.Count == 0)
                return warnings;

            var fields = classType.GetFields(BindingFlags.Public | BindingFlags.Instance);
            foreach (var newKvp in newDataMap)
            {
                if (!oldDataMap.TryGetValue(newKvp.Key, out var oldObj))
                    continue;

                foreach (var field in fields)
                {
                    var oldValue = field.GetValue(oldObj);
                    var newValue = field.GetValue(newKvp.Value.Data);
                    var diffErrors = m_validator.ValidateDiff(classType, field.Name, oldValue, newValue);
                    if (diffErrors == null)
                        continue;

                    foreach (var errorMsg in diffErrors)
                    {
                        warnings.Add(new ValidationError
                        {
                            SheetName = sheetName,
                            RowIndex = newKvp.Value.RowIndex,
                            FieldName = field.Name,
                            RawValue = newValue?.ToString() ?? string.Empty,
                            Message = $"[Diff变更] {errorMsg}"
                        });
                    }
                }
            }

            return warnings;
        }

        protected bool TryConvertToObject(Type classType, DbDataReader row, List<string> headNames, string sheetName, int rowIndex, out object obj)
        {
            obj = Activator.CreateInstance(classType);
            var success = true;

            for (int columnID = 0; columnID < row.FieldCount && columnID < headNames.Count; columnID++)
            {
                var fieldInfo = GetPublicInstanceField(classType, headNames[columnID]);
                if (fieldInfo == null)
                {
                    LogError($"[{sheetName}] 第{rowIndex}行找不到字段 {headNames[columnID]}");
                    success = false;
                    continue;
                }

                var value = GetCellString(row, columnID);
                if (string.IsNullOrEmpty(value))
                    continue;

                m_typeConvert.m_extraInfo = $"{sheetName} 第{rowIndex}行 字段 {fieldInfo.Name}";
                if (!m_typeConvert.TryParse(fieldInfo.FieldType, value, -1, out var parsedValue))
                {
                    LogError($"[{sheetName}] 第{rowIndex}行字段 {fieldInfo.Name} 解析失败，原始值: {value}");
                    success = false;
                    continue;
                }

                if (parsedValue != null || !fieldInfo.FieldType.IsValueType || Nullable.GetUnderlyingType(fieldInfo.FieldType) != null)
                    fieldInfo.SetValue(obj, parsedValue);
            }

            m_typeConvert.m_extraInfo = string.Empty;
            return success;
        }

        private static byte[] SerializeToBytes(object obj)
        {
            return MemoryPackSerializer.Serialize(obj.GetType(), obj);
        }

        protected void AppendCsvRow(StringBuilder builder, DbDataReader row)
        {
            for (int i = 0; i < row.FieldCount; i++)
            {
                if (i > 0)
                    builder.Append(',');

                builder.Append(EscapeCsvCell(GetCellString(row, i)));
            }

            builder.AppendLine();
        }

        private static string EscapeCsvCell(string value)
        {
            value ??= string.Empty;
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0)
                return value;

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static FieldInfo GetPublicInstanceField(Type type, string fieldName)
        {
            return type.GetField(fieldName, BindingFlags.Public | BindingFlags.Instance);
        }

        private static string GetCellString(DbDataReader row, int index)
        {
            if (row == null || index < 0 || index >= row.FieldCount || row.IsDBNull(index))
                return null;

            return row.GetString(index);
        }

        private static string GetPrimaryKeyValue(FieldInfo field, object obj)
        {
            var value = field?.GetValue(obj);
            if (value == null)
                return null;

            if (value is IFormattable formattable)
                return formattable.ToString(null, CultureInfo.InvariantCulture);

            return value.ToString();
        }

        /// <summary>
        /// 公式单元格信息
        /// </summary>
        public struct FormulaCellInfo
        {
            /// <summary>单元格引用（如 "B3"）</summary>
            public string CellRef;
            /// <summary>公式文本（如 "SUM(A1:A10)"）</summary>
            public string FormulaText;
        }

        /// <summary>
        /// 检查标记了 [NoFormula] 的字段对应的 Excel 列是否存在公式
        /// </summary>
        private void CheckNoFormulaFields(
            string xlsxPath,
            string sheetName,
            Type classType,
            List<int> heads,
            List<string> headNames,
            List<ValidationError> allErrors)
        {
            var noFormulaColumns = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int i = 0; i < headNames.Count; i++)
            {
                var field = GetPublicInstanceField(classType, headNames[i]);
                if (field != null && field.GetCustomAttribute<NoFormulaAttribute>() != null)
                {
                    var colLetter = ColumnIndexToLetter(heads[i]);
                    noFormulaColumns[colLetter] = headNames[i];
                }
            }

            if (noFormulaColumns.Count == 0)
                return;

            var formulaCells = ScanFormulaCells(xlsxPath, sheetName);
            foreach (var cell in formulaCells)
            {
                var colLetter = ExtractColumnLetter(cell.CellRef);
                if (!noFormulaColumns.TryGetValue(colLetter, out var fieldName))
                    continue;

                allErrors.Add(new ValidationError
                {
                    SheetName = sheetName,
                    RowIndex = ExtractRowNumber(cell.CellRef),
                    FieldName = fieldName,
                    RawValue = cell.FormulaText,
                    Message = $"字段 '{fieldName}' 标记了 [NoFormula]，但单元格 {cell.CellRef} 的值由公式计算: {cell.FormulaText}"
                });
            }
        }

        /// <summary>
        /// Excel 列号（0-based）转列字母（A, B, ..., Z, AA, AB, ...）
        /// </summary>
        private static string ColumnIndexToLetter(int columnIndex)
        {
            var sb = new StringBuilder();
            int n = columnIndex;
            do
            {
                sb.Insert(0, (char)('A' + n % 26));
                n = n / 26 - 1;
            } while (n >= 0);
            return sb.ToString();
        }

        /// <summary>
        /// 从单元格引用中提取列字母部分（如 "B3" -> "B", "AA12" -> "AA"）
        /// </summary>
        private static string ExtractColumnLetter(string cellRef)
        {
            int i = 0;
            while (i < cellRef.Length && cellRef[i] >= 'A' && cellRef[i] <= 'Z')
                i++;
            return cellRef.Substring(0, i);
        }

        /// <summary>
        /// 从单元格引用中提取行号（如 "B3" -> 3, "AA12" -> 12）
        /// </summary>
        private static int ExtractRowNumber(string cellRef)
        {
            int i = 0;
            while (i < cellRef.Length && cellRef[i] >= 'A' && cellRef[i] <= 'Z')
                i++;
            if (int.TryParse(cellRef.Substring(i), out var row))
                return row;
            return 0;
        }

        private static readonly XNamespace s_nsSpreadsheet =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace s_nsRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace s_nsPackageRel =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>
        /// 预扫描 xlsx 文件中指定 Sheet 的公式单元格
        /// 直接读取底层 XML，检测 &lt;f&gt; 元素
        /// </summary>
        private List<FormulaCellInfo> ScanFormulaCells(string xlsxPath, string sheetName)
        {
            var result = new List<FormulaCellInfo>();

            try
            {
                using var zipStream = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var zip = new ZipArchive(zipStream, ZipArchiveMode.Read);

                var wbEntry = zip.GetEntry("xl/workbook.xml");
                if (wbEntry == null)
                    return result;

                XDocument wbDoc;
                using (var wbStream = wbEntry.Open())
                    wbDoc = XDocument.Load(wbStream);

                string rId = null;
                foreach (var sheet in wbDoc.Descendants(s_nsSpreadsheet + "sheet"))
                {
                    if (sheet.Attribute("name")?.Value == sheetName)
                    {
                        rId = sheet.Attribute(s_nsRelationship + "id")?.Value;
                        break;
                    }
                }

                if (rId == null)
                    return result;

                var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry == null)
                    return result;

                XDocument relsDoc;
                using (var relsStream = relsEntry.Open())
                    relsDoc = XDocument.Load(relsStream);

                string sheetPath = null;
                foreach (var rel in relsDoc.Descendants(s_nsPackageRel + "Relationship"))
                {
                    if (rel.Attribute("Id")?.Value == rId)
                    {
                        sheetPath = "xl/" + rel.Attribute("Target")?.Value;
                        break;
                    }
                }

                if (sheetPath == null)
                    return result;

                var sheetEntry = zip.GetEntry(sheetPath);
                if (sheetEntry == null)
                    return result;

                XDocument sheetDoc;
                using (var sheetStream = sheetEntry.Open())
                    sheetDoc = XDocument.Load(sheetStream);

                var sharedFormulas = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (var cell in sheetDoc.Descendants(s_nsSpreadsheet + "c"))
                {
                    var formula = cell.Element(s_nsSpreadsheet + "f");
                    if (formula == null)
                        continue;

                    var si = formula.Attribute("si")?.Value;
                    var formulaText = formula.Value;
                    if (si != null && !string.IsNullOrEmpty(formulaText))
                        sharedFormulas[si] = formulaText;
                }

                foreach (var cell in sheetDoc.Descendants(s_nsSpreadsheet + "c"))
                {
                    var formula = cell.Element(s_nsSpreadsheet + "f");
                    if (formula == null)
                        continue;

                    var formulaText = formula.Value;
                    if (string.IsNullOrEmpty(formulaText))
                    {
                        var si = formula.Attribute("si")?.Value;
                        if (si != null)
                            sharedFormulas.TryGetValue(si, out formulaText);

                        formulaText ??= "(共享公式)";
                    }

                    result.Add(new FormulaCellInfo
                    {
                        CellRef = cell.Attribute("r")?.Value ?? "??",
                        FormulaText = formulaText
                    });
                }
            }
            catch (Exception e)
            {
                Log($"[公式扫描] 读取 {sheetName} 的公式信息时出错: {e.Message}");
            }

            return result;
        }
    }

    /// <summary>
    /// StaticDataModel.cs 代码生成器
    /// 使用 StringBuilder 结构化生成代码，替代旧的字符串模板 + 魔法注释方式
    /// </summary>
    public class DataModelGenerate
    {
        protected virtual string staticdatamodel_path => Application.dataPath + "/StaticData/";
        protected virtual string bin_path => "StaticData/";
        protected virtual string config_path => "./excelconfig";

        protected Action<string, float, string, string> _EVENT_PROGRESS;
        protected Action _EVENT_END_PROGRESS;
        protected Action<string> _EVENT_LOG;
        protected Action<string> _EVENT_ERROR_LOG;

        private void Progress(float progress, string action, string name) => _EVENT_PROGRESS?.Invoke("DataModel 自动生成", progress, action, name);
        private void EndProgress() => _EVENT_END_PROGRESS?.Invoke();
        private void LogError(string error) => _EVENT_ERROR_LOG?.Invoke(error);
        private void Log(string log) => _EVENT_LOG?.Invoke(log);

        public void Init(Action EVENT_END_PROGRESS,
            Action<string, float, string, string> EVENT_PROGRESS,
            Action<string> EVENT_ERROR_LOG,
            Action<string> EVENT_LOG)
        {
            _EVENT_PROGRESS = EVENT_PROGRESS;
            _EVENT_END_PROGRESS = EVENT_END_PROGRESS;
            _EVENT_LOG = EVENT_LOG;
            _EVENT_ERROR_LOG = EVENT_ERROR_LOG;
        }

        public void GenerateStaticDataModel()
        {
            try
            {
                if (!File.Exists(config_path))
                {
                    LogError("Excel config is not found! Please load excels first!");
                    return;
                }

                var bin = File.ReadAllBytes(config_path);
                ExcelConfig config;
                try
                {
                    config = MemoryPackSerializer.Deserialize<ExcelConfig>(bin);
                }
                catch
                {
                    LogError("配置文件格式不兼容，请先重新执行 Serialize Excels");
                    return;
                }

                if (config?.Files == null)
                {
                    LogError("配置文件内容为空，请先重新执行 Serialize Excels");
                    return;
                }

                Progress(0, "GenerateStaticDataModel", "Start");

                var allSheets = new List<ExcelSheetInfo>();
                var duplicates = new HashSet<string>(StringComparer.Ordinal);
                var fileIndex = 0;
                foreach (var item in config.Files)
                {
                    Progress((float)fileIndex / Math.Max(1, config.Files.Count), "ReadExcel", item.Key);
                    fileIndex++;

                    if (item.Value?.Sheets == null)
                        continue;

                    foreach (var sheet in item.Value.Sheets)
                    {
                        if (sheet == null || string.IsNullOrWhiteSpace(sheet.SheetName) || string.IsNullOrWhiteSpace(sheet.TypeName))
                            continue;

                        if (!duplicates.Add(sheet.SheetName))
                        {
                            LogError($"StaticDataModel 生成失败，发现重复的 Sheet 名称: {sheet.SheetName}");
                            return;
                        }

                        allSheets.Add(new ExcelSheetInfo
                        {
                            SheetName = sheet.SheetName,
                            TypeName = sheet.TypeName
                        });
                    }
                }

                allSheets.Sort((left, right) => string.CompareOrdinal(left.SheetName, right.SheetName));
                var text = GenerateCode(allSheets);

                if (!Directory.Exists(staticdatamodel_path))
                    Directory.CreateDirectory(staticdatamodel_path);

                File.WriteAllText(Path.Combine(staticdatamodel_path, "StaticDataModel.cs"), text);
                Log($"StaticDataModel.cs 已生成，包含 {allSheets.Count} 个数据表");
                Progress(1, "GenerateStaticDataModel", "End");
            }
            catch (Exception e)
            {
                EndProgress();
                Debug.LogException(e);
            }
        }

        /// <summary>
        /// 使用 StringBuilder 结构化生成 StaticDataModel.cs 代码
        /// </summary>
        private string GenerateCode(List<ExcelSheetInfo> sheets)
        {
            var sb = new StringBuilder();

            sb.AppendLine("using System.Collections.Generic;");
            sb.AppendLine("using UnityEngine;");
            sb.AppendLine("using MemoryPack;");
            sb.AppendLine();
            sb.AppendLine("public class StaticDataModel");
            sb.AppendLine("{");

            foreach (var sheet in sheets)
                sb.AppendLine($"    public List<{GetCodeTypeName(sheet.TypeName)}> {GetCollectionPropertyName(sheet.SheetName)} {{ get; private set; }}");

            sb.AppendLine();
            sb.AppendLine("    public void Init()");
            sb.AppendLine("    {");
            foreach (var sheet in sheets)
                sb.AppendLine($"        {GetCollectionPropertyName(sheet.SheetName)} = MemoryPackDeserialize<List<{GetCodeTypeName(sheet.TypeName)}>>(\"{sheet.SheetName}\");");
            sb.AppendLine("    }");

            sb.AppendLine();
            sb.AppendLine("    private T MemoryPackDeserialize<T>(string filename)");
            sb.AppendLine("    {");
            sb.AppendLine($"        var asset = Resources.Load<TextAsset>(\"{bin_path}\" + filename);");
            sb.AppendLine("        if (asset == null)");
            sb.AppendLine("        {");
            sb.AppendLine("            Debug.LogError($\"StaticData asset not found: {filename}\");");
            sb.AppendLine("            return default;");
            sb.AppendLine("        }");
            sb.AppendLine();
            sb.AppendLine("        return MemoryPackSerializer.Deserialize<T>(asset.bytes);");
            sb.AppendLine("    }");
            sb.AppendLine("}");

            return sb.ToString();
        }

        private static string GetCodeTypeName(string typeName)
        {
            return "global::" + typeName.Replace('+', '.');
        }

        private static string GetCollectionPropertyName(string sheetName)
        {
            var identifier = SanitizeIdentifier(sheetName);
            if (identifier.EndsWith("s", StringComparison.Ordinal))
                return identifier + "List";

            return identifier + "s";
        }

        private static string SanitizeIdentifier(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "_";

            var sb = new StringBuilder();
            for (int i = 0; i < value.Length; i++)
            {
                var c = value[i];
                var isValid = i == 0
                    ? char.IsLetter(c) || c == '_'
                    : char.IsLetterOrDigit(c) || c == '_';

                sb.Append(isValid ? c : '_');
            }

            return sb.ToString();
        }
    }
}
