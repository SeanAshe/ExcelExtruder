using UnityEngine;
using Sylvan.Data;
using Sylvan.Data.Excel;
using Sylvan.Data.Csv;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Data.Common;
using System;
using System.Reflection;
using System.Text;
using System.Security.Cryptography;
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

        private ExcelConfig m_config;

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
                return config ?? new ExcelConfig();
            }
            catch
            {
                // 旧格式不兼容，返回空配置以触发全量重建
                Log("配置文件格式不兼容，将执行全量重建");
                return new ExcelConfig();
            }
        }

        /// <summary>
        /// 保存 Excel 配置
        /// </summary>
        private void SaveExcelConfig()
        {
            var bin = MemoryPackSerializer.Serialize(m_config);
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
                DirectoryInfo folder = new DirectoryInfo(EXCELRES_PATH);
                var files = folder.GetFiles("*.xlsx", SearchOption.AllDirectories);
                var count = files.Length;
                var index = 0;
                var skippedCount = 0;

                // 加载已有配置用于增量判断
                m_config = forceAll ? new ExcelConfig() : LoadExcelConfig();

                foreach (FileInfo file in files)
                {
                    if (file.Name.StartsWith("~$")) continue;
                    Progress((float)index / count, "SerializeExcel", file.Name);

                    // 增量判断：比对文件哈希
                    var currentHash = ComputeFileHash(file.FullName);
                    if (!forceAll
                        && m_config.Files.TryGetValue(file.Name, out var existingInfo)
                        && existingInfo.Hash == currentHash)
                    {
                        Log($"[增量跳过] {file.Name} 未变更");
                        skippedCount++;
                        index += 1;
                        continue;
                    }

                    // 需要序列化
                    var info = ConvertOneExcel(file.Name);
                    m_config.Files[info.FileName] = new ExcelFileInfo
                    {
                        Hash = currentHash,
                        Sheets = info.Sheets
                    };
                    index += 1;
                }

                SaveExcelConfig();
                Log($"序列化完成: 处理 {index - skippedCount} 个文件, 跳过 {skippedCount} 个未变更文件");
                Progress(1, "SerializeExcels", "End");
            }
            catch (Exception e)
            {
                EndProgress();
                Debug.LogException(e);
            }
        }

        /// <summary>
        /// 单个 Excel 文件的序列化结果
        /// </summary>
        public struct ExcelConvertResult
        {
            public string FileName;
            public List<string> Sheets;
        }

        public ExcelConvertResult ConvertOneExcel(string fileName)
        {
            ExcelWorkbookType workbooktype = ExcelDataReader.GetWorkbookType(EXCELRES_PATH + fileName);
            var sheets = new List<string>();

            using var fs = new FileInfo(EXCELRES_PATH + fileName).Open(FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var edr = ExcelDataReader.Create(fs, workbooktype, new ExcelDataReaderOptions()
            {
                Schema = ExcelSchema.NoHeaders,
                Culture = CultureInfo.InvariantCulture,
            });

            do
            {
                var sheetName = edr.WorksheetName;
                m_validator.ClearContext();
                sheets.Add(sheetName);
                var heads = new List<int>();
                var headNames = new List<string>();
                Type classType = m_typeConvert.TryGetType(sheetName);
                object genericList = TypeConvert.CreateGeneric(typeof(List<>), classType);
                var list = (IList)genericList;

                // 1. 准备字典，用于 Diff 校验
                Dictionary<string, object> oldDataMap = null;
                var newDataMap = new Dictionary<string, object>();

                using var cdw = new StreamWriter(CSV_PATH + sheetName + ".csv", false);
                int dataRowIndex = 0;

                var allErrors = new List<ValidationError>();
                bool formulaScanned = false;

                while(edr.Read())
                {
                    if (edr.GetString(0) == EXCEL_SKIP) continue;
                    if (edr.GetString(0) == EXCEL_FIELDNAME)
                    {
                        for (int i = 0; i < edr.FieldCount; i++)
                        {
                            if (edr.GetString(i) != null
                                && edr.GetString(i) != ""
                                && edr.GetString(i) != EXCEL_SKIP
                                && edr.GetString(i) != EXCEL_FIELDNAME)
                            {
                                heads.Add(i);
                                headNames.Add(edr.GetString(i));
                            }
                        }

                        // headNames 建立完成后，对 [NoFormula] 字段执行公式扫描
                        if (!formulaScanned)
                        {
                            formulaScanned = true;
                            CheckNoFormulaFields(EXCELRES_PATH + fileName, sheetName, classType, heads, headNames, allErrors);

                            // 此时已有正确的 headNames，可以根据第一列提取主键并反序列化旧数据
                            oldDataMap = headNames.Count > 0 ? LoadBaselineFromBin(classType, headNames[0]) : new Dictionary<string, object>();
                        }
                    }
                    var row = edr.Select(heads.ToArray());
                    Write2Csv(cdw, row);
                    if (edr.GetString(0) == EXCEL_FIELDNAME) continue;

                    dataRowIndex++;

                    // 收集原始值用于校验报告
                    var rawValues = new Dictionary<string, string>();
                    for (int i = 0; i < row.FieldCount && i < headNames.Count; i++)
                    {
                        rawValues[headNames[i]] = row.GetString(i);
                    }

                    var obj = Convert2Object(classType, row, headNames);
                    if (obj == null) continue;

                    // 数据校验（收集错误，不阻断序列化）
                    var errors = m_validator.Validate(obj, classType, sheetName, dataRowIndex, rawValues);
                    allErrors.AddRange(errors);

                    string pk = row.GetString(0);
                    if (!string.IsNullOrEmpty(pk))
                    {
                        newDataMap[pk] = obj;

                        // 判定是否是新增行，如果是，则执行新增行的全局验证（如 [NoReuse]）
                        if (oldDataMap != null && !oldDataMap.ContainsKey(pk))
                        {
                            var newRowErrors = m_validator.ValidateNewRow(classType, sheetName, dataRowIndex, obj, oldDataMap);
                            allErrors.AddRange(newRowErrors);
                        }
                    }

                    list.Add(obj);
                }

                // 2. 循环结束后，执行后置的 Diff 差异校验
                PerformDiffValidation(sheetName, classType, oldDataMap, newDataMap, allErrors);

                // 统一输出所有校验错误
                if (allErrors.Count > 0)
                {
                    LogError($"[校验报告] {sheetName} 共发现 {allErrors.Count} 个校验错误:");
                    foreach (var error in allErrors)
                    {
                        LogError(error.ToString());
                    }
                }

                MemoryPackSerializeAndSave(genericList, classType);

            } while(edr.NextResult());

            return new ExcelConvertResult { FileName = fileName, Sheets = sheets };
        }

        /// <summary>
        /// 从 StaticData 目录加载对应类型的 .bytes 文件，反序列化作为基线数据
        /// key: 主键 (导出后的第一列的值), value: 反序列化后的旧对象
        /// </summary>
        private Dictionary<string, object> LoadBaselineFromBin(Type classType, string pkFieldName)
        {
            var dataMap = new Dictionary<string, object>();
            string binPath = BIN_PATH + classType.ToString() + ".bytes";
            if (!File.Exists(binPath)) return dataMap;

            try
            {
                var bin = File.ReadAllBytes(binPath);
                var listType = typeof(List<>).MakeGenericType(classType);
                var oldList = MemoryPackSerializer.Deserialize(listType, bin) as IList;
                
                if (oldList != null)
                {
                    var pkField = classType.GetField(pkFieldName);
                    if (pkField == null) return dataMap;

                    foreach (var item in oldList)
                    {
                        if (item == null) continue;
                        var pkValue = pkField.GetValue(item)?.ToString();
                        if (!string.IsNullOrEmpty(pkValue))
                        {
                            dataMap[pkValue] = item;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Log($"加载基线数据 {classType.Name}.bytes 失败: {e.Message}");
            }
            return dataMap;
        }

        /// <summary>
        /// 执行新老数据的 Diff 差异校验（支持插行、插列）
        /// </summary>
        private void PerformDiffValidation(
            string sheetName, Type classType,
            Dictionary<string, object> oldDataMap,
            Dictionary<string, object> newDataMap,
            List<ValidationError> allErrors)
        {
            if (oldDataMap == null || oldDataMap.Count == 0 || newDataMap.Count == 0) return;

            int rowIndex = 0; 
            var fields = classType.GetFields(BindingFlags.Public | BindingFlags.Instance);

            foreach (var newKvp in newDataMap)
            {
                rowIndex++;
                string pk = newKvp.Key;
                object newObj = newKvp.Value;

                // 插行：旧数据中没有，视为新增行，安全跳过 Diff
                if (!oldDataMap.TryGetValue(pk, out var oldObj)) continue;

                // 发生更新：对比两行数据
                foreach (var field in fields)
                {
                    object oldValue = field.GetValue(oldObj);
                    object newValue = field.GetValue(newObj);

                    var diffErrors = m_validator.ValidateDiff(classType, field.Name, oldValue, newValue);
                    if (diffErrors != null)
                    {
                        foreach (var errorMsg in diffErrors)
                        {
                            allErrors.Add(new ValidationError
                            {
                                SheetName = sheetName,
                                RowIndex = rowIndex,
                                FieldName = field.Name,
                                RawValue = newValue?.ToString() ?? "",
                                Message = $"[Diff变更] {errorMsg}"
                            });
                        }
                    }
                }
            }
        }

        protected void Write2Csv(StreamWriter ws, DbDataReader row)
        {
            var data = new List<string>();
            for (var i = 0; i < row.FieldCount; i++)
            {
                data.Add(row.GetString(i));
            }
            ws.WriteLine(string.Join(",", data.ToArray()));
        }

        protected object Convert2Object(Type classType, DbDataReader row, List<string> headNames)
        {
            object obj = Activator.CreateInstance(classType);
            for (int columnID = 0; columnID < row.FieldCount; columnID ++)
            {
                var fieldInfo = classType.GetField(headNames[columnID]);
                if (fieldInfo == null)
                {
                    LogError("Can't find the field \"" + headNames[columnID] + "\" in the type \"" + classType.ToString());
                    return null;
                }
                string value = row.GetString(columnID);
                object o;
                if (m_typeConvert.TryParse(fieldInfo.FieldType, value, -1 , out o))
                {
                    fieldInfo.SetValue(obj, o);
                }
            }
            return obj;
        }

        protected void MemoryPackSerializeAndSave(object obj, Type type)
        {
            // 直接使用 MemoryPack 的非泛型 API，无需反射遍历方法
            var bin = MemoryPackSerializer.Serialize(obj.GetType(), obj);
            File.WriteAllBytes(BIN_PATH + type.ToString() + ".bytes", bin);
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
            string xlsxPath, string sheetName, Type classType,
            List<int> heads, List<string> headNames,
            List<ValidationError> allErrors)
        {
            // 找出标记 [NoFormula] 的字段名及其对应的 Excel 列号
            var noFormulaColumns = new Dictionary<string, string>(); // key: 列字母前缀, value: 字段名
            for (int i = 0; i < headNames.Count; i++)
            {
                var field = classType.GetField(headNames[i]);
                if (field != null && field.GetCustomAttribute<NoFormulaAttribute>() != null)
                {
                    var colLetter = ColumnIndexToLetter(heads[i]);
                    noFormulaColumns[colLetter] = headNames[i];
                }
            }

            if (noFormulaColumns.Count == 0) return;

            // 扫描公式单元格
            var formulaCells = ScanFormulaCells(xlsxPath, sheetName);

            foreach (var cell in formulaCells)
            {
                // 从单元格引用提取列字母（如 "B3" -> "B"）
                var colLetter = ExtractColumnLetter(cell.CellRef);
                if (noFormulaColumns.TryGetValue(colLetter, out var fieldName))
                {
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

        // xlsx OpenXML 命名空间
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

                // 1. 读取 workbook.xml，找到目标 Sheet 的 rId
                var wbEntry = zip.GetEntry("xl/workbook.xml");
                if (wbEntry == null) return result;

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
                if (rId == null) return result;

                // 2. 读取 workbook.xml.rels，找到 Sheet 文件路径
                var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry == null) return result;

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
                if (sheetPath == null) return result;

                // 3. 读取 Sheet XML，查找所有含 <f> 的单元格
                var sheetEntry = zip.GetEntry(sheetPath);
                if (sheetEntry == null) return result;

                XDocument sheetDoc;
                using (var sheetStream = sheetEntry.Open())
                    sheetDoc = XDocument.Load(sheetStream);

                // 先收集共享公式的原始文本（si 索引 -> 公式文本）
                var sharedFormulas = new Dictionary<string, string>();
                foreach (var cell in sheetDoc.Descendants(s_nsSpreadsheet + "c"))
                {
                    var formula = cell.Element(s_nsSpreadsheet + "f");
                    if (formula == null) continue;

                    // 共享公式的定义方（有 ref 属性的那个）存储了完整公式
                    var si = formula.Attribute("si")?.Value;
                    var formulaText = formula.Value;
                    if (si != null && !string.IsNullOrEmpty(formulaText))
                        sharedFormulas[si] = formulaText;
                }

                // 收集所有公式单元格，共享公式引用方回溯原始文本
                foreach (var cell in sheetDoc.Descendants(s_nsSpreadsheet + "c"))
                {
                    var formula = cell.Element(s_nsSpreadsheet + "f");
                    if (formula == null) continue;

                    var formulaText = formula.Value;

                    // 共享公式引用方的 <f> 为空，通过 si 索引查找原始公式
                    if (string.IsNullOrEmpty(formulaText))
                    {
                        var si = formula.Attribute("si")?.Value;
                        if (si != null)
                            sharedFormulas.TryGetValue(si, out formulaText);
                        formulaText = formulaText ?? "(共享公式)";
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
        private void Progress(float progress, string action, string name) => _EVENT_PROGRESS?.Invoke("DataModel 自动生成",progress, action, name);
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

                Progress(0, "GenerateStaticDataModel", "Start");

                // 收集所有 Sheet 名称
                var allSheets = new List<string>();
                int fileIndex = 0;
                foreach (var item in config.Files)
                {
                    Progress((float)fileIndex / config.Files.Count, "ReadExcel", item.Key);
                    allSheets.AddRange(item.Value.Sheets);
                    fileIndex++;
                }

                // 使用 StringBuilder 结构化生成代码
                var text = GenerateCode(allSheets);

                if (!Directory.Exists(staticdatamodel_path))
                    Directory.CreateDirectory(staticdatamodel_path);

                File.WriteAllText(staticdatamodel_path + "StaticDataModel.cs", text);
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
        private string GenerateCode(List<string> sheets)
        {
            var sb = new StringBuilder();

            // 1. using 声明
            sb.AppendLine("using System.Collections.Generic;");
            sb.AppendLine("using UnityEngine;");
            sb.AppendLine("using MemoryPack;");
            sb.AppendLine();

            // 2. 类定义开始
            sb.AppendLine("public class StaticDataModel");
            sb.AppendLine("{");

            // 3. 属性声明
            foreach (var sheet in sheets)
            {
                sb.AppendLine($"    public List<{sheet}> {sheet}s {{ get; set; }}");
            }
            sb.AppendLine();

            // 4. Init 方法
            sb.AppendLine("    public void Init()");
            sb.AppendLine("    {");
            foreach (var sheet in sheets)
            {
                sb.AppendLine($"        {sheet}s = MemoryPackDeserialize<List<{sheet}>>(\"{sheet}\");");
            }
            sb.AppendLine("    }");

            // 5. 反序列化辅助方法
            sb.AppendLine();
            sb.AppendLine("    private T MemoryPackDeserialize<T>(string filename)");
            sb.AppendLine("    {");
            sb.AppendLine($"        var bin = Resources.Load<TextAsset>(\"{bin_path}\" + filename).bytes;");
            sb.AppendLine("        return MemoryPackSerializer.Deserialize<T>(bin);");
            sb.AppendLine("    }");

            // 6. 类定义结束
            sb.AppendLine("}");

            return sb.ToString();
        }
    }
}
