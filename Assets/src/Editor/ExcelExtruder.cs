using UnityEngine;
using Sylvan.Data;
using Sylvan.Data.Excel;
using Sylvan.Data.Csv;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Data.Common;
using System;
using System.Reflection;
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
        private Action<string, float, string, string> _EVENT_PROGRESS;
        private Action _EVENT_END_PROGRESS;
        private Action<string> _EVENT_LOG;
        private Action<string> _EVENT_ERROR_LOG;
        private void Progress(float progress, string action, string name) => _EVENT_PROGRESS?.Invoke("Excels 序列化", progress, action, name);
        private void EndProgress() => _EVENT_END_PROGRESS?.Invoke();
        private void LogError(string error) => _EVENT_LOG?.Invoke(error);
        private void Log(string log) => _EVENT_LOG?.Invoke(log);
        public void Init(TypeConvert typeConvert, Action EVENT_END_PROGRESS,
            Action<string, float, string, string> EVENT_PROGRESS,
            Action<string> EVENT_ERROR_LOG,
            Action<string> EVENT_LOG)
        {
            Assembly assembly = null;
            if (File.Exists("./Library/ScriptAssemblies/Assembly-CSharp.dll"))
                assembly = Assembly.LoadFile(System.Environment.CurrentDirectory + "\\Library\\ScriptAssemblies\\Assembly-CSharp.dll");

            m_typeConvert = typeConvert;
            m_typeConvert.Init(assembly);

            if (!Directory.Exists(EXCELRES_PATH)) Directory.CreateDirectory(EXCELRES_PATH);
            if (!Directory.Exists(CSV_PATH)) Directory.CreateDirectory(CSV_PATH);
            if (!Directory.Exists(BIN_PATH)) Directory.CreateDirectory(BIN_PATH);

            _EVENT_PROGRESS += EVENT_PROGRESS;
            _EVENT_END_PROGRESS += EVENT_END_PROGRESS;
            _EVENT_LOG += EVENT_LOG;
            _EVENT_ERROR_LOG += EVENT_ERROR_LOG;
        }
        private Dictionary<string, List<string>> ExcelInfos;
        public void SerializeAllExcel()
        {
            try
            {
                Progress(0, "SerializeExcels", "Start");
                DirectoryInfo folder = new DirectoryInfo(EXCELRES_PATH);
                var files = folder.GetFiles("*.xlsx", SearchOption.AllDirectories);
                var count = files.Length;
                var index = 0;
                ExcelInfos = new Dictionary<string, List<string>>();
                foreach (FileInfo file in files)
                {
                    if (file.Name.StartsWith("~$")) continue;
                    Progress(index / count, "SerializeExcel", file.Name);
                    var info = ConverOneExcel(file.Name);
                    ExcelInfos[info.Item1] = info.Item2;
                    index += 1;
                }
                SaveExcelInfos();
                Progress(1, "SerializeExcels", "End");
            }
            catch (Exception e)
            {
                EndProgress();
                Debug.LogException(e);
            }
        }

        private void SaveExcelInfos()
        {
            var bin = MemoryPackSerializer.Serialize(ExcelInfos);
            File.WriteAllBytes(CONFIG_PATH, bin);
        }

        public (string, List<string>) ConverOneExcel(string fileName)
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
                sheets.Add(sheetName);
                var heads = new List<int>();
                var headNames = new List<string>();
                Type classType = m_typeConvert.TryGetType(sheetName);
                object genericList = TypeConvert.CreateGeneric(typeof(List<>), classType);
                Type listType = genericList.GetType();
                MethodInfo mi = listType.GetMethod("Add");

                using var cdw = new StreamWriter(CSV_PATH + sheetName + ".csv", false);
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
                    }
                    var row = edr.Select(heads.ToArray());
                    Write2Csv(cdw, row);
                    if (edr.GetString(0) == EXCEL_FIELDNAME) continue;
                    var obj = Convert2Object(classType, row, headNames);
                    mi.Invoke(genericList, new object[] { obj });
                }
                MemoryPackSerializeAndSave(genericList, classType);

            } while(edr.NextResult());

            return (fileName, sheets);
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
            var Serializer = typeof(MemoryPack.MemoryPackSerializer);
            var methods = Serializer.GetMethods();
            MethodInfo methodInfo = null;
            foreach (var m in methods)
            {
                if (m.Name == "Serialize" && m.ReturnParameter.ParameterType == typeof(byte[]))
                {
                    methodInfo = m;
                    break;
                }
            }
            if (methodInfo == null)
            {
                Debug.Log(type.ToString());
                return;
            }
            var bit = methodInfo.Invoke(null, new object[] { obj.GetType(), obj, default});
            byte[] buff = bit as byte[];
            File.WriteAllBytes(BIN_PATH + type.ToString() + ".bytes", buff);
        }
    }
    public class DataModelGenerate
    {
        protected virtual string staticdatamodel_path => Application.dataPath + "/StaticData/";
        protected virtual string bin_path => "StaticData/";
        protected virtual string config_path => "./excelconfig";
        private const string STATICDATAMODEL_CONST=
@"using System.Collections.Generic;
using UnityEngine;
using MemoryPack;

public class StaticDataModel
{
    // @Dont delete - for Gen property@

    public void Init()
    {
        // @Dont delete - for Gen Init Func@
    }
    private T MemoryPackDeserialize<T>(string filename)
    {
        var bin = Resources.Load<TextAsset>(""@bin_path@"" + filename).bytes;
        return MemoryPackSerializer.Deserialize<T>(bin);
    }
}
";
        protected Action<string, float, string, string> _EVENT_PROGRESS;
        protected Action _EVENT_END_PROGRESS;
        protected Action<string> _EVENT_LOG;
        protected Action<string> _EVENT_ERROR_LOG;
        private void Progress(float progress, string action, string name) => _EVENT_PROGRESS?.Invoke("DataModel 自动生成",progress, action, name);
        private void EndProgress() => _EVENT_END_PROGRESS?.Invoke();
        private void LogError(string error) => _EVENT_LOG?.Invoke(error);
        private void Log(string log) => _EVENT_LOG?.Invoke(log);
        public void Init(Action EVENT_END_PROGRESS,
            Action<string, float, string, string> EVENT_PROGRESS,
            Action<string> EVENT_ERROR_LOG,
            Action<string> EVENT_LOG)
        {
            _EVENT_PROGRESS += EVENT_PROGRESS;
            _EVENT_END_PROGRESS += EVENT_END_PROGRESS;
            _EVENT_LOG += EVENT_LOG;
            _EVENT_ERROR_LOG += EVENT_ERROR_LOG;
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
                var excelInfos = MemoryPackSerializer.Deserialize<Dictionary<string, List<string>>>(bin);

                Progress(0, "GenerateStaticDataModel", "Start");
                var i = 0;
                var text = STATICDATAMODEL_CONST.Replace("@bin_path@", bin_path);
                foreach (var item in excelInfos)
                {
                    Progress(i / excelInfos.Count, "ReadExcel", $"{item.Key}");
                    var ii = 0;
                    foreach (var sheet in item.Value)
                    {
                        Progress(ii / item.Value.Count, "ReadSheet", $"{sheet}");
                        text = text.Insert(text.IndexOf("// @Dont delete - for Gen property@"), $"public List<{sheet}> {sheet}s  {{ set; get; }}\r\n    ");
                        text = text.Insert(text.IndexOf("// @Dont delete - for Gen Init Func@"), $"{sheet}s = MemoryPackDeserialize<List<{sheet}>>(\"{sheet}\");\r\n        ");
                    }
                    Progress(1, "ReadSheets", "End");
                }
                if (!Directory.Exists(staticdatamodel_path)) Directory.CreateDirectory(staticdatamodel_path);
                File.WriteAllText(staticdatamodel_path + "StaticDataModel.cs", text);
                Progress(1, "GenerateStaticDataModel", "End");
            }
            catch (Exception e)
            {
                EndProgress();
                Debug.LogException(e);
            }
        }
    }
}
