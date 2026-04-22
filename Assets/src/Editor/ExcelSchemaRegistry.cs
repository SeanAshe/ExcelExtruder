using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelExtruder
{
    internal sealed class ExcelSheetSchema
    {
        public string SheetName { get; set; }
        public string OutputName { get; set; }
        public string KeyFieldName { get; set; }
        public FieldInfo KeyField { get; set; }
        public Type DataType { get; set; }

        public string TypeName => DataType.FullName ?? DataType.Name;
        public string KeyTypeName => KeyField.FieldType.FullName ?? KeyField.FieldType.Name;
    }

    internal sealed class ExcelSchemaRegistry
    {
        private readonly Dictionary<string, ExcelSheetSchema> _schemasBySheetName = new Dictionary<string, ExcelSheetSchema>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _sheetByOutputName = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Action<string> _logError;

        public bool HasErrors { get; private set; }

        private ExcelSchemaRegistry(Action<string> logError)
        {
            _logError = logError;
        }

        public static ExcelSchemaRegistry Create(Assembly preferredAssembly, Action<string> logError)
        {
            var registry = new ExcelSchemaRegistry(logError);
            registry.Load(preferredAssembly);
            return registry;
        }

        public bool TryGetSchema(string sheetName, out ExcelSheetSchema schema)
        {
            return _schemasBySheetName.TryGetValue(sheetName, out schema);
        }

        private void Load(Assembly preferredAssembly)
        {
            var seenTypes = new HashSet<Type>();

            if (preferredAssembly != null)
                LoadAssembly(preferredAssembly, seenTypes);

            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (assembly == preferredAssembly)
                    continue;

                LoadAssembly(assembly, seenTypes);
            }
        }

        private void LoadAssembly(Assembly assembly, HashSet<Type> seenTypes)
        {
            foreach (var type in GetLoadableTypes(assembly))
            {
                if (type == null || !seenTypes.Add(type))
                    continue;

                var attribute = type.GetCustomAttribute<ExcelSheetAttribute>(false);
                if (attribute == null)
                    continue;

                Register(type, attribute);
            }
        }

        private void Register(Type type, ExcelSheetAttribute attribute)
        {
            if (string.IsNullOrWhiteSpace(attribute.SheetName))
            {
                ReportError($"[ExcelSheet] {type.FullName} 缺少 SheetName");
                return;
            }

            if (string.IsNullOrWhiteSpace(attribute.KeyField))
            {
                ReportError($"[ExcelSheet] {type.FullName} 缺少 KeyField");
                return;
            }

            if (string.IsNullOrWhiteSpace(attribute.OutputName))
            {
                ReportError($"[ExcelSheet] {type.FullName} 缺少 OutputName");
                return;
            }

            var keyField = type.GetField(attribute.KeyField, BindingFlags.Public | BindingFlags.Instance);
            if (keyField == null)
            {
                ReportError($"[ExcelSheet] {type.FullName} 找不到 KeyField '{attribute.KeyField}'");
                return;
            }

            if (_schemasBySheetName.TryGetValue(attribute.SheetName, out var existingSchema))
            {
                ReportError($"[ExcelSheet] SheetName '{attribute.SheetName}' 重复注册: {existingSchema.TypeName} 和 {type.FullName}");
                return;
            }

            if (_sheetByOutputName.TryGetValue(attribute.OutputName, out var existingSheet))
            {
                ReportError($"[ExcelSheet] OutputName '{attribute.OutputName}' 重复注册: {existingSheet} 和 {attribute.SheetName}");
                return;
            }

            var schema = new ExcelSheetSchema
            {
                SheetName = attribute.SheetName,
                OutputName = attribute.OutputName,
                KeyFieldName = attribute.KeyField,
                KeyField = keyField,
                DataType = type
            };

            _schemasBySheetName.Add(schema.SheetName, schema);
            _sheetByOutputName.Add(schema.OutputName, schema.SheetName);
        }

        private void ReportError(string message)
        {
            HasErrors = true;
            _logError?.Invoke(message);
        }

        private static IEnumerable<Type> GetLoadableTypes(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                var result = new List<Type>();
                foreach (var type in ex.Types)
                {
                    if (type != null)
                        result.Add(type);
                }

                return result;
            }
        }
    }
}
