using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelExtruder
{
    /// <summary>
    /// 单条校验错误信息
    /// </summary>
    public class ValidationError
    {
        /// <summary>Sheet 名称（即类型名）</summary>
        public string SheetName { get; set; }
        /// <summary>数据行号（从1开始，不含表头行）</summary>
        public int RowIndex { get; set; }
        /// <summary>字段名</summary>
        public string FieldName { get; set; }
        /// <summary>原始值</summary>
        public string RawValue { get; set; }
        /// <summary>错误描述</summary>
        public string Message { get; set; }

        public override string ToString()
        {
            return $"[校验失败] {SheetName} 第{RowIndex}行 字段'{FieldName}' 值='{RawValue}': {Message}";
        }
    }

    /// <summary>
    /// 数据校验器，在序列化前对每行数据执行 Attribute 声明的校验规则
    /// 通过 IValidationAttribute 接口统一调用，无需硬编码各类校验逻辑
    /// </summary>
    public class DataValidator
    {
        /// <summary>
        /// 字段校验信息缓存，避免每行都重复反射
        /// key: Type, value: 该类型中所有需要校验的字段信息
        /// </summary>
        private readonly Dictionary<Type, List<FieldValidationInfo>> _cache
            = new Dictionary<Type, List<FieldValidationInfo>>();

        /// <summary>
        /// 去重缓存，服务于 IContextValidationAttribute
        /// </summary>
        private readonly Dictionary<string, HashSet<string>> _uniqueCache = new Dictionary<string, HashSet<string>>();

        /// <summary>
        /// 清理缓存上下文（每次解析新表前调用）
        /// </summary>
        public void ClearContext()
        {
            _uniqueCache.Clear();
        }

        /// <summary>
        /// 校验单个对象的所有字段
        /// </summary>
        public List<ValidationError> Validate(
            object obj,
            Type classType,
            string sheetName,
            int rowIndex,
            Dictionary<string, string> rawValues)
        {
            var errors = new List<ValidationError>();
            if (obj == null) return errors;

            // ========== 通用检测：Excel 公式和错误值 ==========
            if (rawValues != null)
            {
                foreach (var kv in rawValues)
                {
                    var formulaError = CheckFormulaOrError(kv.Value);
                    if (formulaError != null)
                    {
                        errors.Add(new ValidationError
                        {
                            SheetName = sheetName,
                            RowIndex = rowIndex,
                            FieldName = kv.Key,
                            RawValue = kv.Value,
                            Message = formulaError
                        });
                    }
                }
            }

            // ========== Attribute 声明式校验 ==========
            var validations = GetFieldValidations(classType);

            foreach (var validation in validations)
            {
                var value = validation.Field.GetValue(obj);
                string rawValue = null;
                rawValues?.TryGetValue(validation.Field.Name, out rawValue);

                // 统一通过 IValidationAttribute 接口调用校验
                foreach (var attr in validation.Validators)
                {
                    var errorMsg = attr.Validate(value, rawValue);
                    if (errorMsg != null)
                    {
                        errors.Add(new ValidationError
                        {
                            SheetName = sheetName,
                            RowIndex = rowIndex,
                            FieldName = validation.Field.Name,
                            RawValue = rawValue ?? "(null)",
                            Message = errorMsg
                        });
                    }
                }

                // 通过 IContextValidationAttribute 接口调用全局去重校验
                foreach (var attr in validation.ContextValidators)
                {
                    var errorMsg = attr.ValidateContext(validation.Field.Name, rawValue, _uniqueCache);
                    if (errorMsg != null)
                    {
                        errors.Add(new ValidationError
                        {
                            SheetName = sheetName,
                            RowIndex = rowIndex,
                            FieldName = validation.Field.Name,
                            RawValue = rawValue ?? "(null)",
                            Message = errorMsg
                        });
                    }
                }
            }

            return errors;
        }

        /// <summary>
        /// 获取某类型的所有需要校验的字段信息（带缓存）
        /// </summary>
        private List<FieldValidationInfo> GetFieldValidations(Type type)
        {
            if (_cache.TryGetValue(type, out var cached))
                return cached;

            var result = new List<FieldValidationInfo>();
            var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance);

            foreach (var field in fields)
            {
                var validators = new List<IValidationAttribute>();
                var diffValidators = new List<IDiffValidationAttribute>();
                var contextValidators = new List<IContextValidationAttribute>();
                var newRowValidators = new List<INewRowValidationAttribute>();

                // 收集所有校验相关的 Attribute
                foreach (var attr in field.GetCustomAttributes(true))
                {
                    if (attr is IValidationAttribute validationAttr)
                        validators.Add(validationAttr);
                    if (attr is IDiffValidationAttribute diffAttr)
                        diffValidators.Add(diffAttr);
                    if (attr is IContextValidationAttribute contextAttr)
                        contextValidators.Add(contextAttr);
                    if (attr is INewRowValidationAttribute newRowAttr)
                        newRowValidators.Add(newRowAttr);
                }

                if (validators.Count > 0 || diffValidators.Count > 0 || contextValidators.Count > 0 || newRowValidators.Count > 0)
                {
                    result.Add(new FieldValidationInfo
                    {
                        Field = field,
                        Validators = validators,
                        DiffValidators = diffValidators,
                        ContextValidators = contextValidators,
                        NewRowValidators = newRowValidators
                    });
                }
            }

            _cache[type] = result;
            return result;
        }

        /// <summary>
        /// 对单个字段执行 Diff 校验
        /// </summary>
        public List<string> ValidateDiff(Type classType, string fieldName, object oldValue, object newValue)
        {
            var validations = GetFieldValidations(classType);
            var fieldValidation = validations.Find(v => v.Field.Name == fieldName);
            if (fieldValidation == null || fieldValidation.DiffValidators == null) return null;

            var errors = new List<string>();
            foreach (var attr in fieldValidation.DiffValidators)
            {
                var errorMsg = attr.ValidateDiff(oldValue, newValue);
                if (errorMsg != null)
                {
                    errors.Add(errorMsg);
                }
            }
            return errors.Count > 0 ? errors : null;
        }

        /// <summary>
        /// 对单行（判定为新增的行）执行专属校验
        /// </summary>
        public List<ValidationError> ValidateNewRow(
            Type classType,
            string sheetName,
            int rowIndex,
            object newObj,
            Dictionary<string, object> oldDataMap)
        {
            var errors = new List<ValidationError>();
            var validations = GetFieldValidations(classType);

            foreach (var validation in validations)
            {
                if (validation.NewRowValidators.Count == 0) continue;
                
                var newValue = validation.Field.GetValue(newObj);

                foreach (var attr in validation.NewRowValidators)
                {
                    var errorMsg = attr.ValidateNewRow(validation.Field.Name, newValue, oldDataMap);
                    if (errorMsg != null)
                    {
                        errors.Add(new ValidationError
                        {
                            SheetName = sheetName,
                            RowIndex = rowIndex,
                            FieldName = validation.Field.Name,
                            RawValue = newValue?.ToString() ?? "(null)",
                            Message = errorMsg
                        });
                    }
                }
            }
            return errors;
        }

        /// <summary>
        /// 检测单元格值是否包含 Excel 公式或公式错误
        /// </summary>
        /// <param name="rawValue">原始字符串值</param>
        /// <returns>错误描述，无问题返回 null</returns>
        private static string CheckFormulaOrError(string rawValue)
        {
            if (string.IsNullOrEmpty(rawValue))
                return null;

            // 检测未求值的公式（以 = 开头）
            if (rawValue.StartsWith("="))
                return $"单元格包含未求值的 Excel 公式: {rawValue}";

            // 检测 Excel 公式错误值
            if (s_excelErrors.Contains(rawValue))
                return $"单元格包含 Excel 公式错误: {rawValue}";

            return null;
        }

        /// <summary>
        /// Excel 标准公式错误值
        /// </summary>
        private static readonly HashSet<string> s_excelErrors = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "#REF!",    // 引用无效
            "#VALUE!",  // 值类型错误
            "#N/A",     // 值不可用
            "#DIV/0!",  // 除以零
            "#NAME?",   // 名称无法识别
            "#NULL!",   // 交集为空
            "#NUM!",    // 数值无效
        };

        /// <summary>
        /// 字段校验信息
        /// </summary>
        private class FieldValidationInfo
        {
            public FieldInfo Field;
            public List<IValidationAttribute> Validators;
            public List<IDiffValidationAttribute> DiffValidators;
            public List<IContextValidationAttribute> ContextValidators;
            public List<INewRowValidationAttribute> NewRowValidators;
        }
    }
}
