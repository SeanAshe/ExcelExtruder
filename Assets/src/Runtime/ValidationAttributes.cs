using System;
using System.Collections.Generic;

/// <summary>
/// 校验 Attribute 的通用接口
/// 每个校验 Attribute 实现自己的校验逻辑
/// </summary>
public interface IValidationAttribute
{
    /// <summary>
    /// 执行校验
    /// </summary>
    /// <param name="value">解析后的对象值</param>
    /// <param name="rawValue">Excel 中的原始字符串值</param>
    /// <returns>校验通过返回 null，失败返回错误描述</returns>
    string Validate(object value, string rawValue);
}

/// <summary>
/// 标记字段为必填项，值不能为 null 或空字符串
/// 基于原始字符串判断，兼容值类型（int 等默认值为 0 的场景）
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class RequiredAttribute : Attribute, IValidationAttribute
{
    public string Validate(object value, string rawValue)
    {
        if (string.IsNullOrEmpty(rawValue))
            return "字段标记为 [Required]，但值为空";
        return null;
    }
}

/// <summary>
/// 标记数值字段的有效范围
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class RangeAttribute : Attribute, IValidationAttribute
{
    public double Min { get; }
    public double Max { get; }

    public RangeAttribute(double min, double max)
    {
        Min = min;
        Max = max;
    }

    public string Validate(object value, string rawValue)
    {
        if (value == null) return null;

        double numericValue;
        if (TryConvertToDouble(value, out numericValue))
        {
            if (numericValue < Min || numericValue > Max)
                return $"值 {numericValue} 超出范围 [{Min}, {Max}]";
        }
        return null;
    }

    private static bool TryConvertToDouble(object value, out double result)
    {
        if (value is int i) { result = i; return true; }
        if (value is float f) { result = f; return true; }
        if (value is double d) { result = d; return true; }
        if (value is long l) { result = l; return true; }
        if (value is short s) { result = s; return true; }
        if (value is byte b) { result = b; return true; }
        if (value is uint ui) { result = ui; return true; }
        if (value is ulong ul) { result = ul; return true; }
        result = 0;
        return false;
    }
}

/// <summary>
/// 标记字段不允许由 Excel 公式计算得出
/// 序列化时会扫描 xlsx 底层 XML，检测对应列是否存在公式
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class NoFormulaAttribute : Attribute { }

/// <summary>
/// 显式声明某个类型对应的 Excel Sheet、输出名和主键字段。
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public class ExcelSheetAttribute : Attribute
{
    public string SheetName { get; }
    public string KeyField { get; }
    public string OutputName { get; set; }

    public ExcelSheetAttribute(string sheetName, string keyField, string outputName)
    {
        SheetName = sheetName;
        KeyField = keyField;
        OutputName = outputName;
    }
}

/// <summary>
/// Diff 校验接口，仅在数据发生改变时触发
/// </summary>
public interface IDiffValidationAttribute
{
    /// <summary>
    /// 执行差异校验
    /// </summary>
    /// <param name="oldValue">旧的解析值</param>
    /// <param name="newValue">新的解析值</param>
    /// <returns>如果校验失败返回错误信息，通过返回 null</returns>
    string ValidateDiff(object oldValue, object newValue);
}

/// <summary>
/// 标记字段建议不可被修改。只要旧数据存在该字段，新数据如果发生变动则会提示警告。
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class FixedAttribute : Attribute, IDiffValidationAttribute
{
    public string ValidateDiff(object oldValue, object newValue)
    {
        if (!object.Equals(oldValue, newValue))
        {
            return $"该字段标记了 [Fixed]，请谨慎修改。旧值: {oldValue}, 新值: {newValue}";
        }
        return null;
    }
}

/// <summary>
/// 标记数值字段的变化幅度不能超过指定的百分比。
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class MaxChangePercentAttribute : Attribute, IDiffValidationAttribute
{
    public float MaxPercent { get; }

    /// <summary>
    /// </summary>
    /// <param name="maxPercent">最大允许变化比例，例如 0.5 表示 50%</param>
    public MaxChangePercentAttribute(float maxPercent)
    {
        MaxPercent = maxPercent;
    }

    public string ValidateDiff(object oldValue, object newValue)
    {
        try
        {
            double oldNum = Convert.ToDouble(oldValue);
            double newNum = Convert.ToDouble(newValue);

            if (Math.Abs(oldNum) < 1e-6) return null; // 旧值为0时，跳过百分比校验，或者可以根据业务需求报错

            var change = Math.Abs((newNum - oldNum) / oldNum);
            if (change > MaxPercent)
            {
                return $"该字段标记了变化不能超过 {MaxPercent:P0}，旧值({oldNum}) -> 新值({newNum}) 变化了 {change:P2}";
            }
        }
        catch { }

        return null;
    }
}

/// <summary>
/// 上下文校验接口，用于跨行校验（如查重）
/// </summary>
public interface IContextValidationAttribute
{
    /// <summary>
    /// 执行上下文校验
    /// </summary>
    string ValidateContext(string fieldName, string value, Dictionary<string, HashSet<string>> cache);
}

/// <summary>
/// 标记字段在整列中必须唯一，不能有重复值。
/// 用于非主键字段的唯一约束。
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class UniqueAttribute : Attribute, IContextValidationAttribute
{
    public string ValidateContext(string fieldName, string value, Dictionary<string, HashSet<string>> cache)
    {
        return UniqueValidationUtility.ValidateUniqueValue(fieldName, value, cache, "[Unique]");
    }
}

/// <summary>
/// 显式标记导表主键字段。
/// 主键默认要求整列唯一，同时用于增量 Diff 和旧数据基线对比。
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class PrimaryKeyAttribute : Attribute, IContextValidationAttribute
{
    public string ValidateContext(string fieldName, string value, Dictionary<string, HashSet<string>> cache)
    {
        return UniqueValidationUtility.ValidateUniqueValue(fieldName, value, cache, "[PrimaryKey]");
    }
}

internal static class UniqueValidationUtility
{
    public static string ValidateUniqueValue(string fieldName, string value, Dictionary<string, HashSet<string>> cache, string attributeName)
    {
        if (string.IsNullOrEmpty(value)) return null;

        if (!cache.TryGetValue(fieldName, out var set))
        {
            set = new HashSet<string>();
            cache[fieldName] = set;
        }

        if (!set.Add(value))
        {
            return $"该字段标记了 {attributeName}，发现重复的值: {value}";
        }

        return null;
    }
}

/// <summary>
/// 新增行校验接口，仅在数据行被判定为“新增”时触发
/// </summary>
public interface INewRowValidationAttribute
{
    /// <summary>
    /// 执行新增行专属校验
    /// </summary>
    string ValidateNewRow(string fieldName, object newValue, Dictionary<string, object> oldDataMap);
}
