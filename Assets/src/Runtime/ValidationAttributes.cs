using System;

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
