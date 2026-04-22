using System.Collections.Generic;
using MemoryPack;

namespace ExcelExtruder
{
    /// <summary>
    /// Excel 配置根结构，存储所有已序列化 Excel 文件的元信息
    /// 用于增量序列化时判断文件是否变更
    /// </summary>
    [MemoryPackable]
    public partial class ExcelConfig
    {
        /// <summary>
        /// key: Excel 文件名, value: 该文件的元信息
        /// </summary>
        public Dictionary<string, ExcelFileInfo> Files { get; set; } = new Dictionary<string, ExcelFileInfo>(System.StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 单个 Excel 文件的元信息
    /// </summary>
    [MemoryPackable]
    public partial class ExcelFileInfo
    {
        /// <summary>
        /// 文件的 MD5 哈希值，用于判断文件是否变更
        /// </summary>
        public string Hash { get; set; }

        /// <summary>
        /// 该 Excel 文件包含的所有 Sheet 元信息
        /// </summary>
        public List<ExcelSheetInfo> Sheets { get; set; } = new List<ExcelSheetInfo>();
    }

    /// <summary>
    /// 单个 Sheet 的导出元信息
    /// </summary>
    [MemoryPackable]
    public partial class ExcelSheetInfo
    {
        /// <summary>
        /// Excel 中的 Sheet 名称，也是资源文件名
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 对应的数据类型全名，用于代码生成时输出正确的 C# 类型
        /// </summary>
        public string TypeName { get; set; }
    }
}
