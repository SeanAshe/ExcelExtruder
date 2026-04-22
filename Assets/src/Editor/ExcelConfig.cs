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
        public Dictionary<string, ExcelFileInfo> Files { get; set; } = new Dictionary<string, ExcelFileInfo>();
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
        /// 该 Excel 文件包含的所有 Sheet 名称列表
        /// </summary>
        public List<string> Sheets { get; set; } = new List<string>();
    }
}
