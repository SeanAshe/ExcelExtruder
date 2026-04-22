using System.Collections;
using System.Collections.Generic;
using System;
using MemoryPack;

[ExcelSheet(nameof(sCityVO), nameof(m_id), nameof(sCityVO))]
[MemoryPackable]
public partial class sCityVO
{
    [PrimaryKey] public string m_id;
    public string m_name;
    public string m_desc;
    public string m_mapID;
    public string m_portID;
    public int[] m_initGovernmentLevel;
}
[ExcelSheet(nameof(sCityLevelVO), nameof(m_level), nameof(sCityLevelVO))]
[MemoryPackable]
public partial class sCityLevelVO
{
    [PrimaryKey] public int m_level;
    public int m_exp;
    public int m_houseLimit;
}
