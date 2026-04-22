using System.Collections;
using System.Collections.Generic;
using System;
using MemoryPack;

[MemoryPackable]
public partial class sCityVO
{
    [Unique] public string m_id;
    public string m_name;
    public string m_desc;
    [Fixed] public string m_mapID;
    public string m_portID;
    public int[] m_initGovernmentLevel;
}
[MemoryPackable]
public partial class sCityLevelVO
{
    [Unique] public int m_level;
    public int m_exp;
    public int m_houseLimit;
}
