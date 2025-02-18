using System.Collections.Generic;
using UnityEngine;
using MemoryPack;

public class StaticDataModel
{
    public List<sCityVO> sCityVOs  { set; get; }
    public List<sCityLevelVO> sCityLevelVOs  { set; get; }
    // @Dont delete - for Gen property@

    public void Init()
    {
        sCityVOs = MemoryPackDeserialize<List<sCityVO>>("sCityVO");
        sCityLevelVOs = MemoryPackDeserialize<List<sCityLevelVO>>("sCityLevelVO");
        // @Dont delete - for Gen Init Func@
    }
    private T MemoryPackDeserialize<T>(string filename)
    {
        var bin = Resources.Load<TextAsset>("StaticData/" + filename).bytes;
        return MemoryPackSerializer.Deserialize<T>(bin);
    }
}
