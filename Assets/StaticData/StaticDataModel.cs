using System.Collections.Generic;
using UnityEngine;
using MemoryPack;

public class StaticDataModel
{
    public List<global::sCityVO> sCityVOs { get; private set; }
    public List<global::sCityLevelVO> sCityLevelVOs { get; private set; }

    public void Init()
    {
        sCityVOs = MemoryPackDeserialize<List<global::sCityVO>>("sCityVO");
        sCityLevelVOs = MemoryPackDeserialize<List<global::sCityLevelVO>>("sCityLevelVO");
    }

    private T MemoryPackDeserialize<T>(string filename)
    {
        var asset = Resources.Load<TextAsset>("StaticData/" + filename);
        if (asset == null)
        {
            Debug.LogError($"StaticData asset not found: {filename}");
            return default;
        }

        return MemoryPackSerializer.Deserialize<T>(asset.bytes);
    }
}
