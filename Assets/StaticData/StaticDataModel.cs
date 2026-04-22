using System.Collections.Generic;
using UnityEngine;
using MemoryPack;

public class StaticDataModel
{
    public List<global::sCityVO> sCityVOs { get; private set; }
    public Dictionary<string, global::sCityVO> sCityVOsByKey { get; private set; }
    public List<global::sCityLevelVO> sCityLevelVOs { get; private set; }
    public Dictionary<int, global::sCityLevelVO> sCityLevelVOsByKey { get; private set; }

    public void Init()
    {
        sCityVOs = MemoryPackDeserialize<List<global::sCityVO>>("sCityVO");
        sCityVOsByKey = BuildIndex(sCityVOs, item => item.m_id, "sCityVO", "m_id");
        sCityLevelVOs = MemoryPackDeserialize<List<global::sCityLevelVO>>("sCityLevelVO");
        sCityLevelVOsByKey = BuildIndex(sCityLevelVOs, item => item.m_level, "sCityLevelVO", "m_level");
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

    private Dictionary<TKey, TValue> BuildIndex<TKey, TValue>(List<TValue> items, System.Func<TValue, TKey> keySelector, string sheetName, string keyFieldName)
    {
        var dict = new Dictionary<TKey, TValue>();
        if (items == null)
            return dict;

        foreach (var item in items)
        {
            var key = keySelector(item);
            if (!dict.TryAdd(key, item))
                Debug.LogError($"StaticData key duplicated: {sheetName}.{keyFieldName} = {key}");
        }

        return dict;
    }
}
