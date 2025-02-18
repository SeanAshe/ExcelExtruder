using UnityEditor;
using UnityEngine;
using ExcelExtruder;

public class EditorUI : EditorWindow
{
    [MenuItem("Tools/Excel/Serialize Excels", false, 0)]
    static void OpenLoadingScene()
    {
        var typeConvert = new TypeConvert(EVENT_ERROR_LOG);
        var excelSerialize = new ExcelSerialize();
        excelSerialize.Init(typeConvert, EVENT_END_PROGRESS, EVENT_PROGRESS, EVENT_ERROR_LOG, EVENT_LOG);
        excelSerialize.SerializeAllExcel();
        EditorUtility.ClearProgressBar();
    }

    [MenuItem("Tools/Excel/Generate StaticDataModel", false, 1)]
    static void GenerateStaticDataModel()
    {
        var dataModelGenerate = new DataModelGenerate();
        dataModelGenerate.Init(EVENT_END_PROGRESS, EVENT_PROGRESS, EVENT_ERROR_LOG, EVENT_LOG);
        dataModelGenerate.GenerateStaticDataModel();
        EditorUtility.ClearProgressBar();
    }
    private static void EVENT_LOG(string log) => Debug.Log(log);
    private static void EVENT_ERROR_LOG(string error) => Debug.LogError(error);
    private static void EVENT_END_PROGRESS() => EditorUtility.ClearProgressBar();
    private static void EVENT_PROGRESS(string title, float p, string a, string n) =>
        EditorUtility.DisplayProgressBar(title, $"{p * 100f}%, {a} => {n}", p);
}
