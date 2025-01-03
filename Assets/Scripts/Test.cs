using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

#if UNITY_EDITOR
using UnityEditor;
#endif

public class Test : MonoBehaviour
{
    private ExcelSheet _sheet = new ExcelSheet();

    // Start is called before the first frame update
    void Start()
    {
        var sheet = new ExcelSheet("test");
        sheet.Load("test");
        sheet[0, 0] = "1"; // 第一行第一列赋值为 1
        sheet[1, 2] = "2"; // 第二行第三列赋值为 2

        sheet.Save("test", "Sheet1", ExcelFormat.Csv);
    }

    // Update is called once per frame
    void Update() { }

#if UNITY_EDITOR
    private void OnGUI()
    {
        EditorGUILayout.BeginHorizontal();

        if (GUILayout.Button("Save", GUILayout.Width(200), GUILayout.Height(80)))
        {
            _sheet[0, 1] = "Hello World";
            _sheet[1, 2] = "123";

            _sheet.Save("test");
        }

        if (GUILayout.Button("Load", GUILayout.Width(200), GUILayout.Height(80)))
        {
            _sheet.Load("test");

            for (int i = _sheet.Start.x; i < _sheet.End.x; i++)
            {
                for (int j = _sheet.Start.y; j < _sheet.End.y; j++)
                {
                    var value = _sheet[i, j];
                    if (string.IsNullOrEmpty(value)) continue;
                    Debug.Log($"Sheet[{i}, {j}]: {value}");
                }
            }

            Debug.Log(_sheet.Start);
            Debug.Log(_sheet.End);
        }

        EditorGUILayout.EndHorizontal();
    }
#endif
}