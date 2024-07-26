using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using UnityEngine;
using OfficeOpenXml;

/// <summary>
/// Excel 文件存储和读取器
/// </summary>
public partial class ExcelSheet
{
    public static string SAVE_PATH = Application.streamingAssetsPath + "/Excel/";

    private int _rowCount = 0; // 最大行数

    private int _colCount = 0; // 最大列数

    public int RowCount { get => _rowCount; }

    public int ColCount { get => _colCount; }

    private Dictionary<Index, string> _sheetDic = new Dictionary<Index, string>(); // 缓存当前数据的字典

    public ExcelSheet() { }

    public ExcelSheet(string filePath, string sheetName = null, FileFormat format = FileFormat.Xlsx) {
        Load(filePath, sheetName, format);
    }

    public string this[int row, int col] {
        get {
            // 越界检查
            if (row >= _rowCount || row < 0)
                Debug.LogError($"ExcelSheet: Row {row} out of range!");
            if (col >= _colCount || col < 0)
                Debug.LogError($"ExcelSheet: Column {col} out of range!");

            // 不存在结果，则返回空字符串
            return _sheetDic.GetValueOrDefault(new Index(row, col), "");
        }
        set {
            _sheetDic[new Index(row, col)] = value;

            // 记录最大行数和列数
            if (row >= _rowCount) _rowCount = row + 1;
            if (col >= _colCount) _colCount = col + 1;
        }
    }

    /// <summary>
    /// 存储 Excel 文件
    /// </summary>
    /// <param name="filePath">文件路径，不需要写文件扩展名</param>
    /// <param name="sheetName">表名，如果没有指定表名，则使用文件名。若使用 csv 格式，则忽略此参数</param>
    /// <param name="format">保存的文件格式</param>
    public void Save(string filePath, string sheetName = null, FileFormat format = FileFormat.Xlsx) {
        string fullPath  = SAVE_PATH + filePath + FileFormatToExtension(format); // 文件完整路径
        var    index     = fullPath.LastIndexOf("/", StringComparison.Ordinal);
        var    directory = fullPath[..index];

        if (!Directory.Exists(directory)) { // 如果文件所在的目录不存在，则先创建目录
            Directory.CreateDirectory(directory);
        }

        switch (format) {
            case FileFormat.Xlsx:
                SaveAsXlsx(fullPath, sheetName);
                break;
            case FileFormat.Csv:
                SaveAsCsv(fullPath);
                break;
            default: throw new ArgumentOutOfRangeException(nameof(format), format, null);
        }

        Debug.Log($"ExcelSheet: Save sheet \"{filePath}::{sheetName}\" successfully.");
    }

    /// <summary>
    /// 读取 Excel 文件
    /// </summary>
    /// <param name="filePath">文件路径，不需要写文件扩展名</param>
    /// <param name="sheetName">表名，如果没有指定表名，则使用文件名</param>
    /// <param name="format">保存的文件格式</param>
    public void Load(string filePath, string sheetName = null, FileFormat format = FileFormat.Xlsx) {
        // 清空当前数据
        Clear();
        string fullPath = SAVE_PATH + filePath + FileFormatToExtension(format); // 文件完整路径

        if (!File.Exists(fullPath)) { // 不存在文件，则报错
            Debug.LogError($"ExcelSheet: Can't find path \"{fullPath}\".");
            return;
        }

        switch (format) {
            case FileFormat.Xlsx:
                LoadFromXlsx(fullPath, sheetName);
                break;
            case FileFormat.Csv:
                LoadFromCsv(fullPath);
                break;
            default: throw new ArgumentOutOfRangeException(nameof(format), format, null);
        }

        Debug.Log($"ExcelSheet: Load sheet \"{filePath}::{sheetName}\" successfully.");
    }

    public void Clear() {
        _sheetDic.Clear();
        _rowCount = 0;
        _colCount = 0;
    }
}

public partial class ExcelSheet
{
    public struct Index
    {
        public int Row;
        public int Col;

        public Index(int row, int col) {
            Row = row;
            Col = col;
        }
    }

    /// <summary>
    /// 保存的文件格式
    /// </summary>
    public enum FileFormat
    {
        Xlsx,
        Csv
    }

    private string FileFormatToExtension(FileFormat format) {
        return $".{format.ToString().ToLower()}";
    }

    private void SaveAsXlsx(string fullPath, string sheetName) {
        var index    = fullPath.LastIndexOf("/", StringComparison.Ordinal);
        var fileName = fullPath[(index + 1)..];
        sheetName ??= fileName[..fileName.IndexOf(".", StringComparison.Ordinal)]; // 如果没有指定表名，则使用文件名

        var       fileInfo = new FileInfo(fullPath);
        using var package  = new ExcelPackage(fileInfo);

        if (!File.Exists(fullPath) ||                         // 不存在 Excel
            package.Workbook.Worksheets[sheetName] == null) { // 或者没有表，则添加表
            package.Workbook.Worksheets.Add(sheetName);       // 创建表时，Excel 文件也会被创建
        }

        var sheet = package.Workbook.Worksheets[sheetName];

        var cells = sheet.Cells;
        cells.Clear(); // 先清空数据

        foreach (var pair in _sheetDic) {
            var i = pair.Key.Row;
            var j = pair.Key.Col;
            cells[i + 1, j + 1].Value = pair.Value;
        }

        package.Save(); // 保存文件
    }

    private void SaveAsCsv(string fullPath) {
        using FileStream fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write);

        Index idx = new Index(0, 0);
        for (int i = 0; i < _rowCount; i++) {
            idx.Row = i;
            idx.Col = 0;

            // 写入第一个 value
            var value = _sheetDic.GetValueOrDefault(idx, "");
            if (!string.IsNullOrEmpty(value))
                fs.Write(Encoding.UTF8.GetBytes(value));

            // 写入后续 value，需要添加 ","
            for (int j = 1; j < _colCount; j++) {
                idx.Col = j;
                value   = "," + _sheetDic.GetValueOrDefault(idx, "");
                fs.Write(Encoding.UTF8.GetBytes(value));
            }

            // 写入 "\n"
            fs.Write(Encoding.UTF8.GetBytes("\n"));
        }
    }

    private void LoadFromXlsx(string fullPath, string sheetName) {
        var index    = fullPath.LastIndexOf("/", StringComparison.Ordinal);
        var fileName = fullPath[(index + 1)..];
        sheetName ??= fileName[..fileName.IndexOf(".", StringComparison.Ordinal)]; // 如果没有指定表名，则使用文件名

        var fileInfo = new FileInfo(fullPath);

        using var package = new ExcelPackage(fileInfo);

        var sheet = package.Workbook.Worksheets[sheetName];

        if (sheet == null) { // 不存在表，则报错
            Debug.LogError($"ExcelSheet: Can't find sheet \"{sheetName}\" in file \"{fullPath}\"");
            return;
        }

        _rowCount = sheet.Dimension.Rows;
        _colCount = sheet.Dimension.Columns;

        var cells = sheet.Cells;
        for (int i = 0; i < _rowCount; i++) {
            for (int j = 0; j < _colCount; j++) {
                var value = cells[i + 1, j + 1].Text;
                if (string.IsNullOrEmpty(value)) continue; // 有数据才记录
                _sheetDic.Add(new Index(i, j), value);
            }
        }
    }

    private void LoadFromCsv(string fullPath) {
        // 读取文件
        string[] lines = File.ReadAllLines(fullPath); // 读取所有行
        for (int i = 0; i < lines.Length; i++) {
            string[] line = lines[i].Split(','); // 读取一行，逗号分割
            for (int j = 0; j < line.Length; j++) {
                if (line[j] != "") // 有数据才记录
                    _sheetDic.Add(new Index(i, j), line[j]);
            }

            // 更新最大行数和列数
            _colCount = Mathf.Max(_colCount, line.Length);
            _rowCount = i + 1;
        }
    }
}