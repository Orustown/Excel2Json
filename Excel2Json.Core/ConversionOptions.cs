namespace Excel2Json.Core;

// 转换参数模型：CLI/GUI 共用，保持选项一致。
public sealed record ConversionOptions(
    string ExcelPath,
    string OutputPath,
    int HeaderRows = 3,
    bool Lowercase = false,
    bool ExportArray = false,
    string EncodingName = "utf-8",
    string DateFormat = "yyyy-MM-dd HH:mm:ss",
    bool ForceSheetName = false,
    string ExcludePrefix = "",
    bool CellJson = false,
    bool AllString = false,
    bool SingleLineArray = false,
    string? SheetName = null);
