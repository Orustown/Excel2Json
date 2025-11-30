using System.Data;
using System.Text;
using System.Linq;
using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Excel2Json.Core;

/// <summary>
/// Core conversion service: reads an Excel/CSV workbook and writes JSON output based on provided options.
/// 核心转换服务：按配置读取 Excel/CSV 并输出 JSON，可供 CLI/GUI 复用。
/// </summary>
public sealed class ExcelJsonConverter
{
    public ExcelJsonConverter()
    {
        // Needed for non-UTF8 code pages when running on .NET Core/NET 5+.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public async Task<ConversionResult> ConvertAsync(ConversionOptions options, CancellationToken cancellationToken = default)
    {
        // 先构建预览（带格式化），写同样的内容到文件，确保导出与预览一致。
        var preview = BuildPreview(options, cancellationToken);
        var encoding = Encoding.GetEncoding(options.EncodingName);

        var directory = Path.GetDirectoryName(options.OutputPath);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);

        var tempPath = Path.Combine(string.IsNullOrWhiteSpace(directory) ? Environment.CurrentDirectory : directory, Path.GetRandomFileName());
        await File.WriteAllTextAsync(tempPath, preview.Json, encoding, cancellationToken).ConfigureAwait(false);
        File.Move(tempPath, options.OutputPath, overwrite: true);

        return new ConversionResult(options.OutputPath, preview.Sheets, preview.Rows);
    }

    /// <summary>
    /// Generate JSON without writing to disk so the UI can show a live preview.
    /// 不输出文件的快速预览，文件对象与导出完全一致。
    /// </summary>
    public Task<ConversionPreview> PreviewAsync(ConversionOptions options, CancellationToken cancellationToken = default)
    {
        // 提供快速预览数据（不落盘），UI/CLI 可共用。
        var preview = BuildPreview(options, cancellationToken);
        return Task.FromResult(preview);
    }

    /// <summary>
    /// Detect all sheet names in the workbook for UI selection.
    /// 检测工作簿中的所有 Sheet 名，便于 UI 下拉选择。
    /// </summary>
    public IReadOnlyList<string> GetSheetNames(string excelPath, int headerRows = 3)
    {
        // 读取工作簿，列出 Sheet 名以供下拉选择。
        var dataSet = LoadWorkbook(excelPath, headerRows, null);
        return dataSet.Tables.Cast<DataTable>().Select(t => t.TableName).ToList();
    }

    private static DataSet LoadWorkbook(string excelPath, int headerRows, string? sheetName)
    {
        // 根据扩展名选择 Excel 或 CSV 读取，并可按名称截取单个 Sheet。
        var ext = Path.GetExtension(excelPath).ToLowerInvariant();
        DataSet dataSet;
        if (ext is ".csv")
        {
            dataSet = LoadCsv(excelPath, headerRows);
        }
        else
        {
            using var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            dataSet = reader.AsDataSet(CreateConfig(headerRows));
        }

        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            var target = dataSet.Tables.Cast<DataTable>().FirstOrDefault(t => t.TableName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            if (target == null)
                throw new InvalidOperationException($"指定的 Sheet 不存在: {sheetName}");

            var clone = dataSet.Clone();
            clone.Tables.Add(target.Copy());
            return clone;
        }

        if (dataSet.Tables.Count < 1)
            throw new InvalidOperationException($"Excel/CSV 为空或未找到 Sheet: {excelPath}");

        return dataSet;
    }

    private static ExcelDataSetConfiguration CreateConfig(int headerRows)
    {
        // 配置表头行过滤，确保行号从 headerRows 之后开始输出。
        var tableConfig = new ExcelDataTableConfiguration
        {
            UseHeaderRow = true,
            FilterRow = rowReader => rowReader.Depth >= Math.Max(headerRows - 1, 0)
        };

        return new ExcelDataSetConfiguration
        {
            UseColumnDataType = true,
            ConfigureDataTable = _ => tableConfig
        };
    }

    private static JToken BuildJson(DataSet dataSet, ConversionOptions opt, out int maxDepth)
    {
        // 根据配置构建对象或 Sheet 包装；过滤空表/前缀。
        var validSheets = new List<DataTable>();
        foreach (DataTable sheet in dataSet.Tables)
        {
            if (!string.IsNullOrEmpty(opt.ExcludePrefix) && sheet.TableName.StartsWith(opt.ExcludePrefix, StringComparison.OrdinalIgnoreCase))
                continue;
            if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                validSheets.Add(sheet);
        }

        object? data;
        if (!opt.ForceSheetName && validSheets.Count == 1)
        {
            // 单 Sheet 且不强制包装：直接输出该 Sheet。
            data = ConvertSheet(validSheets[0], opt);
        }
        else
        {
            // 多 Sheet 或强制包装：以 Sheet 名为键。
            data = validSheets.ToDictionary(sheet => sheet.TableName, sheet => ConvertSheet(sheet, opt));
        }

        var token = data is JToken jToken ? jToken : JToken.FromObject(data ?? new object());
        maxDepth = CalculateDepth(token);
        return token;
    }

    private static object ConvertSheet(DataTable sheet, ConversionOptions opt)
    {
        // 输出数组或字典（首列为键）两种模式。
        return opt.ExportArray
            ? ConvertSheetToArray(sheet, opt)
            : ConvertSheetToDict(sheet, opt);
    }

    private static List<object?> ConvertSheetToArray(DataTable sheet, ConversionOptions opt)
    {
        // 将每行转成对象，放入数组。
        var values = new List<object?>();
        var firstDataRow = Math.Max(opt.HeaderRows - 1, 0);
        for (var i = firstDataRow; i < sheet.Rows.Count; i++)
        {
            var row = sheet.Rows[i];
            values.Add(ConvertRowToDict(sheet, row, opt, firstDataRow));
        }
        return values;
    }

    private static Dictionary<string, object?> ConvertSheetToDict(DataTable sheet, ConversionOptions opt)
    {
        // 以首列或自动编号为键，构成字典。
        var map = new Dictionary<string, object?>();
        var firstDataRow = Math.Max(opt.HeaderRows - 1, 0);
        for (var i = firstDataRow; i < sheet.Rows.Count; i++)
        {
            var row = sheet.Rows[i];
            var id = row[sheet.Columns[0]]?.ToString();
            if (string.IsNullOrWhiteSpace(id))
                id = $"row_{i}";

            map[id] = ConvertRowToDict(sheet, row, opt, firstDataRow);
        }

        return map;
    }

    private static Dictionary<string, object?> ConvertRowToDict(DataTable sheet, DataRow row, ConversionOptions opt, int firstDataRow)
    {
        // 按列遍历一行，应用前缀过滤、大小写转换、类型归一化。
        var rowData = new Dictionary<string, object?>();
        var colIndex = 0;
        foreach (DataColumn column in sheet.Columns)
        {
            var columnName = column.ToString() ?? string.Empty;
            if (!string.IsNullOrEmpty(opt.ExcludePrefix) && columnName.StartsWith(opt.ExcludePrefix, StringComparison.OrdinalIgnoreCase))
                continue;

            object? value = row[column];
            value = NormalizeCell(value, sheet, column, row, firstDataRow, opt);

            var fieldName = string.IsNullOrWhiteSpace(columnName) ? $"col_{colIndex}" : columnName;
            if (opt.Lowercase)
                fieldName = fieldName.ToLowerInvariant();

            rowData[fieldName] = value;
            colIndex++;
        }

        return rowData;
    }

    private static object? NormalizeCell(object? value, DataTable sheet, DataColumn column, DataRow row, int firstDataRow, ConversionOptions opt)
    {
        // 空值回落默认；整数化；可解析单元格 JSON；可强制转字符串。
        if (value == null || value is DBNull)
        {
            value = GetColumnDefault(sheet, column, firstDataRow);
        }
        else if (value is double num && (int)num == num)
        {
            value = (int)num;
        }

        if (opt.CellJson && value is string text)
        {
            var cellText = text.Trim();
            if (cellText.StartsWith("[") || cellText.StartsWith("{"))
            {
                try
                {
                    var cellObj = JsonConvert.DeserializeObject(cellText);
                    if (cellObj != null)
                        value = cellObj;
                }
                catch
                {
                    // ignore parse errors and fall back to string
                }
            }
        }

        if (opt.AllString && value is not string)
            value = value?.ToString();

        return value;
    }

    private static object GetColumnDefault(DataTable sheet, DataColumn column, int firstDataRow)
    {
        // 找到首个非空同列值的默认类型，否则空串。
        for (var i = firstDataRow; i < sheet.Rows.Count; i++)
        {
            var value = sheet.Rows[i][column];
            var type = value?.GetType();
            if (type != null && type != typeof(DBNull))
            {
                if (type.IsValueType)
                    return Activator.CreateInstance(type) ?? string.Empty;
                break;
            }
        }

        return string.Empty;
    }

    private static ConversionPreview BuildPreview(ConversionOptions options, CancellationToken cancellationToken)
    {
        // 构建预览并计算行/深度；预览 JSON 已按选项格式化。
        cancellationToken.ThrowIfCancellationRequested();
        var dataSet = LoadWorkbook(options.ExcelPath, options.HeaderRows, options.SheetName);

        cancellationToken.ThrowIfCancellationRequested();
        var token = BuildJson(dataSet, options, out var maxDepth);
        var rows = dataSet.Tables.Cast<DataTable>().Sum(t => Math.Max(0, t.Rows.Count - Math.Max(options.HeaderRows - 1, 0)));

        var formatted = FormatJson(token, options);
        return new ConversionPreview(formatted, dataSet.Tables.Count, rows, maxDepth, options.SingleLineArray);
    }

    public static string FormatJson(JToken token, ConversionOptions opt)
    {
        try
        {
            var settings = new JsonSerializerSettings { DateFormatString = opt.DateFormat, Formatting = Formatting.Indented };

            // 单行数组模式：所有数组元素逐行输出，其余使用标准缩进。
            if (opt.SingleLineArray)
            {
                var sb = new StringBuilder();
                WriteTokenPerLine(token, sb, depth: 0, settings, indentSize: 2);
                return sb.ToString();
            }

            return JsonConvert.SerializeObject(token, settings);
        }
        catch (Exception ex)
        {
            // 序列化失败时输出原始 ToString，附加异常信息便于日志排查。
            return $"// FormatJson failed: {ex.Message}{Environment.NewLine}{token}";
        }
    }

    private static void WriteTokenPerLine(JToken token, StringBuilder sb, int depth, JsonSerializerSettings settings, int indentSize)
    {
        switch (token.Type)
        {
            case JTokenType.Object:
                WriteObjectPerLine((JObject)token, sb, depth, settings, indentSize);
                break;
            case JTokenType.Array:
                WriteArrayPerLine((JArray)token, sb, depth, settings, indentSize);
                break;
            default:
                sb.Append(JsonConvert.SerializeObject(token, settings));
                break;
        }
    }

    private static void WriteObjectPerLine(JObject obj, StringBuilder sb, int depth, JsonSerializerSettings settings, int indentSize)
    {
        sb.Append('{');
        var properties = obj.Properties().ToList();
        if (properties.Count == 0)
        {
            sb.Append('}');
            return;
        }
        
        for (var i = 0; i < properties.Count; i++)
        {
            var prop = properties[i];
            sb.Append(JsonConvert.SerializeObject(prop.Name, settings));
            sb.Append(": ");
            WriteTokenPerLine(prop.Value, sb, depth + 1, settings, indentSize);
            if (i < properties.Count - 1)
                sb.Append(',');
        }
        sb.Append('}');
    }

    private static void WriteArrayPerLine(JArray array, StringBuilder sb, int depth, JsonSerializerSettings settings, int indentSize)
    {
        sb.Append('[');
        if (array.Count == 0)
        {
            sb.Append(']');
            return;
        }

        sb.Append(Environment.NewLine);
        var childIndent = new string(' ', (depth + 1) * indentSize);
        var parentIndent = new string(' ', depth * indentSize);
        for (var i = 0; i < array.Count; i++)
        {
            sb.Append(childIndent);
            WriteTokenPerLine(array[i], sb, depth + 1, settings, indentSize);
            if (i < array.Count - 1)
                sb.Append(',');
            sb.Append(Environment.NewLine);
        }
        sb.Append(parentIndent);
        sb.Append(']');
    }

    private static DataSet LoadCsv(string path, int headerRows)
    {
        // 按表头行读取 CSV，构建 DataTable。
        var dataSet = new DataSet();
        var table = new DataTable(Path.GetFileNameWithoutExtension(path) ?? "csv");
        var rows = new List<string[]>();

        using var parser = new TextFieldParser(path, Encoding.UTF8)
        {
            TextFieldType = FieldType.Delimited,
            HasFieldsEnclosedInQuotes = true
        };
        parser.SetDelimiters(",");

        while (!parser.EndOfData)
        {
            rows.Add(parser.ReadFields() ?? Array.Empty<string>());
        }

        if (rows.Count == 0)
        {
            dataSet.Tables.Add(table);
            return dataSet;
        }

        var headerIndex = Math.Min(Math.Max(headerRows - 1, 0), rows.Count - 1);
        var header = rows[headerIndex];
        var columnCount = header.Length;

        for (var i = 0; i < columnCount; i++)
        {
            var name = string.IsNullOrWhiteSpace(header[i]) ? $"col_{i}" : header[i];
            table.Columns.Add(name);
        }

        for (var i = headerIndex + 1; i < rows.Count; i++)
        {
            var row = table.NewRow();
            var fields = rows[i];
            for (var c = 0; c < columnCount; c++)
            {
                row[c] = c < fields.Length ? fields[c] : string.Empty;
            }
            table.Rows.Add(row);
        }

        dataSet.Tables.Add(table);
        return dataSet;
    }

    private static int CalculateDepth(JToken token)
    {
        // 计算 JSON 最大嵌套深度，便于 UI 展示。
        if (token is not JContainer container || !container.HasValues)
            return 1;

        var depths = container.Children().Select(CalculateDepth).ToList();
        return 1 + (depths.Count == 0 ? 0 : depths.Max());
    }

}

public sealed record ConversionResult(string OutputPath, int Sheets, int Rows);

public sealed record ConversionPreview(string Json, int Sheets, int Rows, int MaxDepth, bool SingleLineArrayUsed);
