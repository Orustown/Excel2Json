using System.Data;
using System.Text;
using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json;

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
        var preview = BuildPreview(options, cancellationToken);
        var encoding = Encoding.GetEncoding(options.EncodingName);

        Directory.CreateDirectory(Path.GetDirectoryName(options.OutputPath) ?? ".");
        await File.WriteAllTextAsync(options.OutputPath, preview.Json, encoding, cancellationToken).ConfigureAwait(false);

        return new ConversionResult(options.OutputPath, preview.Sheets, preview.Rows);
    }

    /// <summary>
    /// Generate JSON without writing to disk so the UI can show a live preview.
    /// 不输出文件的快速预览，文件对象与导出完全一致。
    /// </summary>
    public Task<ConversionPreview> PreviewAsync(ConversionOptions options, CancellationToken cancellationToken = default)
    {
        var preview = BuildPreview(options, cancellationToken);
        return Task.FromResult(preview);
    }

    /// <summary>
    /// Detect all sheet names in the workbook for UI selection.
    /// 检测工作簿中的所有 Sheet 名，便于 UI 下拉选择。
    /// </summary>
    public IReadOnlyList<string> GetSheetNames(string excelPath, int headerRows = 3)
    {
        var dataSet = LoadWorkbook(excelPath, headerRows, null);
        return dataSet.Tables.Cast<DataTable>().Select(t => t.TableName).ToList();
    }

    private static DataSet LoadWorkbook(string excelPath, int headerRows, string? sheetName)
    {
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

    private static string BuildJson(DataSet dataSet, ConversionOptions opt)
    {
        var validSheets = new List<DataTable>();
        foreach (DataTable sheet in dataSet.Tables)
        {
            if (!string.IsNullOrEmpty(opt.ExcludePrefix) && sheet.TableName.StartsWith(opt.ExcludePrefix, StringComparison.OrdinalIgnoreCase))
                continue;
            if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                validSheets.Add(sheet);
        }

        var jsonSettings = new JsonSerializerSettings
        {
            DateFormatString = opt.DateFormat,
            Formatting = opt.SingleLineArray ? Formatting.None : Formatting.Indented
        };

        if (!opt.ForceSheetName && validSheets.Count == 1)
        {
            var single = ConvertSheet(validSheets[0], opt);
            return JsonConvert.SerializeObject(single, jsonSettings);
        }

        var data = new Dictionary<string, object?>();
        foreach (var sheet in validSheets)
        {
            data[sheet.TableName] = ConvertSheet(sheet, opt);
        }

        return JsonConvert.SerializeObject(data, jsonSettings);
    }

    private static object ConvertSheet(DataTable sheet, ConversionOptions opt)
    {
        return opt.ExportArray
            ? ConvertSheetToArray(sheet, opt)
            : ConvertSheetToDict(sheet, opt);
    }

    private static List<object?> ConvertSheetToArray(DataTable sheet, ConversionOptions opt)
    {
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
        cancellationToken.ThrowIfCancellationRequested();
        var dataSet = LoadWorkbook(options.ExcelPath, options.HeaderRows, options.SheetName);

        cancellationToken.ThrowIfCancellationRequested();
        var json = BuildJson(dataSet, options);
        var rows = dataSet.Tables.Cast<DataTable>().Sum(t => Math.Max(0, t.Rows.Count - Math.Max(options.HeaderRows - 1, 0)));

        return new ConversionPreview(json, dataSet.Tables.Count, rows);
    }

    private static DataSet LoadCsv(string path, int headerRows)
    {
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
}

public sealed record ConversionResult(string OutputPath, int Sheets, int Rows);

public sealed record ConversionPreview(string Json, int Sheets, int Rows);
