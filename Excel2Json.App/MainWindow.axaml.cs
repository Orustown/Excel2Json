using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Threading;
using Excel2Json.Core;
using Newtonsoft.Json.Linq;

namespace Excel2Json.App;

public partial class MainWindow : Window
{
    private readonly ExcelJsonConverter _converter = new();
    private Button _convertButton = null!;
    private Button _browseExcelButton = null!;
    private Button _browseJsonButton = null!;
    private TextBox _excelPathBox = null!;
    private TextBox _jsonPathBox = null!;
    private NumericUpDown _headerRowsBox = null!;
    private ComboBox _encodingBox = null!;
    private ComboBox _dateFormatBox = null!;
    private TextBox _excludePrefixBox = null!;
    private ComboBox _sheetNameBox = null!;
    private CheckBox _lowercaseBox = null!;
    private CheckBox _arrayBox = null!;
    private CheckBox _forceSheetBox = null!;
    private CheckBox _cellJsonBox = null!;
    private CheckBox _allStringBox = null!;
    private CheckBox _singleLineArrayBox = null!;
    private TextBlock _statusText = null!;
    private TextBox _logBox = null!;
    private TextBlock _previewStatusText = null!;
    private TextBlock _previewBlock = null!;
    private CancellationTokenSource? _previewCts;
    private const string DefaultSheetPlaceholder = "全部 Sheet（默认）";

    public MainWindow()
    {
        InitializeComponent();
        BindControls();
        BindOptionChangeHandlers();
        _ = UpdateSheetComboAsync();

        _convertButton.Click += OnConvertClicked;
        _browseExcelButton.Click += OnBrowseExcel;
        _browseJsonButton.Click += OnBrowseJson;
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
        AddHandler(DragDrop.DropEvent, OnDrop);

        SchedulePreviewRefresh();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private void BindControls()
    {
        _convertButton = this.FindControl<Button>("ConvertButton") ?? throw new InvalidOperationException("ConvertButton not found");
        _browseExcelButton = this.FindControl<Button>("BrowseExcelButton") ?? throw new InvalidOperationException("BrowseExcelButton not found");
        _browseJsonButton = this.FindControl<Button>("BrowseJsonButton") ?? throw new InvalidOperationException("BrowseJsonButton not found");
        _excelPathBox = this.FindControl<TextBox>("ExcelPathBox") ?? throw new InvalidOperationException("ExcelPathBox not found");
        _jsonPathBox = this.FindControl<TextBox>("JsonPathBox") ?? throw new InvalidOperationException("JsonPathBox not found");
        _headerRowsBox = this.FindControl<NumericUpDown>("HeaderRowsBox") ?? throw new InvalidOperationException("HeaderRowsBox not found");
        _encodingBox = this.FindControl<ComboBox>("EncodingBox") ?? throw new InvalidOperationException("EncodingBox not found");
        _dateFormatBox = this.FindControl<ComboBox>("DateFormatBox") ?? throw new InvalidOperationException("DateFormatBox not found");
        _excludePrefixBox = this.FindControl<TextBox>("ExcludePrefixBox") ?? throw new InvalidOperationException("ExcludePrefixBox not found");
        _sheetNameBox = this.FindControl<ComboBox>("SheetNameBox") ?? throw new InvalidOperationException("SheetNameBox not found");
        _lowercaseBox = this.FindControl<CheckBox>("LowercaseBox") ?? throw new InvalidOperationException("LowercaseBox not found");
        _arrayBox = this.FindControl<CheckBox>("ArrayBox") ?? throw new InvalidOperationException("ArrayBox not found");
        _forceSheetBox = this.FindControl<CheckBox>("ForceSheetBox") ?? throw new InvalidOperationException("ForceSheetBox not found");
        _cellJsonBox = this.FindControl<CheckBox>("CellJsonBox") ?? throw new InvalidOperationException("CellJsonBox not found");
        _allStringBox = this.FindControl<CheckBox>("AllStringBox") ?? throw new InvalidOperationException("AllStringBox not found");
        _singleLineArrayBox = this.FindControl<CheckBox>("SingleLineArrayBox") ?? throw new InvalidOperationException("SingleLineArrayBox not found");
        _statusText = this.FindControl<TextBlock>("StatusText") ?? throw new InvalidOperationException("StatusText not found");
        _logBox = this.FindControl<TextBox>("LogBox") ?? throw new InvalidOperationException("LogBox not found");
        _previewStatusText = this.FindControl<TextBlock>("PreviewStatusText") ?? throw new InvalidOperationException("PreviewStatusText not found");
        _previewBlock = this.FindControl<TextBlock>("PreviewBlock") ?? throw new InvalidOperationException("PreviewBlock not found");
    }

    private void BindOptionChangeHandlers()
    {
        void ListenBox(TextBox box) => box.PropertyChanged += OnOptionChanged;
        void ListenCheck(CheckBox box) => box.PropertyChanged += OnOptionChanged;
        void ListenSpin(NumericUpDown box) => box.PropertyChanged += OnOptionChanged;
        void ListenCombo(ComboBox box)
        {
            box.PropertyChanged += OnOptionChanged;
            box.SelectionChanged += (_, _) => SchedulePreviewRefresh();
        }

        ListenBox(_excelPathBox);
        ListenBox(_jsonPathBox);
        ListenSpin(_headerRowsBox);
        ListenCombo(_encodingBox);
        ListenCombo(_dateFormatBox);
        ListenBox(_excludePrefixBox);
        ListenCombo(_sheetNameBox);
        ListenCheck(_lowercaseBox);
        ListenCheck(_arrayBox);
        ListenCheck(_forceSheetBox);
        ListenCheck(_cellJsonBox);
        ListenCheck(_allStringBox);
        ListenCheck(_singleLineArrayBox);
    }

    private void OnOptionChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (e.Property == TextBox.TextProperty
            || e.Property == ComboBox.SelectedItemProperty
            || e.Property == NumericUpDown.ValueProperty
            || e.Property == ToggleButton.IsCheckedProperty)
            SchedulePreviewRefresh();
    }

    private async void OnBrowseExcel(object? sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = false,
            Filters = new List<FileDialogFilter>
            {
                new() { Name = "Excel/CSV", Extensions = { "xlsx", "xls", "csv" } },
                new() { Name = "All", Extensions = { "*" } }
            }
        };
        var result = await dialog.ShowAsync(this);
        if (result?.Length > 0)
        {
            _excelPathBox.Text = result[0];
            if (string.IsNullOrWhiteSpace(_jsonPathBox.Text))
            {
                var candidate = Path.ChangeExtension(result[0], ".json");
                _jsonPathBox.Text = candidate;
            }

            await UpdateSheetComboAsync();
            SchedulePreviewRefresh();
        }
    }

    private async void OnBrowseJson(object? sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            DefaultExtension = "json",
            Filters = new List<FileDialogFilter>
            {
                new() { Name = "JSON", Extensions = { "json" } },
                new() { Name = "All", Extensions = { "*" } }
            }
        };

        dialog.InitialFileName = !string.IsNullOrWhiteSpace(_jsonPathBox.Text)
            ? Path.GetFileName(_jsonPathBox.Text)
            : "output.json";

        var result = await dialog.ShowAsync(this);
        if (!string.IsNullOrWhiteSpace(result))
        {
            _jsonPathBox.Text = result;
            SchedulePreviewRefresh();
        }
    }

    private async void OnConvertClicked(object? sender, RoutedEventArgs e)
    {
        try
        {
            _statusText.Text = "正在导出...";
            _convertButton.IsEnabled = false;
            Log("开始转换");

            var options = BuildOptions();
            var summary = await _converter.ConvertAsync(options);

            _statusText.Text = $"完成：{summary.OutputPath}";
            Log($"完成，Sheet: {summary.Sheets}，行: {summary.Rows}");
            SchedulePreviewRefresh();
        }
        catch (Exception ex)
        {
            _statusText.Text = "失败";
            Log($"错误：{ex.Message}");
            await new MessageBox("导出失败", ex.Message).ShowDialog(this);
        }
        finally
        {
            _convertButton.IsEnabled = true;
        }
    }

    private void OnDragOver(object? sender, DragEventArgs e)
    {
        if (e.Data.Contains(DataFormats.Files))
        {
            e.DragEffects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private async void OnDrop(object? sender, DragEventArgs e)
    {
        if (!e.Data.Contains(DataFormats.Files)) return;
        var files = e.Data.GetFiles();
        var first = files?.FirstOrDefault();
        if (first == null) return;

        var localPath = first.Path.LocalPath;
        var ext = Path.GetExtension(localPath).ToLowerInvariant();
        if (ext is ".xlsx" or ".xls" or ".csv")
        {
            _excelPathBox.Text = localPath;
            if (string.IsNullOrWhiteSpace(_jsonPathBox.Text))
                _jsonPathBox.Text = Path.ChangeExtension(localPath, ".json");
            Log($"已选择文件：{localPath}");
            await UpdateSheetComboAsync();
            SchedulePreviewRefresh();
        }
    }

    private ConversionOptions BuildOptions(bool requireOutputPath = true)
    {
        var headerRows = (int)Math.Max(1, Math.Round(_headerRowsBox.Value ?? 3));

        var excelPath = _excelPathBox.Text?.Trim() ?? string.Empty;
        var jsonPath = _jsonPathBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            throw new InvalidOperationException("请选择有效的 Excel/CSV 文件路径。");
        if (requireOutputPath && string.IsNullOrWhiteSpace(jsonPath))
            throw new InvalidOperationException("请设置输出 JSON 路径。");
        if (!requireOutputPath && string.IsNullOrWhiteSpace(jsonPath))
            jsonPath = Path.ChangeExtension(excelPath, ".json");

        return new ConversionOptions(
            ExcelPath: excelPath,
            OutputPath: jsonPath,
            HeaderRows: Math.Max(headerRows, 1),
            Lowercase: _lowercaseBox.IsChecked ?? false,
            ExportArray: _arrayBox.IsChecked ?? false,
            EncodingName: GetComboValue(_encodingBox, "utf-8"),
            DateFormat: GetComboValue(_dateFormatBox, "yyyy-MM-dd HH:mm:ss"),
            ForceSheetName: _forceSheetBox.IsChecked ?? false,
            ExcludePrefix: _excludePrefixBox.Text?.Trim() ?? string.Empty,
            CellJson: _cellJsonBox.IsChecked ?? false,
            AllString: _allStringBox.IsChecked ?? false,
            SingleLineArray: _singleLineArrayBox.IsChecked ?? false,
            SheetName: GetSheetSelection()
        );
    }

    private void SchedulePreviewRefresh()
    {
        _previewCts?.Cancel();
        _previewCts = new CancellationTokenSource();
        var token = _previewCts.Token;
        _ = RefreshPreviewAsync(token);
    }

    private async Task RefreshPreviewAsync(CancellationToken token)
    {
        try
        {
            UpdatePreviewStatus("预览中...");
            await Task.Delay(200, token);
            var excelPath = _excelPathBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    SetPreviewPlain("// 请选择 Excel/CSV 文件以生成预览。");
                    _previewStatusText.Text = "等待选择文件...";
                }, DispatcherPriority.Background);
                return;
            }
            if (!File.Exists(excelPath))
            {
                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    SetPreviewPlain($"// 未找到文件：{excelPath}");
                    _previewStatusText.Text = "等待可用文件...";
                }, DispatcherPriority.Background);
                return;
            }

            var options = BuildOptions(requireOutputPath: false);
            var preview = await _converter.PreviewAsync(options, token);

            token.ThrowIfCancellationRequested();
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                RenderHighlightedJson(preview.Json);
                _previewStatusText.Text = $"预览就绪：Sheet {preview.Sheets}，行 {preview.Rows}";
            }, DispatcherPriority.Background);
        }
        catch (OperationCanceledException)
        {
            // ignore
        }
        catch (Exception ex)
        {
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                SetPreviewPlain($"// 预览失败：{ex.Message}");
                _previewStatusText.Text = "预览失败";
            }, DispatcherPriority.Background);
        }
    }

    private void UpdatePreviewStatus(string status)
    {
        _previewStatusText.Text = status;
    }

    private void Log(string message)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss");
        var builder = new StringBuilder(_logBox.Text ?? string.Empty);
        builder.AppendLine($"[{timestamp}] {message}");
        _logBox.Text = builder.ToString();
        _logBox.CaretIndex = _logBox.Text.Length;
    }

    private void SetPreviewPlain(string text)
    {
        if (_previewBlock is null) return;
        var block = _previewBlock;
        block.Inlines.Clear();
        block.Text = text;
        block.TextAlignment = TextAlignment.Center;
        block.TextWrapping = TextWrapping.Wrap;
    }

    private static string GetComboValue(ComboBox combo, string fallback = "")
    {
        if (combo.SelectedItem is string s)
            return s.Trim();
        if (combo.SelectedItem is ComboBoxItem item && item.Content is string content)
            return content.Trim();
        return fallback;
    }

    private string? GetSheetSelection()
    {
        var value = GetComboValue(_sheetNameBox, string.Empty);
        if (string.IsNullOrWhiteSpace(value) || value == DefaultSheetPlaceholder)
            return null;
        return value;
    }

    private async Task UpdateSheetComboAsync()
    {
        var path = _excelPathBox.Text?.Trim();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            _sheetNameBox.ItemsSource = new List<string> { DefaultSheetPlaceholder };
            _sheetNameBox.SelectedIndex = 0;
            return;
        }

        try
        {
            var headerRows = (int)Math.Max(1, Math.Round(_headerRowsBox.Value ?? 3));
            var sheets = await Task.Run(() => _converter.GetSheetNames(path, headerRows));
            var items = new List<string> { DefaultSheetPlaceholder };
            items.AddRange(sheets);
            _sheetNameBox.ItemsSource = items;
            _sheetNameBox.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            Log($"获取 Sheet 列表失败：{ex.Message}");
        }
    }

    private void RenderHighlightedJson(string json)
    {
        if (_previewBlock is null) return;
        var block = _previewBlock;
        block.Inlines.Clear();
        block.TextAlignment = TextAlignment.Left;
        block.TextWrapping = TextWrapping.NoWrap;
        try
        {
            var token = JToken.Parse(json);
            AppendToken(token, block.Inlines, 0);
        }
        catch
        {
            block.Text = json;
        }
    }

    private static readonly IBrush KeyBrush = new SolidColorBrush(Color.FromRgb(180, 200, 255));
    private static readonly IBrush StringBrush = new SolidColorBrush(Color.FromRgb(190, 225, 180));
    private static readonly IBrush NumberBrush = new SolidColorBrush(Color.FromRgb(255, 210, 140));
    private static readonly IBrush BooleanBrush = new SolidColorBrush(Color.FromRgb(255, 170, 170));
    private static readonly IBrush NullBrush = new SolidColorBrush(Color.FromRgb(150, 155, 165));
    private static readonly IBrush PunctuationBrush = new SolidColorBrush(Color.FromRgb(120, 130, 150));

    private void AppendToken(JToken token, InlineCollection inlines, int indent)
    {
        switch (token.Type)
        {
            case JTokenType.Object:
                AppendObject((JObject)token, inlines, indent);
                break;
            case JTokenType.Array:
                AppendArray((JArray)token, inlines, indent);
                break;
            case JTokenType.String:
                inlines.Add(new Run($"\"{token.Value<string>()}\"") { Foreground = StringBrush });
                break;
            case JTokenType.Integer:
            case JTokenType.Float:
                inlines.Add(new Run(token.ToString()) { Foreground = NumberBrush });
                break;
            case JTokenType.Boolean:
                inlines.Add(new Run(token.ToString().ToLowerInvariant()) { Foreground = BooleanBrush });
                break;
            case JTokenType.Null:
            case JTokenType.Undefined:
                inlines.Add(new Run("null") { Foreground = NullBrush });
                break;
            default:
                inlines.Add(new Run(token.ToString()) { Foreground = StringBrush });
                break;
        }
    }

    private void AppendObject(JObject obj, InlineCollection inlines, int indent)
    {
        inlines.Add(new Run("{") { Foreground = PunctuationBrush });
        var properties = obj.Properties().ToList();
        if (properties.Count == 0)
        {
            inlines.Add(new Run("}") { Foreground = PunctuationBrush });
            return;
        }

        inlines.Add(new Run("\n"));
        for (var i = 0; i < properties.Count; i++)
        {
            var prop = properties[i];
            inlines.Add(new Run(new string(' ', (indent + 1) * 2)));
            inlines.Add(new Run($"\"{prop.Name}\"") { Foreground = KeyBrush });
            inlines.Add(new Run(": ") { Foreground = PunctuationBrush });
            AppendToken(prop.Value, inlines, indent + 1);
            if (i < properties.Count - 1)
                inlines.Add(new Run(",") { Foreground = PunctuationBrush });
            inlines.Add(new Run("\n"));
        }
        inlines.Add(new Run(new string(' ', indent * 2)));
        inlines.Add(new Run("}") { Foreground = PunctuationBrush });
    }

    private void AppendArray(JArray array, InlineCollection inlines, int indent)
    {
        inlines.Add(new Run("[") { Foreground = PunctuationBrush });
        if (array.Count == 0)
        {
            inlines.Add(new Run("]") { Foreground = PunctuationBrush });
            return;
        }

        var singleLine = _singleLineArrayBox.IsChecked ?? false;
        if (singleLine)
        {
            for (var i = 0; i < array.Count; i++)
            {
                // 将数组压缩为单行：逐项拼接，不换行。
                var tempHost = new TextBlock();
                AppendToken(array[i], tempHost.Inlines, indent + 1);
                foreach (var inline in tempHost.Inlines)
                    inlines.Add(inline);

                if (i < array.Count - 1)
                    inlines.Add(new Run(", ") { Foreground = PunctuationBrush });
            }

            inlines.Add(new Run("]") { Foreground = PunctuationBrush });
            return;
        }

        inlines.Add(new Run("\n"));
        for (var i = 0; i < array.Count; i++)
        {
            inlines.Add(new Run(new string(' ', (indent + 1) * 2)));
            AppendToken(array[i], inlines, indent + 1);
            if (i < array.Count - 1)
                inlines.Add(new Run(",") { Foreground = PunctuationBrush });
            inlines.Add(new Run("\n"));
        }
        inlines.Add(new Run(new string(' ', indent * 2)));
        inlines.Add(new Run("]") { Foreground = PunctuationBrush });
    }
}
