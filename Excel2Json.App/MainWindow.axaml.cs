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
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Excel2Json.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

namespace Excel2Json.App;

public partial class MainWindow : Window
{
    private readonly ExcelJsonConverter _converter = new();
    private Button _convertButton = null!;
    private Button _browseExcelButton = null!;
    private Button _browseJsonButton = null!;
    private Button _clearSelectionButton = null!;
    private Button _copyPreviewButton = null!;
    private Button _viewJsonButton = null!;
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
    private TextBlock _previewWatermark = null!;
    private CancellationTokenSource? _previewCts;
    private const string DefaultSheetPlaceholder = "全部 Sheet（默认）";
    private const string PreviewWatermarkText = "// 请拖拽或选择 Excel/CSV 文件以生成预览。";
    private string _lastPreviewText = string.Empty;

    public MainWindow()
    {
        InitializeComponent();
        BindControls();
        BindOptionChangeHandlers();
        _ = UpdateSheetComboAsync();
        ShowPreviewWatermark(PreviewWatermarkText);

        _convertButton.Click += OnConvertClicked;
        _browseExcelButton.Click += OnBrowseExcel;
        _browseJsonButton.Click += OnBrowseJson;
        _clearSelectionButton.Click += OnClearSelection;
        _copyPreviewButton.Click += OnCopyPreview;
        _viewJsonButton.Click += OnViewJson;
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
        _clearSelectionButton = this.FindControl<Button>("ClearSelectionButton") ?? throw new InvalidOperationException("ClearSelectionButton not found");
        _copyPreviewButton = this.FindControl<Button>("CopyPreviewButton") ?? throw new InvalidOperationException("CopyPreviewButton not found");
        _viewJsonButton = this.FindControl<Button>("ViewJsonButton") ?? throw new InvalidOperationException("ViewJsonButton not found");
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
        _previewWatermark = this.FindControl<TextBlock>("PreviewWatermark") ?? throw new InvalidOperationException("PreviewWatermark not found");
        UpdateViewButtonState();
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
        if (!StorageProvider.CanOpen)
        {
            await new MessageBox("不支持", "当前环境不支持文件选择。").ShowDialog(this);
            return;
        }

        var result = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            AllowMultiple = false,
            FileTypeFilter = new List<FilePickerFileType>
            {
                new("Excel/CSV") { Patterns = new[] { "*.xlsx", "*.xls", "*.csv" } },
                new("All files") { Patterns = new[] { "*.*" } }
            }
        });

        var file = result.FirstOrDefault();
        var localPath = file?.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(localPath))
            return;

        _excelPathBox.Text = localPath;
        if (string.IsNullOrWhiteSpace(_jsonPathBox.Text))
        {
            var candidate = Path.ChangeExtension(localPath, ".json");
            _jsonPathBox.Text = candidate;
        }

        await UpdateSheetComboAsync();
        SchedulePreviewRefresh();
    }

    private async void OnBrowseJson(object? sender, RoutedEventArgs e)
    {
        if (!StorageProvider.CanSave)
        {
            await new MessageBox("不支持", "当前环境不支持文件保存。").ShowDialog(this);
            return;
        }

        var result = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            DefaultExtension = "json",
            SuggestedFileName = !string.IsNullOrWhiteSpace(_jsonPathBox.Text)
                ? Path.GetFileName(_jsonPathBox.Text)
                : "output.json",
            FileTypeChoices = new List<FilePickerFileType>
            {
                new("JSON") { Patterns = new[] { "*.json" } },
                new("All files") { Patterns = new[] { "*.*" } }
            }
        });

        var localPath = result?.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(localPath))
            return;

        _jsonPathBox.Text = localPath;
        SchedulePreviewRefresh();
    }

    private void OnClearSelection(object? sender, RoutedEventArgs e)
    {
        _excelPathBox.Text = string.Empty;
        _jsonPathBox.Text = string.Empty;
        ResetComboItems(_sheetNameBox);
        _sheetNameBox.ItemsSource = new List<string> { DefaultSheetPlaceholder };
        _sheetNameBox.SelectedIndex = 0;
        _logBox.Text = string.Empty;
        ShowPreviewWatermark(PreviewWatermarkText);
        _previewStatusText.Text = "等待选择文件...";
        UpdateViewButtonState();
    }

    private async void OnConvertClicked(object? sender, RoutedEventArgs e)
    {
        try
        {
            _statusText.Text = "正在导出...";
            _convertButton.IsEnabled = false;
            Log("开始转换");

            var options = BuildOptions();
            var preview = await _converter.PreviewAsync(options);
            var rootArrayPerLine = options.SingleLineArray && IsRootArray(preview.Json);
            var outputJson = BuildFormattedJson(preview.Json, rootArrayPerLine, options.DateFormat);

            var encoding = Encoding.GetEncoding(options.EncodingName);
            Directory.CreateDirectory(Path.GetDirectoryName(options.OutputPath) ?? ".");
            await File.WriteAllTextAsync(options.OutputPath, outputJson, encoding, CancellationToken.None);

            RenderPreviewText(outputJson);
            _statusText.Text = $"完成：{options.OutputPath}";
            Log($"完成，Sheet: {preview.Sheets}，行: {preview.Rows}");
            SchedulePreviewRefresh();
            UpdateViewButtonState();
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
        UpdateViewButtonState();
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
                    ShowPreviewWatermark(PreviewWatermarkText);
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
            var rootArrayPerLine = options.SingleLineArray && IsRootArray(preview.Json);
            var formatted = BuildFormattedJson(preview.Json, rootArrayPerLine, options.DateFormat);

            token.ThrowIfCancellationRequested();
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                RenderPreviewText(formatted);
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
                _lastPreviewText = string.Empty;
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

    private void RenderPreviewText(string text)
    {
        if (_previewBlock is null) return;
        var block = _previewBlock;
        block.Inlines!.Clear();
        block.Text = text;
        block.TextAlignment = TextAlignment.Left;
        block.TextWrapping = TextWrapping.NoWrap;
        block.HorizontalAlignment = HorizontalAlignment.Stretch;
        block.VerticalAlignment = VerticalAlignment.Stretch;
        block.Foreground = PreviewTextBrush;
        block.Opacity = 1.0;
        _lastPreviewText = text;
        if (_previewWatermark is not null)
            _previewWatermark.IsVisible = false;
    }

    private async void OnCopyPreview(object? sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastPreviewText))
        {
            _previewStatusText.Text = "暂无可复制内容";
            return;
        }

        try
        {
            var clipboard = Clipboard;
            if (clipboard is not null)
            {
                await clipboard.SetTextAsync(_lastPreviewText);
                _previewStatusText.Text = "已复制到剪切板";
            }
            else
            {
                _previewStatusText.Text = "无法访问剪切板";
            }
        }
        catch (Exception ex)
        {
            _previewStatusText.Text = $"复制失败：{ex.Message}";
        }
    }

    private void OnViewJson(object? sender, RoutedEventArgs e)
    {
        var path = _jsonPathBox.Text?.Trim();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            _previewStatusText.Text = "未找到可查看的文件";
            UpdateViewButtonState();
            return;
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            };
            Process.Start(psi);
            _previewStatusText.Text = "已打开文件";
        }
        catch (Exception ex)
        {
            _previewStatusText.Text = $"打开失败：{ex.Message}";
        }
    }

    private void SetPreviewPlain(string text)
    {
        if (_previewBlock is null) return;
        var block = _previewBlock;
        block.Inlines!.Clear();
        block.Text = text;
        block.TextAlignment = TextAlignment.Left;
        block.TextWrapping = TextWrapping.NoWrap;
        block.HorizontalAlignment = HorizontalAlignment.Stretch;
        block.VerticalAlignment = VerticalAlignment.Stretch;
        block.Foreground = PreviewTextBrush;
        block.Opacity = 1.0;
        _lastPreviewText = text;
        if (_previewWatermark is not null)
            _previewWatermark.IsVisible = false;
    }

    private void ShowPreviewWatermark(string text)
    {
        if (_previewWatermark is null || _previewBlock is null) return;
        _previewBlock.Inlines!.Clear();
        _previewBlock.Text = string.Empty;
        _previewWatermark.Text = text;
        _previewWatermark.IsVisible = true;
        _lastPreviewText = string.Empty;
    }

    private static string GetComboValue(ComboBox combo, string fallback = "")
    {
        if (combo.SelectedItem is string s)
            return s.Trim();
        if (combo.SelectedItem is ComboBoxItem item && item.Content is string content)
            return content.Trim();
        return fallback;
    }

    private void UpdateViewButtonState()
    {
        var path = _jsonPathBox.Text?.Trim();
        _viewJsonButton.IsEnabled = !string.IsNullOrWhiteSpace(path) && File.Exists(path);
    }

    private string? GetSheetSelection()
    {
        var value = GetComboValue(_sheetNameBox, string.Empty);
        if (string.IsNullOrWhiteSpace(value) || value == DefaultSheetPlaceholder)
            return null;
        return value;
    }

    private static void ResetComboItems(ComboBox combo)
    {
        combo.ItemsSource = null;
        combo.Items?.Clear();
    }

    private static bool IsRootArray(string json)
    {
        try
        {
            var token = JToken.Parse(json);
            return token is JArray;
        }
        catch
        {
            return false;
        }
    }

    private static string BuildFormattedJson(string json, bool rootArrayPerLine, string dateFormat)
    {
        try
        {
            var token = JToken.Parse(json);
            using var sw = new StringWriter();
            using var writer = new JsonTextWriter(sw) { Formatting = Formatting.None };
            var serializer = JsonSerializer.Create(new JsonSerializerSettings { DateFormatString = dateFormat });
            WriteTokenStream(token, writer, serializer, depth: 1, rootArrayPerLine, isRoot: true);
            writer.Flush();
            return sw.ToString();
        }
        catch
        {
            return json;
        }
    }

    private static void WriteTokenStream(JToken token, JsonTextWriter writer, JsonSerializer serializer, int depth, bool rootArrayPerLine, bool isRoot = false, bool inlineObjects = false)
    {
        switch (token.Type)
        {
            case JTokenType.Object:
                WriteObjectStream((JObject)token, writer, serializer, depth, rootArrayPerLine, isRoot, inlineObjects);
                break;
            case JTokenType.Array:
                WriteArrayStream((JArray)token, writer, serializer, depth, rootArrayPerLine, isRoot, inlineObjects);
                break;
            default:
                serializer.Serialize(writer, token is JValue jValue ? jValue.Value : token.ToString());
                break;
        }
    }

    private static readonly string NewLine = Environment.NewLine;

    private static void WriteObjectStream(JObject obj, JsonTextWriter writer, JsonSerializer serializer, int depth, bool rootArrayPerLine, bool isRoot, bool inlineObjects)
    {
        writer.WriteStartObject();
        var properties = obj.Properties().ToList();
        if (properties.Count == 0)
        {
            writer.WriteEndObject();
            return;
        }

        if (inlineObjects)
        {
            for (var i = 0; i < properties.Count; i++)
            {
                var prop = properties[i];
                if (i > 0)
                    writer.WriteRaw(", ");
                writer.WritePropertyName(prop.Name);
                writer.WriteRaw(": ");
                WriteTokenStream(prop.Value, writer, serializer, depth + 1, rootArrayPerLine, isRoot: false, inlineObjects: true);
            }
            writer.WriteEndObject();
        }
        else
        {
            writer.WriteWhitespace(NewLine);
            var childIndent = new string(' ', depth * 2);
            var parentIndent = new string(' ', (depth - 1) * 2);
            for (var i = 0; i < properties.Count; i++)
            {
                var prop = properties[i];
                writer.WriteWhitespace(childIndent);
                writer.WritePropertyName(prop.Name);
                writer.WriteRaw(": ");
                WriteTokenStream(prop.Value, writer, serializer, depth + 1, rootArrayPerLine, isRoot: false, inlineObjects: false);
                if (i < properties.Count - 1)
                    writer.WriteRaw(",");
                writer.WriteWhitespace(NewLine);
            }
            writer.WriteWhitespace(parentIndent);
            writer.WriteEndObject();
        }
    }

    private static void WriteArrayStream(JArray array, JsonTextWriter writer, JsonSerializer serializer, int depth, bool rootArrayPerLine, bool isRoot, bool inlineObjects)
    {
        if (inlineObjects)
        {
            serializer.Serialize(writer, array);
            return;
        }

        var perLine = rootArrayPerLine && isRoot;
        writer.Formatting = perLine ? Formatting.None : Formatting.Indented;

        if (perLine)
        {
            writer.WriteStartArray();
            if (array.Count == 0)
            {
                writer.WriteEndArray();
                return;
            }

            writer.WriteWhitespace(NewLine);
            var childIndent = new string(' ', depth * 2);
            var parentIndent = new string(' ', (depth - 1) * 2);
            for (var i = 0; i < array.Count; i++)
            {
                writer.WriteWhitespace(childIndent);
                // 根数组每项单行：直接紧凑序列化每个元素
                serializer.Serialize(writer, array[i]);
                if (i < array.Count - 1)
                    writer.WriteRaw(",");
                writer.WriteWhitespace(NewLine);
            }
            writer.WriteWhitespace(parentIndent);
            writer.WriteEndArray();
            return;
        }

        writer.WriteStartArray();
        if (array.Count == 0)
        {
            writer.WriteEndArray();
            return;
        }

        writer.WriteWhitespace(NewLine);
        var childIndentNonPerLine = new string(' ', depth * 2);
        var parentIndentNonPerLine = new string(' ', (depth - 1) * 2);
        for (var i = 0; i < array.Count; i++)
        {
            writer.WriteWhitespace(childIndentNonPerLine);
            WriteTokenStream(array[i], writer, serializer, depth + 1, rootArrayPerLine, isRoot: false, inlineObjects: false);
            if (i < array.Count - 1)
                writer.WriteRaw(",");
            writer.WriteWhitespace(NewLine);
        }

        writer.WriteWhitespace(parentIndentNonPerLine);
        writer.WriteEndArray();
    }

    private async Task UpdateSheetComboAsync()
    {
        var path = _excelPathBox.Text?.Trim();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            ResetComboItems(_sheetNameBox);
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
            ResetComboItems(_sheetNameBox);
            _sheetNameBox.ItemsSource = items;
            _sheetNameBox.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            Log($"获取 Sheet 列表失败：{ex.Message}");
        }
    }

    private void RenderHighlightedJson(string json, bool rootArrayPerLine)
    {
        if (_previewBlock is null) return;
        var block = _previewBlock;
        block.Inlines!.Clear();
        block.Text = string.Empty;
        block.TextAlignment = TextAlignment.Left;
        block.TextWrapping = TextWrapping.NoWrap;
        block.HorizontalAlignment = HorizontalAlignment.Stretch;
        block.VerticalAlignment = VerticalAlignment.Stretch;
        block.Foreground = PreviewTextBrush;
        block.Opacity = 1.0;
        if (_previewWatermark is not null)
            _previewWatermark.IsVisible = false;
        try
        {
            var token = JToken.Parse(json);
            var isRootArray = token is JArray;
            var effectiveRootArrayPerLine = rootArrayPerLine && isRootArray;
            var copyText = BuildFormattedJson(json, effectiveRootArrayPerLine, GetComboValue(_dateFormatBox, "yyyy-MM-dd HH:mm:ss"));
            AppendToken(token, block.Inlines!, 0, effectiveRootArrayPerLine, isRootArray, inlineObjects: false);
            _lastPreviewText = copyText;
        }
        catch
        {
            block.Text = json;
            _lastPreviewText = json;
        }
    }

    private static readonly IBrush KeyBrush = new SolidColorBrush(Color.FromRgb(180, 200, 255));
    private static readonly IBrush StringBrush = new SolidColorBrush(Color.FromRgb(190, 225, 180));
    private static readonly IBrush NumberBrush = new SolidColorBrush(Color.FromRgb(255, 210, 140));
    private static readonly IBrush BooleanBrush = new SolidColorBrush(Color.FromRgb(255, 170, 170));
    private static readonly IBrush NullBrush = new SolidColorBrush(Color.FromRgb(150, 155, 165));
    private static readonly IBrush PunctuationBrush = new SolidColorBrush(Color.FromRgb(120, 130, 150));
    private static readonly IBrush PreviewTextBrush = new SolidColorBrush(Color.FromRgb(233, 237, 244));
    private static readonly IBrush WatermarkBrush = new SolidColorBrush(Color.FromRgb(140, 145, 155)) { Opacity = 0.65 };

    private void AppendToken(JToken token, InlineCollection inlines, int indent, bool rootArrayPerLine, bool isRootArray, bool inlineObjects)
    {
        if (inlines is null) return;

        switch (token.Type)
        {
            case JTokenType.Object:
                AppendObject((JObject)token, inlines, indent, rootArrayPerLine, isRootArray, inlineObjects);
                break;
            case JTokenType.Array:
                AppendArray((JArray)token, inlines, indent, rootArrayPerLine, isRootArray);
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

    private void AppendObject(JObject obj, InlineCollection inlines, int indent, bool rootArrayPerLine, bool isRootArray, bool inlineObjects)
    {
        if (obj.Count == 0)
        {
            inlines.Add(new Run("{") { Foreground = PunctuationBrush });
            inlines.Add(new Run("}") { Foreground = PunctuationBrush });
            return;
        }

        var properties = obj.Properties().ToList();
        inlines.Add(new Run("{") { Foreground = PunctuationBrush });
        if (inlineObjects)
        {
            for (var i = 0; i < properties.Count; i++)
            {
                var prop = properties[i];
                inlines.Add(new Run($"\"{prop.Name}\"") { Foreground = KeyBrush });
                inlines.Add(new Run(": ") { Foreground = PunctuationBrush });
                AppendToken(prop.Value, inlines, indent + 1, rootArrayPerLine, isRootArray: false, inlineObjects: true);
                if (i < properties.Count - 1)
                    inlines.Add(new Run(", ") { Foreground = PunctuationBrush });
            }
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
            AppendToken(prop.Value, inlines, indent + 1, rootArrayPerLine, isRootArray: false, inlineObjects: false);
            if (i < properties.Count - 1)
                inlines.Add(new Run(",") { Foreground = PunctuationBrush });
            inlines.Add(new Run("\n"));
        }
        inlines.Add(new Run(new string(' ', indent * 2)));
        inlines.Add(new Run("}") { Foreground = PunctuationBrush });
    }

    private void AppendArray(JArray array, InlineCollection inlines, int indent, bool rootArrayPerLine, bool isRootArray)
    {
        inlines.Add(new Run("[") { Foreground = PunctuationBrush });
        if (array.Count == 0)
        {
            inlines.Add(new Run("]") { Foreground = PunctuationBrush });
            return;
        }

        var perLine = rootArrayPerLine && isRootArray;
        inlines.Add(new Run("\n"));
        for (var i = 0; i < array.Count; i++)
        {
            inlines.Add(new Run(new string(' ', (indent + 1) * 2)));
            if (perLine)
            {
                var temp = new TextBlock();
                AppendToken(array[i], temp.Inlines!, indent + 1, rootArrayPerLine: false, isRootArray: false, inlineObjects: true);
                foreach (var piece in temp.Inlines!)
                    inlines.Add(piece);
            }
            else
            {
                AppendToken(array[i], inlines, indent + 1, rootArrayPerLine: false, isRootArray: false, inlineObjects: false);
            }

            if (i < array.Count - 1)
                inlines.Add(new Run(",") { Foreground = PunctuationBrush });
            inlines.Add(new Run("\n"));
        }
        inlines.Add(new Run(new string(' ', indent * 2)));
        inlines.Add(new Run("]") { Foreground = PunctuationBrush });
    }
}
