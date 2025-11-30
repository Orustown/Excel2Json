# Excel/CSV 转 JSON 工具

基于 Avalonia 的跨平台桌面应用，用于将 Excel/CSV（.xlsx/.xls/.csv）转换为 JSON。左侧配置输入与导出选项，右侧提供所见即所得的 JSON 预览；核心转换逻辑集中在 `Excel2Json.Core`，便于复用。

## 项目结构
- `Excel2Json.Core/`：纯转换逻辑，读取 Excel/CSV、构建 JSON、预览、列出 Sheet。
- `Excel2Json.App/`：Avalonia UI，包含窗口、控件绑定与彩色 JSON 预览。
- 解决方案：`Excel2Json.Modern.sln`。

## 构建与运行
```powershell
dotnet restore Excel2Json.Modern.sln
dotnet build Excel2Json.Modern.sln
dotnet run --project Excel2Json.App
```
发布示例（Win x64 非自包含）：
```powershell
dotnet publish Excel2Json.App -c Release -r win-x64 --self-contained false
```

## 主要功能
- 选择或拖拽 Excel/CSV，自动建议同名 JSON 输出路径。
- 格式设置：表头行数、导出编码（utf-8/utf-16/gbk 等）、日期格式、目标 Sheet 自动检测。
- 导出选项：
  - Key 小写：仅作用于键名。
  - 导出为数组：顶层输出数组而非对象。
  - 保留 Sheet 包装：即使单 Sheet 也保留一层。
  - 识别单元格 JSON：值可解析为对象/数组。
  - 值转字符串：仅作用于值，键名不变。
  - 数组按行输出：根数组每个元素单独一行，便于对照 Excel 行。
- JSON 预览：与导出格式完全一致，支持复制、无换行折行；水印提示无文件时的状态。
- 日志与状态提示：记录操作与错误；提供“查看”按钮打开已导出的 JSON。

## 使用步骤
1. 选择 Excel/CSV（可拖拽），确认或调整输出 JSON 路径。
2. 在“格式设置”“导出选项”中配置编码、日期格式、Sheet、前缀排除等。
3. 右侧实时查看预览，确认无误后点击“导出 JSON”，在日志区域查看结果；可直接复制预览或点击“查看”打开导出文件。

## 依赖
- .NET 8
- Avalonia 11
- ExcelDataReader / Newtonsoft.Json

## 贡献
欢迎提交 Issue/PR，提交前建议执行 `dotnet build Excel2Json.Modern.sln` 进行验证。
