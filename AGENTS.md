# Repository Guidelines

## 项目结构与模块
- 解决方案：`Excel2Json.Modern.sln` 统一管理各项目。
- 核心库：`Excel2Json.Core/` 提供转换逻辑（`ExcelJsonConverter`、`ConversionOptions`），与界面无关。
- 桌面端：`Excel2Json.App/` 基于 Avalonia `.axaml`，引用核心库，承载主窗口与视图模型。
- 构建产物：`bin/`、`obj/` 为生成目录，请勿提交。

## 构建、测试与开发命令
- 首次恢复依赖：  
  ```powershell
  dotnet restore Excel2Json.Modern.sln
  ```
- 调试构建：  
  ```powershell
  dotnet build Excel2Json.Modern.sln
  ```
- 运行桌面应用：  
  ```powershell
  dotnet run --project Excel2Json.App
  ```
- 发布（Release）：  
  ```powershell
  dotnet publish Excel2Json.App -c Release -r win-x64 --self-contained false
  ```
  若需其他平台，请调整 `-r`。

## 代码风格与命名
- 语言：C#，启用可空引用类型与隐式 using。
- 缩进：4 空格；花括号换行，遵循 .NET 常规风格。
- 命名：类型/方法用 PascalCase，本地变量/字段用 camelCase；结尾明确标注 DTO/Options/Result（如 `ConversionOptions`、`ConversionResult`）。
- UI：Avalonia XAML 与对应 `.cs` 同目录，控件命名需便于绑定。

## 测试规范
- 当前无自动化测试，推荐后续使用 xUnit；放置于 `tests/` 或 `<Project>.Tests/`，命名空间与被测代码一致。
- 测试命名强调行为，如 `ConvertAsync_WritesExpectedRows`。
- 添加测试后，用 `dotnet test` 运行。

## 提交与合并请求
- 提交信息：简短祈使句（如 `Add sheet filter`，`Fix encoding fallback`），按功能分组，避免格式化与功能混杂。
- 合并请求：提供变更摘要、验证步骤（命令或 UI 截图）、关联任务编号，并说明新依赖或配置改动。

## 架构概览与安全
- 核心服务用 `ExcelDataReader` 读取 Excel，`Newtonsoft.Json` 输出 JSON，按指定编码与表头行数写入文件。
- UI 是覆盖核心库的轻量 Avalonia 壳，业务逻辑保持在 `Excel2Json.Core` 以便 CLI/服务复用。
- 处理文件时优先只读，缺失或占用应明确提示错误。
