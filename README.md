# Hollysys VS Code 扩展

## 简介
Hollysys 是一个适用于 Visual Studio Code 的扩展，旨在为 Structured Text (ST) 文件提供特定的功能支持，包括文件夹创建、Excel 文件生成与更新、POU 和 ST 文件的处理等。

## 主要功能
- **新建 Hollysys 工程**：创建必要的文件夹结构，并生成初始的 ST 框架 Excel 文件。
- **更新 Excel 文件**：根据现有文件生成用于 POU 替换、典型回路和 ST 替换的 Excel 文件。
- **更新 ST 变量表**：检查当前活动文件是否为 ST 文件，并显示相应的提示信息。
- **替换 POU**：根据 Excel 文件中的映射关系替换 POU 文件中的点名。
- **生成典型回路**：从 Excel 文件中读取数据并生成对应的 XML 回路文件。
- **生成 ST 顺控**：暂未实现具体功能。
- **替换 ST**：暂未实现具体功能。
- **更新 POU 变量表**：暂未实现具体功能。

## 安装
1. 打开 Visual Studio Code。
2. 进入扩展视图（快捷键 `Ctrl+Shift+X` 或 `Cmd+Shift+X`）。
3. 在搜索栏中输入 `Hollysys`。
4. 点击安装按钮进行安装。

## 使用方法
### 新建 Hollysys 工程
1. 打开命令面板（快捷键 `Ctrl+Shift+P` 或 `Cmd+Shift+P`）。
2. 输入 `新建 hollysys` 并选择。
3. 扩展将自动创建所需的文件夹结构，并生成 `ST框架.xlsx` 文件。

### 更新 Excel 文件
1. 打开命令面板（快捷键 `Ctrl+Shift+P` 或 `Cmd+Shift+P`）。
2. 输入 `更新 excel` 并选择。
3. 扩展将根据现有的文件生成用于 POU 替换、典型回路和 ST 替换的 Excel 文件。

### 更新 ST 变量表
1. 打开一个 ST 文件。
2. 打开命令面板（快捷键 `Ctrl+Shift+P` 或 `Cmd+Shift+P`）。
3. 输入 `更新 ST 变量表` 并选择。
4. 如果当前文件是 ST 文件，则会显示更新成功的提示信息；否则会显示警告信息。

### 替换 POU
1. 打开命令面板（快捷键 `Ctrl+Shift+P` 或 `Cmd+Shift+P`）。
2. 输入 `替换 POU` 并选择。
3. 扩展将根据 `点名替换.xlsx` 文件中的映射关系替换 `POU替换输入` 文件夹中的 POU 文件中的点名，并生成新的文件到 `POU替换输出` 文件夹。

### 生成典型回路
1. 打开命令面板（快捷键 `Ctrl+Shift+P` 或 `Cmd+Shift+P`）。
2. 输入 `生成回路` 并选择。
3. 扩展将从 `典型回路.xlsx` 文件中读取数据并生成对应的 XML 回路文件到 `典型回路输出` 文件夹。

### 生成 ST 顺控
1. 打开命令面板（快捷键 `Ctrl+Shift+P` 或 `Cmd+Shift+P`）。
2. 输入 `生成 ST 顺控` 并选择。
3. 目前该功能尚未实现，后续版本将提供具体功能。

### 替换 ST
1. 打开命令面板（快捷键 `Ctrl+Shift+P` 或 `Cmd+Shift+P`）。
2. 输入 `替换 ST` 并选择。
3. 目前该功能尚未实现，后续版本将提供具体功能。

### 更新 POU 变量表
1. 打开命令面板（快捷键 `Ctrl+Shift+P` 或 `Cmd+Shift+P`）。
2. 输入 `更新 POU 变量表` 并选择。
3. 目前该功能尚未实现，后续版本将提供具体功能。

## 依赖项
- `fast-xml-parser`: 用于解析和生成 XML 文件。
- `xlsx`: 用于读取和写入 Excel 文件。

## 贡献者
- 作者：红烧肉。

## 许可证
本项目采用 MIT 许可证。详情请参见 [LICENSE](LICENSE) 文件。