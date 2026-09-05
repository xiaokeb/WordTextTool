# Word 文本处理工具 - Code Wiki

##  项目概述

### 1.1 项目定位

**Word 文本处理工具**是一款功能完善的文档批量处理工具，支持三大核心功能：
1. **文本替换** - 批量替换 Word 文档中的文本（支持正文、表格、页眉页脚），保持原文档格式不变
2. **记录编号** - 向 Excel 文件写入编号信息，并支持导出为 PDF 格式
3. **后缀修改** - 批量删除文件名中的"-修改版"后缀

### 1.2 核心价值

| 特性 | 描述 |
|------|------|
| 批量处理 | 支持同时处理多个 Word/Excel 文档 |
| 格式保持 | 替换文本时保持原有格式不变 |
| 智能替换 | 基于类别（项目编号、公司名称等）进行批量替换 |
| 页眉页脚 | 支持替换页眉、页脚中的内容 |
| PDF导出 | 支持将 Excel 文件按指定页面设置导出为 PDF |
| 后缀管理 | 批量清理文件名中的冗余后缀 |
| 独立运行 | 打包为单个 exe 文件，无需 Python 环境 |

---

## 🏗️ 架构设计

### 2.1 整体架构

```
┌──────────────────────────────────────────────────────────────────┐
│                       Word 文本处理工具                           │
├──────────────────────────────────────────────────────────────────┤
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌──────┐│
│  │WordProcessor │  │ExcelProcessor│  │ExcelToPdfProc│  │Suffix││
│  │ (文档处理)   │  │ (编号处理)   │  │  (PDF导出)   │  │Proc  ││
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘  └─────┘│
│         │                 │                  │              │    │
│         ▼                 ▼                  ▼              ▼    │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌──────┐│
│  │  python-docx │  │   openpyxl   │  │   pywin32    │  │ os   ││
│  └──────────────┘  └──────────────  └──────────────┘  └──────┘│
│                                                                  │
│  ┌──────────────────────────────────────────────────────────────┤
│  │                    WordProcessorGUI (PyQt5)                   │
│  │  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐       │
│  │  │ 文本替换选项卡│  │ 记录编号选项卡│  │ 后缀修改选项卡│       │
│  │  └──────────────┘  └──────────────┘  └──────────────┘       │
│  └──────────────────────────────────────────────────────────────┘
└──────────────────────────────────────────────────────────────────┘
```

### 2.2 模块职责

| 模块 | 职责 | 关键类/函数 |
|------|------|-------------|
| **WordProcessor** | Word文档读取、文本替换、格式保持、页眉页脚处理 | `process_document()`, `replace_text_in_paragraph()`, `_process_headers()`, `_process_footers()` |
| **ExcelProcessor** | Excel文件读取、编号写入 | `write_number_to_excel()` |
| **ExcelToPdfProcessor** | Excel页面设置、PDF导出 | `export_to_pdf()` |
| **SuffixProcessor** | 文件名后缀批量删除 | `remove_suffix_from_files()` |
| **WordProcessorGUI** | 用户界面、交互逻辑、线程管理 | `_create_text_replace_tab()`, `_create_number_record_tab()`, `_create_suffix_modify_tab()` |
| **ExportPdfThread** | PDF导出后台线程，避免阻塞UI和COM线程问题 | `run()`, `export_complete`信号 |
| **ProcessingThread** | Word文档批量处理后台线程 | `run()`, `processing_complete`信号 |
| **WriteNumberThread** | Excel编号写入后台线程 | `run()`, `write_complete`信号 |
| **SuffixModifyThread** | 文件后缀修改后台线程 | `run()`, `modify_complete`信号 |

### 2.3 数据流

```
文件选择 → 验证过滤 → 显示列表 → 设置规则 → 多线程处理 → 保存输出
     ↓              ↓            ↓           ↓              ↓
 浏览文件      大小检查      Treeview     表格编辑       进度更新
 浏览文件夹    权限验证                    重置默认       结果提示
```

---

## 📦 核心模块详解

### 3.1 WordProcessor 类

**功能定位**：负责 Word 文档的读取、文本替换和格式保持

**文件位置**：`src/main.py`

#### 3.1.1 类属性

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `MAX_FILE_SIZE` | int | 50MB | 最大文件大小限制 |

#### 3.1.2 核心方法

##### `get_word_files(path)`

获取指定路径下的所有 Word 文件

| 参数 | 类型 | 说明 |
|------|------|------|
| `path` | str/Path/list | 文件路径、文件夹路径或文件列表 |

**返回值**：`list` - Word 文件路径列表

**处理流程**：
1. 验证路径存在性
2. 解析路径类型（文件/文件夹/列表）
3. 过滤 `.doc` 和 `.docx` 文件
4. 验证文件大小和访问权限

##### `replace_text_in_paragraph(paragraph, replacements)`

在段落中替换文本，保持格式不变

| 参数 | 类型 | 说明 |
|------|------|------|
| `paragraph` | Paragraph | docx 段落对象 |
| `replacements` | dict | 替换规则字典 {原词: 新词} |

**处理流程**：
1. 按原词长度降序排序（长词优先），避免部分匹配问题
2. 查找匹配文本
3. 调用 `_replace_all_occurrences()` 执行替换

##### `_replace_all_occurrences(paragraph, old_text, new_text)`

替换段落中所有匹配的文本，处理跨 Run 的文本替换，确保完整替换所有匹配项。

##### `_replace_single_occurrence(paragraph, old_text, new_text)`

**核心算法**：替换段落中单个匹配项

**处理逻辑**：
1. 合并所有 Run 的文本
2. 查找目标文本位置
3. 定位文本跨越的 Run 范围
4. 在第一个 Run 中执行替换，删除后续 Run 中的相关部分

##### `process_document(doc_path, replacements, output_path, callback=None)`

处理单个 Word 文档

| 参数 | 类型 | 说明 |
|------|------|------|
| `doc_path` | str | 源文档路径 |
| `replacements` | dict | 替换规则 |
| `output_path` | str | 输出路径 |
| `callback` | function | 完成回调 |

**处理范围**：正文段落、表格（含嵌套表格）、页眉、页脚

**返回值**：`bool` - 处理成功返回 True

##### `_process_tables(tables, replacements)`

处理文档中的所有表格，支持嵌套表格，递归处理每个单元格中的段落。

##### `_process_headers(sections, replacements)`

处理文档所有节的页眉，包括页眉中的段落和表格。

##### `_process_footers(sections, replacements)`

处理文档所有节的页脚，包括页脚中的段落和表格。

---

### 3.2 ExcelProcessor 类

**功能定位**：负责将编号写入 Excel 文件

**文件位置**：`src/main.py`

#### 3.2.1 核心方法

##### `write_number_to_excel(excel_path, output_path, number_text, use_fixed_mode=False, fixed_number="", number_format="", callback=None)`

将编号写入 Excel 文件

| 参数 | 类型 | 说明 |
|------|------|------|
| `excel_path` | str | Excel 文件路径 |
| `output_path` | str | 输出文件路径 |
| `number_text` | str | 生成的编号文本 |
| `use_fixed_mode` | bool | 是否使用固定编号模式 |
| `fixed_number` | str | 固定编号内容 |
| `number_format` | str | 编号格式 |

**写入位置**：
- 自动模式：编号写入 E1:F1 合并单元格
- 固定模式：固定编号写入 E1:F1，格式编号写入 E2:F2

---

### 3.3 ExcelToPdfProcessor 类

**功能定位**：负责 Excel 文件的页面设置和 PDF 导出

**文件位置**：`src/main.py`

**支持多种导出方法（按优先级自动检测）**：

| 优先级 | 方法 | 依赖 | 效果 |
|--------|------|------|------|
| 1 | pywin32 (COM) | pywin32 + Microsoft Excel | 最佳（保留Excel原格式） |
| 2 | LibreOffice CLI | LibreOffice | 良好（跨平台） |
| 3 | reportlab | reportlab | 备选（纯Python） |
| - | 无可用 | - | 显示安装提示 |

#### 3.3.1 核心方法

##### `_detect_method()`

自动检测可用的PDF导出方法，结果保存在`self.method`中。

##### `export_to_pdf(excel_path, output_path, callback=None)`

将 Excel 文件导出为 PDF

| 参数 | 类型 | 说明 |
|------|------|------|
| `excel_path` | str | 输入 Excel 文件路径 |
| `output_path` | str | 输出 PDF 文件路径 |
| `callback` | function | 完成回调 |

**页面设置**（pywin32方式）：
- 缩放：将所有列缩放为一页（FitToPagesWide=1, FitToPagesTall=0）
- 页边距：上下 1.3cm，左右 0.9cm
- 选中所有工作表后统一导出

**返回值**：`tuple` - (成功: bool, 消息: str)

##### 私有方法

- `_export_with_pywin32()` - 使用COM接口导出（效果最佳）
- `_export_with_libreoffice()` - 使用LibreOffice命令行导出
- `_export_with_reportlab()` - 使用纯Python库导出（备选）

---

### 3.4 SuffixProcessor 类（新增）

**功能定位**：负责批量修改文件名，删除指定后缀

**文件位置**：`src/main.py`

#### 3.4.1 核心方法

##### `remove_suffix_from_files(files, suffix="-修改版", callback=None)`

批量删除文件名中所有的指定后缀（包括中间和末尾）

| 参数 | 类型 | 说明 |
|------|------|------|
| `files` | list | 文件路径列表 |
| `suffix` | str | 要删除的后缀，默认"-修改版" |
| `callback` | function | 进度回调 (当前索引, 总数, 文件名) |

**返回值**：`tuple` - (成功数量, 总数)

**处理逻辑**：
1. 使用 `replace()` 删除文件名中所有的后缀
2. 自动处理文件名冲突（添加序号后缀）
3. 只有名称实际发生变化时才重命名

---

### 3.5 WordProcessorGUI 类

**功能定位**：用户界面和交互逻辑

**文件位置**：`src/main.py`

#### 3.5.1 界面结构

```
主窗口 (QMainWindow)
├── 选项卡容器 (QTabWidget)
│   ├── 文本替换选项卡
│   │   ├── 文件选择区域（浏览文件/浏览文件夹）
│   │   ├── 已选文件列表区域（QTableWidget，支持双击删除）
│   │   ├── 替换规则设置区域（QTableWidget，双击编辑单元格）
│   │   ├── 保存设置区域
│   │   ├── 处理进度区域（QProgressBar）
│   │   └── 操作按钮区域（开始修改/清空列表）
│   ├── 记录编号选项卡
│   │   ├── 编号模式选择（自动/固定）
│   │   ├── 编号格式设置
│   │   ├── 固定编号输入
│   │   ├── Excel 文件选择
│   │   ├── 保存设置（复用路径用于编号写入和PDF导出）
│   │   ├── 处理进度区域
│   │   └── 操作按钮区域（写入编号/重置编号/导出PDF/一键转换）
│   └── 后缀修改选项卡
│       ├── 文件选择区域（浏览文件/浏览文件夹）
│       ├── 已选文件列表区域（QTableWidget，显示原文件名和修改后）
│       ├── 操作按钮区域（添加文件/删除文件）
│       ├── 处理进度区域
│       └── 执行按钮区域（开始修改/清空列表）
```

#### 3.5.2 核心方法

##### `_create_text_replace_tab()`

创建文本替换选项卡的所有组件

##### `_start_processing()`

启动批量处理（后台线程）

**处理流程**：
1. 验证文件列表和替换规则
2. 禁用开始按钮防止重复点击
3. 创建后台线程执行处理
4. 实时更新进度条和状态

##### `_get_all_replacements()`

获取所有替换规则并转换为字典

##### `_generate_number()`

生成编号，序号自动递增（超过99后变为3位数）

##### `_create_number_record_tab()`

创建记录编号选项卡的所有组件

##### `_start_write_number()`

启动编号写入（后台线程）

##### `_start_export_pdf()`

启动 PDF 导出，复用保存设置中的路径

##### `_start_one_click_convert()`

一键转换：先写入编号，再导出 PDF

##### `_create_suffix_modify_tab()`

创建后缀修改选项卡的所有组件

##### `_start_suffix_processing()`

启动后缀修改处理

#### 3.5.3 UI 优化说明

**表格样式统一**：
- 所有表格表头高度统一为 22px
- 文本替换文件列表行高 20px，单元格内边距 1px 6px
- 替换规则设置行高 24px，单元格内边距 2px 6px
- 后缀修改文件列表行高 20px，单元格内边距 1px 6px

**编辑单元格优化**：
- 替换规则表格编辑时无边框间隔，内容完整显示
- 固定行高，默认不显示滚动条
- 重置按钮与表格不重叠

---

## 🧵 多线程模型

### 4.1 线程架构

为避免UI阻塞和COM线程问题，所有耗时操作均放在独立的后台线程中执行：

| 线程类 | 用途 | 避免的问题 |
|--------|------|------------|
| **ProcessingThread** | Word文档批量替换 | UI卡顿、长时间无响应 |
| **WriteNumberThread** | Excel编号写入 | UI卡顿、openpyxl阻塞 |
| **ExportPdfThread** | Excel转PDF | UI卡顿、COM线程死锁 |
| **SuffixModifyThread** | 文件名后缀修改 | UI卡顿、文件系统操作阻塞 |

### 4.2 PDF导出线程特别说明

**问题背景**：`pywin32`通过COM接口调用Excel，需要在线程内单独初始化COM apartment。如果在主GUI线程中执行COM操作，会导致：
- UI完全冻结
- "RPC server unavailable"错误
- 程序闪退

**解决方案**：`ExportPdfThread`继承自`QThread`，在`run()`方法中：
1. 调用`pythoncom.CoInitialize()`初始化COM
2. 创建Excel COM对象并执行导出
3. 无论成功失败都调用`pythoncom.CoUninitialize()`清理
4. 通过`export_complete`信号返回结果

### 4.3 启动性能优化

`pdf_processor`采用懒加载模式（初始化为None），避免启动时加载`pywin32`：
```python
self.pdf_processor = None  # 延迟初始化，避免启动时加载pywin32
```

首次使用PDF功能时才创建实例，显著缩短程序启动时间。

---

## 🧪 单元测试

### 5.1 测试文件

测试文件位于 `tests/test_main.py`，包含完整的单元测试套件。

### 5.2 测试范围

| 测试类 | 测试内容 | 测试数量 |
|--------|----------|----------|
| **TestWordProcessor** | 文件验证、文本替换、跨Run替换、表格处理、异常处理 | 10 |
| **TestExcelProcessor** | 自动模式、固定模式、错误处理 | 3 |
| **TestSuffixProcessor** | 后缀删除、自定义后缀、冲突处理、异常处理 | 6 |
| **TestExcelToPdfProcessor** | pywin32可用性、错误处理 | 2 |
| **TestUtilityFunctions** | 路径操作、文件名冲突 | 2 |
| **TestIntegration** | 端到端工作流、表格+段落组合替换 | 2 |
| **总计** | | **25+** |

### 5.3 运行测试

```bash
cd e:\SourceCode\Projects\WordModification
python tests/test_main.py
```

### 5.4 测试结果

最新测试结果：27个测试全部通过（2个PDF相关测试在缺少pywin32时跳过）。

```
Ran 27 tests in 1.566s
OK (skipped=2)
```

### 5.5 关键测试场景

- **跨Run文本替换**：验证大段文字和表格内的文本被完整替换
- **多规则同时替换**：验证多个替换规则同时生效
- **空规则处理**：验证空值规则被安全跳过
- **文件不存在**：验证不存在的文件被正确处理
- **后缀多位置删除**：验证文件名中多个"-修改版"都被删除
- **集成工作流**：验证 Excel→写入编号→PDF导出的完整流程

---

## 🔧 技术栈

### 4.1 依赖列表

| 依赖 | 版本 | 用途 |
|------|------|------|
| Python | 3.7+ | 编程语言 |
| PyQt5 | 5.x | GUI 框架 |
| python-docx | latest | Word 文档处理 |
| openpyxl | latest | Excel 文档处理 |
| pywin32 | latest | Excel COM 接口（PDF 导出） |
| PyInstaller | latest | 打包工具 |

### 4.2 安装命令

```bash
# 安装核心依赖
pip install python-docx openpyxl pywin32

# 安装 PyQt5
pip install PyQt5

# 安装打包工具
pip install pyinstaller
```

---

## 🚀 运行方式

### 5.1 开发模式

```bash
cd src
python main.py
```

### 5.2 打包发布

```bash
# 使用打包脚本
build.bat

# 或手动执行
pyinstaller --clean build.spec
```

打包完成后，可执行文件位于 `dist/Word 文本处理工具.exe`

---

## 📁 项目结构

```
WordModification/
├── .idea/                    # IDE 配置（可忽略）
│   ├── inspectionProfiles/
│   ├── .gitignore
│   ├── misc.xml
│   └── modules.xml
├── doc/                      # 文档目录
│   ├── dist-README.md        # 分发说明
│   ├── 需求文档.md           # 需求规格说明书
│   └── CODE_WIKI.md          # 代码Wiki文档
├── src/                      # 源代码目录
│   ── main.py               # 主程序入口
── .gitignore                # Git 忽略配置
├── README.md                 # 项目说明
── requirements.txt          # 依赖配置
```

---

## 🎯 使用流程

### 6.1 文本替换流程

1. **选择文件**：点击"浏览文件"或"浏览文件夹"
2. **管理列表**：通过"+"/"-"按钮或双击管理文件列表
3. **设置规则**：在表格中编辑替换规则（双击单元格编辑）
4. **设置路径**：选择保存路径（可选）
5. **开始处理**：点击"开始修改"按钮

### 6.2 记录编号流程

1. **选择模式**：自动编号模式或固定编号模式
2. **设置格式**：输入编号格式（使用 `yy` 作为序号占位符）
3. **选择文件**：选择要写入的 Excel 文件
4. **设置路径**：选择保存路径（可选，同时用于编号写入和 PDF 导出）
5. **执行操作**：
   - **写入编号**：仅执行编号写入
   - **导出PDF**：仅执行格式转换（选中所有工作表，所有列缩放为一页，上下页边距1.3cm，左右页边距0.9cm）
   - **一键转换**：先写入编号再导出 PDF

### 6.3 后缀修改流程

1. **选择文件**：点击"浏览文件"或"浏览文件夹"
2. **管理列表**：通过"+"/"-"按钮或双击管理文件列表
3. **预览效果**：列表中显示原文件名和修改后的文件名
4. **执行修改**：点击"开始修改"按钮

---

## ⚠️ 异常处理

### 7.1 文件相关错误

| 错误类型 | 处理方式 | 用户提示 |
|----------|----------|----------|
| 文件不存在 | 返回 False | 显示错误提示 |
| 文件过大 | 自动过滤 | 显示警告 |
| 无访问权限 | 返回 False | 显示错误提示 |
| 格式不支持 | 跳过处理 | 静默跳过 |
| Excel 被占用 | 返回 False | 提示检查文件是否被占用 |

### 7.2 路径相关错误

| 错误类型 | 处理方式 | 用户提示 |
|----------|----------|----------|
| 路径不存在 | 返回 None | 显示错误提示 |
| 无写入权限 | 测试验证 | 显示错误提示 |
| 路径非法 | 拒绝操作 | 显示错误提示 |

---

## 🔒 安全特性

### 8.1 文件安全检查

- **大小限制**：最大 50MB，自动过滤超大文件
- **权限验证**：检查文件读取权限和文件夹写入权限
- **路径规范化**：使用 `resolve()` 获取绝对路径，防止路径遍历攻击

### 8.2 输入验证

- **空值处理**：查找内容为空时规则不生效
- **异常捕获**：全面的 try-except 包裹

---

## 📊 性能指标

| 指标 | 参考值 |
|------|--------|
| 单文档处理 (100KB) | ~0.5 秒 |
| 批量处理 (10个×1MB) | ~8 秒 |
| 内存占用 | ~50MB |

---

##  版本历史

| 版本 | 日期 | 变更说明 |
|------|------|----------|
| v1.0.0 | 2026-04-03 | 初始版本，Tkinter 实现 |
| v2.0.0 | 2026-05-28 | 重构为 PyQt5，优化界面 |
| v3.0.0 | 2026-06-01 | 新增后缀修改选项卡，优化文本替换逻辑（跨Run替换、页眉页脚支持） |
| v4.0.0 | 2026-06-10 | 新增 ExcelToPdfProcessor，记录编号选项卡新增导出PDF和一键转换功能 |
| v5.0.0 | 2026-06-12 | UI 全面优化：表格行高统一、编辑样式优化、按钮布局调整 |
| v6.0.0 | 2026-06-15 | 记录编号选项卡行为逻辑重构：保存路径复用、新增一键转换按钮、PDF导出独立功能 |
| v6.1.0 | 2026-06-16 | **关键Bug修复**：PDF导出/一键转换从主线程移至ExportPdfThread后台线程，修复COM线程死锁和程序闪退问题；pdf_processor懒加载优化启动速度；新增完整单元测试套件（27个测试） |
| v6.2.0 | 2026-06-17 | **PDF导出多方法支持**：ExcelToPdfProcessor自动检测pywin32/LibreOffice/reportlab三种导出方法，按优先级回退；解决用户环境无pywin32时的"缺少pywin32库"问题；错误信息更友好；新增requirements.txt依赖说明 |
| v6.3.0 | 2026-08-04 | **PDF导出稳定性优化**：_detect_method改为类级别缓存避免重复检测；移除启动时临时处理器创建，改为首次使用时检测；ExportPdfThread新增method_detected信号，检测完成后更新按钮tooltip；修复test_no_method_message测试；30个单元测试全部通过 |
| v6.4.0 | 2026-08-04 | **PDF导出方案精简**：移除LibreOffice和reportlab方案，仅保留pywin32方案；新增WPS表格支持（Ket.Application），自动检测Excel/WPS；优化页面设置：FitToPagesTall改为1（所有行缩放为一页）；新增自动调整行高（AutoFit），确保长文本完整显示；29个单元测试全部通过 |

---

**文档版本**: v2.1  
**最后更新**: 2026-06-16  
**文档状态**: 已完成
