# PCR结果分析软件

一个用于分析PCR扩增结果的Windows桌面应用程序，支持解析不同厂商的Excel数据格式。

## 功能特性

- ✅ 支持多种厂商的Excel数据格式
- ✅ 自动识别数据格式并解析
- ✅ 可视化展示PCR扩增曲线
- ✅ 显示实验信息和Ct值
- ✅ 导出分析结果

## 系统要求

- Windows 10/11
- Python 3.8 或更高版本

## 安装步骤

### 1. 安装Python依赖

```bash
pip install -r requirements.txt
```

### 2. 运行程序

```bash
python pcr_analyzer.py
```

## 使用方法

1. **打开文件**: 点击"打开Excel文件"按钮，选择PCR结果Excel文件
2. **查看数据**: 在"实验数据"标签页查看解析的实验信息和数据表格
3. **查看曲线**: 在"扩增曲线"标签页查看PCR扩增曲线可视化
4. **导出结果**: 点击"导出结果"按钮保存分析结果

## 支持的格式

程序支持以下格式的Excel文件：

- **厂商A格式**: 包含"实验数据"和"扩增曲线"工作表的格式
- **标准格式**: 包含通道信息（HEX, CY5, ROX等）的标准格式
- **自定义格式**: 可通过扩展解析器支持更多格式

## 项目结构

```
qpcr/
├── pcr_analyzer.py      # 主程序
├── excel_parser.py      # Excel解析器
├── data_visualizer.py   # 数据可视化模块
├── requirements.txt     # 依赖包列表
├── README.md           # 说明文档
└── build_exe.py        # 打包脚本
```

## 扩展支持新格式

要支持新的厂商格式，可以在`excel_parser.py`中添加新的解析器类：

```python
class NewVendorParser(BaseParser):
    def parse(self, file_path):
        # 实现解析逻辑
        pass
```

然后在`ExcelParser`类的`__init__`方法中注册：

```python
self.parsers['new_vendor'] = NewVendorParser()
```

## 打包为可执行文件

使用PyInstaller打包：

```bash
pip install pyinstaller
python build_exe.py
```

打包后的exe文件将在`dist`目录中。

## 许可证

MIT License

## 技术支持

如有问题或建议，请提交Issue。

