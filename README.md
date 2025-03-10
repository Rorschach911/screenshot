```markdown:d:\Python\Github\screenshot\README.md
# 网页截图自动化工具

一个基于 Python 的自动化工具，用于批量获取网页截图并生成 PPT 报告。

## 功能特点

- 批量处理网页截图
- 自动生成 PPT 报告
- 友好的图形用户界面
- 实时进度显示
- 支持大批量链接处理
- 自动调整截图大小和布局

## 系统要求

- Windows 操作系统
- Python 3.7+
- Google Chrome 浏览器
- Microsoft PowerPoint

## 依赖包

```bash
pip install selenium
pip install webdriver-manager
pip install pandas
pip install python-pptx
pip install pywin32
pip install openpyxl
```

## 使用说明

### Excel 文件格式要求

Excel 文件必须包含以下列：
- 媒体名称
- 发布时间
- 链接

### 操作步骤

1. 启动程序
2. 选择输入的 Excel 文件（包含网页链接）
3. 选择输出的 PPT 文件
4. 输入页面标题
5. 点击"执行"开始处理
6. 等待处理完成

### 界面说明

- 导入 Excel：选择包含链接信息的 Excel 文件
- 导出 PPT：选择 PPT 保存位置
- 页面标题：设置 PPT 中显示的标题
- 执行进度：显示当前处理进度
- 链接列表：显示所有待处理的链接及其状态

## 项目结构

- `main.py`: 程序入口点
- `ui.py`: 图形界面实现
- `screenshot.py`: 网页截图核心功能
- `excel_handler.py`: Excel 文件处理
- `ppt_handler.py`: PPT 文件生成和处理

## 注意事项

1. 确保 Excel 文件格式正确，包含所有必需列
2. 处理过程中请勿关闭 Chrome 浏览器
3. 确保有足够的磁盘空间存储临时文件
4. 建议定期保存 PPT 文件
5. 处理大量链接时可能需要较长时间

## 错误处理

程序会自动处理以下情况：
- Excel 文件格式错误
- 网页加载失败
- PPT 保存失败
- 临时文件清理

## 技术实现

- 使用 Selenium 进行网页自动化
- 使用 python-pptx 处理 PPT 文件
- 使用 pandas 处理 Excel 数据
- 使用 tkinter 构建图形界面
- 使用 win32com 进行 PPT 操作

## 更新日志

### v1.0.0
- 实现基础的网页截图功能
- 支持 Excel 批量导入
- 添加进度显示功能
- 实现自动 PPT 生成
```

这个 README.md 文件涵盖了项目的主要功能、使用方法、依赖要求等关键信息。如果你觉得还需要补充其他内容，请告诉我。