# Excel自动调整行高列宽工具

## 功能介绍

这个工具用于自动调整Excel文件的行高和列宽，使其内容显示美观清晰。主要功能包括：

- 自动计算并调整最佳列宽和行高
- 支持中英文混合文本的正确显示
- 正确处理多行文本和换行符
- 考虑字体大小等因素进行调整
- 支持自定义调整参数
- 支持批量处理多个Excel文件

## 安装依赖

使用前请确保已安装以下Python库：

```bash
pip install openpyxl
```

## 使用方法

### 命令行使用

```bash
python xlsx_autofit_columns.py [文件路径] [选项]
```

#### 示例：

```bash
# 处理单个文件
python xlsx_autofit_columns.py 示例文件.xlsx

# 指定输出文件
python xlsx_autofit_columns.py 示例文件.xlsx -o 输出文件.xlsx

# 使用自定义参数
python xlsx_autofit_columns.py 示例文件.xlsx --width-factor 1.5 --height-factor 1.4

# 使用代码中DEFAULT_FILES列表中的文件
python xlsx_autofit_columns.py --use-defaults
```

### 在代码中设置默认文件

您可以直接在代码中的 `DEFAULT_FILES`列表中添加要处理的文件路径：

```python
DEFAULT_FILES = [
    '问题模板.xlsx',
    'sql模板.xlsx'
    # 添加更多文件路径
]
```

## 参数说明

| 参数                | 说明                                    | 默认值                    |
| ------------------- | --------------------------------------- | ------------------------- |
| `file`            | Excel文件路径（可选）                   | -                         |
| `-o, --output`    | 输出文件路径                            | 原文件名_beautifuler.xlsx |
| `--width-factor`  | 列宽调整系数                            | 1.3                       |
| `--height-factor` | 行高调整系数                            | 1.3                       |
| `--min-width`     | 最小列宽                                | 8                         |
| `--max-width`     | 最大列宽                                | 120                       |
| `--min-height`    | 最小行高                                | 20                        |
| `--max-height`    | 最大行高                                | 409                       |
| `--font-size`     | 默认字体大小                            | 12                        |
| `--use-defaults`  | 使用代码中DEFAULT_FILES列表中的文件路径 | -                         |

### 功能开关选项

| 参数                         | 说明                              | 默认值 |
| ---------------------------- | --------------------------------- | ------ |
| `--disable-size-limits`    | 禁用单元格尺寸限制                | 启用   |
| `--disable-font-autofit`   | 禁用字体大小自适应                | 禁用   |
| `--disable-cell-wrap`      | 禁用单元格自动换行                | 启用   |
| `--disable-cell-alignment` | 禁用单元格对齐方式                | 启用   |
| `--horizontal-alignment`   | 水平对齐方式：left, center, right | center |
| `--vertical-alignment`     | 垂直对齐方式：top, center, bottom | center |

## 配置选项说明

您可以通过修改代码中的 `DEFAULT_CONFIG`字典来更改默认配置：

```python
DEFAULT_CONFIG = {
    'width_factor': 1.3,      # 列宽调整系数 (增加以显示更充分)
    'height_factor': 1.3,     # 行高调整系数 (增加以显示更充分)
    'min_width': 8,           # 最小列宽
    'max_width': 120,         # 最大列宽 (增加以显示更充分)
    'min_height': 20,         # 最小行高
    'max_height': 409,        # 最大行高（Excel限制）
    'default_font_size': 12,  # 默认字体大小
    'enable_size_limits': True,    # 是否启用单元格尺寸限制
    'enable_font_autofit': False,   # 是否启用字体大小自适应
    'enable_cell_wrap': True,      # 是否启用单元格自动换行
    'enable_cell_alignment': True,  # 是否启用单元格对齐方式
    'horizontal_alignment': 'center',  # 水平对齐方式：'left', 'center', 'right'
    'vertical_alignment': 'center',    # 垂直对齐方式：'top', 'center', 'bottom'
}
```

## 常见问题解答

### 1. 为什么我的Excel文件无法处理？

可能的原因：

- 文件正在被其他程序（如Excel）使用
- 没有足够的文件访问权限
- 文件格式不正确或已损坏

解决方法：

- 关闭可能正在使用此文件的其他程序
- 确保您有足够的文件访问权限
- 尝试将文件复制到其他位置后再处理
- 使用另一个Excel文件测试

### 2. 为什么处理后的文件某些单元格显示不完整？

可能的原因：

- 单元格内容过长，超出了最大列宽限制
- 单元格内有特殊格式或对象

解决方法：

- 增加最大列宽限制：`--max-width 150`
- 增加列宽调整系数：`--width-factor 1.5`
- 启用单元格自动换行（默认已启用）

### 3. 为什么包含换行符的单元格没有正确显示？

本工具已经修复了处理换行符的问题。如果您仍然遇到问题，请确保：

- 使用最新版本的工具
- 单元格中的换行符是真正的换行符（\n），而不是文本中的"\n"字符
- 单元格已启用自动换行（默认已启用）

### 4. 如何批量处理多个文件？

方法一：在代码中设置默认文件列表

```python
DEFAULT_FILES = [
    '文件1.xlsx',
    '文件2.xlsx',
    '文件3.xlsx'
]
```

然后运行：`python xlsx_autofit_columns.py --use-defaults`

方法二：使用批处理脚本循环处理多个文件

## 更新日志

### v1.1.0

- 修复了处理换行符的问题，现在能正确计算包含换行符的单元格高度
- 添加了详细的README文档

### v1.0.0

- 初始版本发布

## 许可证

本工具采用MIT许可证。

## 作者

如有问题或建议，请联系作者qq1534563895。
