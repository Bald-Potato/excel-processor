# Excel 处理工具

## 功能介绍

这是一个用于处理Excel文件中starttime列数据的工具，具有以下功能：

1. **基本功能**：提取Excel文件中包含"starttime"的列数据，并保存到新的Excel文件中
2. **判断数据合法性**：检查提取出的starttime列内每个单元格的数据是否是40的整数倍
3. **时间格式转换**：将提取出的starttime列内每个单元格的数据由毫秒单位转化为hh:mm:ss:ff的格式
4. **自定义后缀**：允许用户为提取的文件设置自定义后缀，最多支持10个自定义后缀
5. **自动跳过临时文件**：自动跳过带有"~$"前缀的Excel临时文件

## 安装方法

### 方法一：使用打包好的应用程序

1. 在`dist`目录下找到`Excel处理工具_单文件.app`
2. 将其复制到Applications文件夹或任意位置
3. 双击运行应用程序

### 方法二：使用独立可执行文件

1. 在`dist`目录下找到`Excel处理工具`可执行文件
2. 将其复制到任意位置
3. 双击或通过终端运行该文件

### 方法三：从源代码运行

1. 确保已安装Python 3.9或更高版本
2. 安装所需依赖：`pip install pandas numpy PyQt6`
3. 运行前端程序：`python frontend_gui.py`

## 使用方法

1. 启动应用程序
2. 点击"浏览"按钮选择要处理的文件夹
3. 根据需要勾选功能选项：
   - 判断数据合法性
   - 采用24小时制
4. 添加自定义后缀（可选）：
   - 点击"添加自定义后缀"按钮添加自定义后缀输入框（最多10个）
   - 在输入框中输入自定义后缀名称
   - 如需删除，点击"删除自定义后缀"按钮
5. 点击"开始处理"按钮开始处理文件
6. 处理结果将显示在日志区域

## 输出文件说明

- 处理后的文件将保存在原文件夹同级的"文件夹名_processed"目录中
- 文件命名规则：
  - 如果启用了自定义后缀，文件名格式为：`[前缀]原文件名_自定义后缀.xlsx`
  - 如果没有自定义后缀，文件名格式为：`[前缀]原文件名_已处理[序号].xlsx`
  - 当启用"判断数据合法性"功能且数据不合法时，文件名前缀为"FALSE_"

## 重新打包应用程序

如果您修改了源代码并希望重新打包应用程序，可以运行：

```bash
python build_app.py
```

这将生成两个版本的应用程序：
1. `Excel处理工具_单文件.app` - 单文件版本的Mac应用程序
2. `Excel处理工具` - 独立可执行文件

## 注意事项

- 处理大量文件时可能需要等待一段时间
- 应用程序会自动跳过带有"~$"前缀的Excel临时文件
- 自定义后缀按照添加顺序与捕获到的starttime列一一对应
- 当捕获到的列数大于自定义后缀数量时，多余的列将使用默认的"已处理+序号"后缀