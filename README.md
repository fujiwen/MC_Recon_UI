# MC对账明细工具

## 项目说明

这是一个用于处理收货记录和生成供应商对账明细的工具。该工具使用Python开发，并使用PyQt5构建图形用户界面。

## 运行环境要求

- Python 3.10或更高版本
- 依赖包：pandas, numpy, openpyxl, PyQt5

## 不使用虚拟环境运行

### 方法一：使用批处理文件

1. 双击运行`run_without_venv.bat`文件
2. 批处理文件会自动检查Python环境和依赖项，并启动程序

### 方法二：手动运行

1. 确保已安装Python 3.10或更高版本
2. 安装依赖项：
   ```
   pip install -r requirements.txt
   ```
3. 运行程序：
   ```
   python MC_Recon_UI.py
   ```

## 构建可执行文件

如果需要构建为独立的可执行文件，可以使用以下命令：

```
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --icon=favicon.ico --name="MC对账明细工具" MC_Recon_UI.py
```

构建完成后，可执行文件将位于`dist`目录中。

## 自动构建

本项目已配置GitHub Actions工作流，当代码推送到main分支时，会自动构建Windows可执行文件。构建结果可在GitHub Actions的构建工件中下载。