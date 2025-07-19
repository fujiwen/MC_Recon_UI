import os
import re
import subprocess
import sys
import zipfile
from datetime import datetime

# 设置控制台编码为UTF-8
os.environ["PYTHONIOENCODING"] = "utf-8"

# 打印标题
print("="*50)
print("      MC对账明细工具 - 自动打包脚本")
print("="*50)

# 更新版本号
print("[1/3] 正在更新版本号...")
result = subprocess.run([sys.executable, "update_version.py"], capture_output=True, text=True, encoding="utf-8")
print(result.stdout)

# 获取当前版本号
with open("MC_Recon_UI.py", "r", encoding="utf-8") as f:
    content = f.read()
    version_match = re.search(r"VERSION = '([\d\.]+)'", content)
    current_version = version_match.group(1) if version_match else "未知"

# 编译资源文件
print("[2/5] 正在编译资源文件...")
resource_result = subprocess.run(["pyrcc5", "-o", "resources.py", "resources.qrc"], capture_output=True, text=True, encoding="utf-8")
if resource_result.returncode != 0:
    print("资源文件编译失败！")
    print("错误信息:")
    print(resource_result.stderr)
    sys.exit(1)

# 打包应用程序
print(f"[3/5] 正在打包应用程序 v{current_version}...")
result = subprocess.run(["pyinstaller", "MC对账明细工具.spec", "--clean"], capture_output=True, text=True, encoding="utf-8")

# 检查打包结果
exe_path = os.path.join("dist", f"MC对账明细工具_v{current_version}.exe")
if os.path.exists(exe_path):
    print("[4/5] 打包完成！")
    
    # 获取文件信息
    file_time = datetime.fromtimestamp(os.path.getmtime(exe_path))
    file_size = os.path.getsize(exe_path) / (1024 * 1024)  # 转换为MB
    
    print("\n应用程序信息:")
    print(f"  文件名称: {exe_path}")
    print(f"  文件大小: {file_size:.2f} MB")
    print(f"  创建时间: {file_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  版本号: {current_version}")
    
    # 创建压缩包
    print("[5/5] 正在创建压缩包...")
    zip_filename = f"MC对账明细工具_v{current_version}.zip"
    zip_path = os.path.join("dist", zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # 添加新打包的exe文件
        zipf.write(exe_path, os.path.basename(exe_path))
    
    # 获取压缩包信息
    zip_size = os.path.getsize(zip_path) / (1024 * 1024)  # 转换为MB
    
    print("\n压缩包信息:")
    print(f"  文件名称: {zip_filename}")
    print(f"  文件大小: {zip_size:.2f} MB")
    
    # 打开输出目录
    print(f"\n打包文件位于: {os.path.abspath('dist')}")
else:
    print("[3/3] 打包失败！")
    print("错误信息:")
    print(result.stderr)

input("\n按回车键退出...")