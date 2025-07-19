import re
import os
import sys
from datetime import datetime

def update_version():
    # 读取当前版本号
    version_pattern = re.compile(r"VERSION\s*=\s*['\"]([0-9]+)\.([0-9]+)\.([0-9]+)['\"]")
    
    mc_recon_path = 'MC_Recon_UI.py'
    version_file_path = 'file_version_info.txt'
    
    # 读取MC_Recon_UI.py文件内容
    with open(mc_recon_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 查找版本号
    match = version_pattern.search(content)
    if not match:
        print("无法在MC_Recon_UI.py中找到版本号")
        return False
    
    # 解析版本号
    major, minor, patch = map(int, match.groups())
    
    # 增加修订版本号
    patch += 1
    new_version = f"{major}.{minor}.{patch}"
    
    # 更新MC_Recon_UI.py中的版本号
    new_content = version_pattern.sub(f"VERSION = '{major}.{minor}.{patch}'", content)
    with open(mc_recon_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    # 更新file_version_info.txt中的版本号
    if os.path.exists(version_file_path):
        with open(version_file_path, 'r', encoding='utf-8') as f:
            version_content = f.read()
        
        # 更新filevers和prodvers
        filevers_pattern = re.compile(r"filevers=\(([0-9]+),\s*([0-9]+),\s*([0-9]+),\s*([0-9]+)\)")
        version_content = filevers_pattern.sub(f"filevers=({major}, {minor}, {patch}, 0)", version_content)
        
        prodvers_pattern = re.compile(r"prodvers=\(([0-9]+),\s*([0-9]+),\s*([0-9]+),\s*([0-9]+)\)")
        version_content = prodvers_pattern.sub(f"prodvers=({major}, {minor}, {patch}, 0)", version_content)
        
        # 更新FileVersion和ProductVersion
        file_version_pattern = re.compile(r"StringStruct\(u'FileVersion',\s*u'[0-9\.]+'\)")
        version_content = file_version_pattern.sub(f"StringStruct(u'FileVersion', u'{new_version}')", version_content)
        
        product_version_pattern = re.compile(r"StringStruct\(u'ProductVersion',\s*u'[0-9\.]+'\)")
        version_content = product_version_pattern.sub(f"StringStruct(u'ProductVersion', u'{new_version}')", version_content)
        
        with open(version_file_path, 'w', encoding='utf-8') as f:
            f.write(version_content)
    
    print(f"版本号已更新为: {new_version}")
    return True

if __name__ == "__main__":
    update_version()