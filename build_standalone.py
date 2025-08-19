#!/usr/bin/env python3
"""
独立打包脚本
创建完全自包含的可执行文件，包含所有配置文件
"""

import os
import sys
import shutil
import subprocess
import json

def create_standalone_build():
    print("====================================")
    print("   Building SeaTable Excel Sync Tool")
    print("====================================")
    
    # 1. Install dependencies
    print("1. Installing dependencies...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], check=True)
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    
    # 2. Clean old files
    print("\n2. Cleaning previous build files...")
    for folder in ["dist", "build"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
    
    for file in os.listdir("."):
        if file.endswith(".spec"):
            os.remove(file)
    
    # 3. Collect config files
    config_files = []
    if os.path.exists("config"):
        config_files = [f"config/{f}" for f in os.listdir("config") if f.endswith(".json")]
    print(f"\n3. Found config files: {', '.join(config_files)}")
    
    # 4. 构建PyInstaller命令
    cmd = [
        "pyinstaller",
        "--onefile",
        "--console",
        "--name", "seatable-sync",
        "--noupx",  # 禁用UPX压缩，避免DLL加载问题
        "--clean",  # 清理缓存
        "--hidden-import", "seatable_api",
        "--hidden-import", "pandas",
        "--hidden-import", "numpy",
        "--hidden-import", "openpyxl",
        "--hidden-import", "dotenv",
        "--hidden-import", "json",
        "--hidden-import", "datetime"
    ]
    
    # Windows特定选项
    if sys.platform.startswith("win"):
        cmd.extend([
            "--collect-all", "seatable_api",  # 收集所有seatable_api依赖
            "--collect-all", "pandas",        # 收集所有pandas依赖
            "--collect-all", "openpyxl",      # 收集所有openpyxl依赖
            "--noconsole" if "--noconsole" in sys.argv else "--console"
        ])
    
    # 添加JSON配置文件
    json_files = [f for f in os.listdir(".") if f.endswith(".json")]
    for json_file in json_files:
        cmd.extend(["--add-data", f"{json_file}:."])
    
    # 添加.env示例文件
    if os.path.exists(".env.example"):
        cmd.extend(["--add-data", ".env.example:."])
    
    # 添加主文件
    cmd.append("main-name-pro.py")
    
    print(f"\n4. Executing build command...")
    print(f"Command: {' '.join(cmd)}")
    
    try:
        subprocess.run(cmd, check=True)
        print("\n[OK] Build successful!")
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] Build failed: {e}")
        return False
    
    # 5. Create deployment package
    print("\n5. Creating deployment package...")
    
    # Create deployment directory
    deploy_dir = "seatable-sync-deploy"
    if os.path.exists(deploy_dir):
        shutil.rmtree(deploy_dir)
    os.makedirs(deploy_dir)
    
    # Copy executable file
    exe_name = "seatable-sync.exe" if sys.platform.startswith("win") else "seatable-sync"
    src_exe = os.path.join("dist", exe_name)
    dst_exe = os.path.join(deploy_dir, exe_name)
    
    if os.path.exists(src_exe):
        shutil.copy2(src_exe, dst_exe)
        print(f"[OK] Copied executable: {exe_name}")
    else:
        print(f"[ERROR] Executable not found: {src_exe}")
        return False
    
    # Copy JSON config files
    json_files = [f for f in os.listdir(".") if f.endswith(".json")]
    for json_file in json_files:
        shutil.copy2(json_file, deploy_dir)
        print(f"[OK] Copied config file: {json_file}")
    
    # Copy .env example file
    if os.path.exists(".env.example"):
        shutil.copy2(".env.example", deploy_dir)
        print("[OK] Copied .env example file")
    
    # Copy documentation
    if os.path.exists("README.md"):
        shutil.copy2("README.md", deploy_dir)
    if os.path.exists("PREPROCESS_GUIDE.md"):
        shutil.copy2("PREPROCESS_GUIDE.md", deploy_dir)
    
    # 创建使用说明
    readme_content = """# SeaTable Excel 同步工具部署包

## 使用步骤：

1. 配置环境变量（推荐）：
   cp .env.example .env
   编辑 .env 文件，填入你的SeaTable Token

2. 运行同步工具：
   # Windows:
   seatable-sync.exe
   
   # Linux/macOS:
   ./seatable-sync

## 配置方式：

### .env文件配置（推荐）
1. 复制 .env.example 为 .env
2. 编辑 .env 文件，填入配置信息：
   - SEATABLE_SERVER_URL=你的SeaTable服务器地址
   - DEFAULT_SEATABLE_API_TOKEN=默认API Token
   - MEMO_ANA2025_SEATABLE_API_TOKEN=2025年数据配置专用Token
   - MEMO_ANA2024_SEATABLE_API_TOKEN=2024年数据配置专用Token
3. 直接运行: ./seatable-sync

## 配置文件说明：

- memo-ana2025.json: 2025年目标数据同步配置
- memo-ana2024.json: 2024年目标数据同步配置

每个配置文件支持：
- menu_description: 菜单显示描述
- tables: 表格配置列表
- 独立的SeaTable API Token配置

## 功能特性：

- 支持多配置文件自动发现
- 支持每个配置文件独立的API Token
- 智能菜单系统，显示配置描述
- Excel到SeaTable数据同步
- 字段映射和关联匹配
- 自动生成带日期的备份文件
- 跨平台支持（Windows, Linux, macOS）

## 注意事项：

- 确保网络能访问SeaTable服务
- 确保API Token有相应的表格权限
- Excel文件路径在配置文件中指定
- 支持.xlsx格式Excel文件
- 程序会根据关联字段匹配更新数据
"""
    
    with open(os.path.join(deploy_dir, "USAGE.txt"), "w", encoding="utf-8") as f:
        f.write(readme_content)
    
    print("\n====================================")
    print("[SUCCESS] SeaTable Excel Sync tool package created successfully!")
    print(f"Package location: {deploy_dir}/")
    print(f"Executable: {deploy_dir}/{exe_name}")
    print("Share the entire folder with your team")
    print("====================================")
    
    return True

if __name__ == "__main__":
    success = create_standalone_build()
    if not success:
        sys.exit(1)