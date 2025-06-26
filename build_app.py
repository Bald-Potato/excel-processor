# build_app.py
# 用于将Python脚本打包为Mac应用程序的脚本

import os
import sys
import subprocess

def run_command(command):
    """运行命令并打印输出"""
    print(f"执行: {command}")
    process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    
    # 实时打印输出
    for line in iter(process.stdout.readline, b''):
        line_str = line.decode('utf-8').rstrip()
        print(line_str)
    
    process.wait()
    return process.returncode

def main():
    # 确保PyInstaller已安装
    print("检查PyInstaller是否已安装...")
    try:
        import PyInstaller
        print("PyInstaller已安装")
    except ImportError:
        print("安装PyInstaller...")
        if run_command("pip install pyinstaller") != 0:
            print("安装PyInstaller失败")
            return 1
    
    # 创建spec文件
    print("\n创建spec文件...")
    spec_content = """
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['frontend_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 创建单文件应用
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel处理工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

# 创建Mac应用程序包
app = BUNDLE(
    exe,
    name='Excel处理工具_单文件.app',
    icon=None,
    bundle_identifier='com.excel.processor',
    info_plist={
        'NSPrincipalClass': 'NSApplication',
        'NSAppleScriptEnabled': False,
        'CFBundleDocumentTypes': [],
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': 'True',
    },
)
"""
    
    with open("Excel处理工具.spec", "w", encoding="utf-8") as f:
        f.write(spec_content)
    
    # 运行PyInstaller
    print("\n开始打包应用程序...")
    if run_command("pyinstaller --clean -y Excel处理工具.spec") != 0:
        print("打包应用程序失败")
        return 1
    
    print("\n应用程序打包完成!")
    print("应用程序位置: dist/Excel处理工具.app")
    print("您可以将此应用程序复制到Applications文件夹中")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())