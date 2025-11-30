"""
打包脚本 - 将PCR分析软件打包为Windows可执行文件
"""
import PyInstaller.__main__
import os
import sys

def build_exe():
    """构建可执行文件"""
    
    # PyInstaller参数
    args = [
        'pcr_analyzer.py',           # 主程序文件
        '--name=PCR分析软件',         # 可执行文件名称
        '--onefile',                  # 打包为单个文件
        '--windowed',                 # 无控制台窗口（GUI应用）
        '--icon=NONE',                # 图标文件（如果有）
        '--add-data=excel_parser.py;.',  # 包含解析器模块
        '--add-data=data_visualizer.py;.', # 包含可视化模块
        '--hidden-import=openpyxl',   # 隐藏导入
        '--hidden-import=pandas',      # 隐藏导入
        '--hidden-import=numpy',       # 隐藏导入
        '--hidden-import=matplotlib',  # 隐藏导入
        '--hidden-import=PyQt5',       # 隐藏导入
        '--collect-all=matplotlib',   # 收集matplotlib所有数据
        '--collect-all=PyQt5',        # 收集PyQt5所有数据
    ]
    
    print("开始打包...")
    print("命令:", ' '.join(args))
    
    try:
        PyInstaller.__main__.run(args)
        print("\n打包完成！可执行文件位于 dist 目录")
    except Exception as e:
        print(f"打包失败: {e}")
        sys.exit(1)

if __name__ == '__main__':
    build_exe()

