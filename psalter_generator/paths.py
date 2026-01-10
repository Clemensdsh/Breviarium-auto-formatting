# psalter_generator/paths.py
import sys
import os

def get_base_path():
    """获取资源文件的基础路径"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的临时目录
        return sys._MEIPASS
    else:
        # 开发环境 - 返回项目根目录（main.py所在目录）
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def resource_path(relative_path):
    """获取资源文件的绝对路径"""
    return os.path.join(get_base_path(), relative_path)

# 常用路径
CONTENT_DIR = resource_path("content")
GABC_DIR = resource_path("gabc")
IMAGES_DIR = resource_path("images")
ICON_PATH = resource_path("cross_of_saint_stephen.png")