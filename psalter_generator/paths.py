"""资源路径（兼容 PyInstaller 打包）"""

import os
import sys

_FROZEN = hasattr(sys, '_MEIPASS')
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 只读资源目录：打包后为临时解压目录，开发时为项目根目录
RESOURCE_DIR = sys._MEIPASS if _FROZEN else _PROJECT_ROOT
# 工作目录：打包后为 exe 所在目录，开发时为项目根目录
WORKING_DIR = os.path.dirname(sys.executable) if _FROZEN else _PROJECT_ROOT

ICON_PATH = os.path.join(RESOURCE_DIR, "cross_of_saint_stephen.png")
