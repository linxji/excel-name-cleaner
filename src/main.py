# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""Excel 名称清理工具 —— 入口。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.app import main  # noqa: E402

if __name__ == "__main__":
    main()
