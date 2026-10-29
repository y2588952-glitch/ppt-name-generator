# -*- coding: utf-8 -*-
"""pytest 配置：确保能 import 项目根目录下的 ppt_processor。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
