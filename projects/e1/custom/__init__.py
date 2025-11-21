"""
中国科学技术大学本科毕业论文格式 - 自定义模块

该模块包含了e1项目的解析器、格式化器和样式管理器
"""

from .parser import E1Parser
from .formatter import E1Formatter
from .styles import E1StyleManager

__all__ = ['E1Parser', 'E1Formatter', 'E1StyleManager']
