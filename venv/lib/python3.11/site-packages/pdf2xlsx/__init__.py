# -*- coding: utf-8 -*-

"""
By default load the managment function and the gui when the module is loaded.
"""
from .managment import do_it
from .gui import main as gui_main
# setup
__version__ = '1.0.0'
__all__ = ["__main__", "managment", "gui", "logger", "config", "utility", "invoice"]
