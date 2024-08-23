# -*- coding: utf-8 -*-
"""
Fire up the GUI by default
"""
from .gui import main as gui_main
from .config import init_conf

init_conf()
gui_main()
