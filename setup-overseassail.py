#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from cx_Freeze import setup, Executable
import matplotlib

include_files = []
include_files.append((matplotlib.get_data_path(), 'mpl-data'))

############ Force copy of basemap data from local directory ###################
include_files.append((r"C:\Python27\Lib\site-packages\mpl_toolkits\basemap\data", 'mpl-basemap-data'))

build_exe_options = {"packages": ["LatLon", "sys"],
                     "excludes": ['collections.abc'],
                     "include_files": include_files,
                     "includes": ["matplotlib.backends.backend_tkagg"]}

base = None

setup(  name = "overseassail-gen",
        version = "0.3",
        description = "overseassail-gen",
        options = {"build_exe": build_exe_options},
        executables = [Executable("overseassail-gen.py", base=base)])




#####
"""

0. Force copy of basemap data in setup.py

1. added import for FileDialog, pyproj in the main.py
2. modify import of basemap in the main.py

import sys, os
if getattr(sys, 'frozen', False):
    os.environ['BASEMAPDATA'] = os.path.join(os.path.dirname(sys.executable), 'mpl-basemap-data')
from mpl_toolkits.basemap import Basemap
import FileDialog, pyproj

3. In line 238 ofPython27\Lib\site-pakages\mpl-toolkits\basemap\pyproj.py

#pyproj_datadir = os.sep.join([os.path.dirname(__file__), 'data'])
import sys
pyproj_datadir = os.path.join(os.path.dirname(sys.executable), 'mpl-basemap-data')
"""
