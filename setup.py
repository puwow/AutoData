#-*- coding:gbk -*-
from distutils.core import setup
import py2exe
import sys
from glob import glob
import os
import shutil

delFiles=["build","dist"]
for item in delFiles:
    if os.path.isdir(item):
        print "remove directory: %s"%(item)

sys.argv.append("py2exe")
includes=["dbhash","appdirs","packaging","packaging.version","packaging.specifiers","packaging.requirements"]
options={"py2exe":{"packages":["pytz","packaging"],"includes":includes,"compressed":1, "optimize":1, "bundle_files":1, "dll_excludes":["MSVCP90.dll"]}}
data_files=[("images", glob(r'images/*.png')),("images",glob(r'images/*.ico')), (".", glob(r'config.ini')), (".",glob(r'memo.she')), (".", glob(r'version.md'))]
setup(
        windows=[{"script":'MBKAuto.py', "icon_resources":[(1, "images/app2.ico")]}], 
        options=options, 
        data_files=data_files, 
        name="MBKAutoData",
        description=u'手机银行自动化测试数据管理工具',
        version="1.1.0"
        )
