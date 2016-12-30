#!/usr/bin/python
# -*- coding: UTF-8 -*-

from distutils.core import setup  
import py2exe  
  
options = {"py2exe":  
            {   "compressed": 1,     
                "optimize": 2,      
                "bundle_files": 1   #所有文件打包成一个exe文件
            }     
          }     
setup(        
    version = "1.0.0",     
    description = u"顺力机械生产型号合并软件",     
    name = u"顺力机械生产型号合并软件",     
    options = options,     
    zipfile = None, # 不生成zip库文件    
    console = [{"script": "sljx_merge.py", "icon_resources": [(0, "tigeek.ico")] }],       
    )

