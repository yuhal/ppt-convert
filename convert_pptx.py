# -*- coding: utf-8 -*-
# !python3
"""
PPT convert PPTX
"""

from changeOffice import Change

c = Change("./ppt")

c.ppt2pptx()

print(c.get_allPath())
