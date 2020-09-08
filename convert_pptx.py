# -*- coding: utf-8 -*-
# !python3
"""
PPT convert PPTX
"""

from changeOffice import Change

change = Change("./")

change.ppt2pptx()

print(change.get_allPath())
