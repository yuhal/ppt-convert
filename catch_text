# -*- coding: utf-8 -*-
# !python3
"""
PPT catch TEXT
"""

from pptx import Presentation

def ppt_catch_text(filename):
    """
    Text stores strings in dictionary format,
    one for each text run in presentation
    """
    prs = Presentation(filename)
    text = {}

    for x in range(len(prs.slides)):
        text[x] = []
        for shape in prs.slides[x].shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text[x].append(run.text)

    return text

ppt_catch_text('xxx.pptx')
