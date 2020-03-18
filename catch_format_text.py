# -*- coding: utf-8 -*-
# !python3
"""
PPT catch TEXT
"""

from pptx import Presentation

def ppt_catch_format_text(filename):
    """
    Text stores strings in dictionary format,
    And according to the paragraph format of the text,
    For every text run in the demo,
    """
    prs = Presentation(filename)
    txt_oa = {}
    for x in range(len(prs.slides)):
        txt_oa[x] = []
        for shape in prs.slides[x].shapes:

            if shape.shape_type._member_name == 'TEXT_BOX' \
            or shape.shape_type._member_name == 'AUTO_SHAPE' \
            or shape.shape_type._member_name == 'PLACEHOLDER':
                shape_txt = shape.text.encode('utf-8').strip().decode()

                if len(shape_txt) > 0 :
                    txt_oa[x].extend( shape_txt.split('\n') )

            if shape.shape_type._member_name == 'TABLE':
                tb = shape.table
                tb_row_size = len(shape.table.rows)
                tb_col_size = len(shape.table.columns)

                for ri in range(0,tb_row_size):
                    row_text_da = []

                    for ci in range(0,tb_col_size):
                        row_text_da.append(tb.cell(ri,ci).text_frame.text)

                    row_text = ' '.join(row_text_da).encode('utf-8').strip().decode()
                    if len(row_text) > 0 :
                        txt_oa[x].append( row_text )
    return txt_oa

print(ppt_catch_format_text('xxx.pptx'))
