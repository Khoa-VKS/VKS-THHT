# -*- coding: utf-8 -*-
import re

text = """“TM. -2270.C370/.,-2/0 Cự +5-2722/05/NG:. NA NỘI DUNG YÊU CẦU Trong giai đoạn truy tố, không thao tác thêm mới "Gia hạn thời hạn quyết định việc truy tố" được. Dẫn đến báo quá hạn."""

def clean_yeu_cau(t):
    # Tìm chuỗi NỘI DUNG YÊU CẦU (kể cả NÔI DUNG, NOI DUNG, YÊU CẦU)
    # Tìm kiếm không phân biệt hoa thường và không phân biệt dấu
    patterns = [
        r'n[ộoô]i\s*dung\s*y[êe]u\s*c[ầa]u',
        r'y[êe]u\s*c[ầa]u'
    ]
    for p in patterns:
        m = re.search(p, t, re.IGNORECASE)
        if m:
            return t[m.end():].lstrip(': -.\t\n').strip()
    return t

print("Result:")
print(clean_yeu_cau(text))
