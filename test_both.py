# -*- coding: utf-8 -*-
import re

def test_cleaning():
    # Case 1: STT 3 from user screenshot
    raw_yc_3 = """“TM. -2270.C370/.,-2/0 Cự +5-2722/05/NG:. NA NỘI DUNG YÊU CẦU Trong giai đoạn truy tố, không thao tác thêm mới "Gia hạn thời hạn quyết định việc truy tố" được. Dẫn đến báo quá hạn."""
    
    # Case 2: STT 1 from user screenshot
    raw_bp_1 = """1. Sau khi tách vụ án, người dùng không kiểm tra BPNC đã được tách cùng bị can hay chưa.
2. Đã thực hiện chuyển vụ án khi chưa kiểm tra phần BPNC trong vụ án sau khi tách.
3. Khu vực 17 đã hoàn thành luôn việc nhập vụ án. .SEaS mm CỔ L4 ki mkl A k"""

    # 1. Clean Yêu Cầu
    m_yc = re.search(r'(?:n[ộoô]i\s*dung\s*)?y[êe]u\s*c[ầa]u', raw_yc_3, re.IGNORECASE)
    clean_yc_3 = raw_yc_3[m_yc.end():].lstrip(': -.\t\n').strip() if m_yc else raw_yc_3

    # 2. Clean Biện Pháp
    lines = raw_bp_1.split('\n')
    last_line = lines[-1].strip()
    # Tìm dấu chấm câu tiếng Việt, sau đó là rác
    # Ví dụ: "nhập vụ án. .SEaS mm CỔ L4 ki mkl A k"
    m_junk = re.search(r'([\.\?\!])\s*(\.|\-|\_|[A-Za-z0-9\s\.\,\/À-ỹ]{3,})$', last_line)
    if m_junk:
        # Cắt đúng tại dấu chấm câu kết thúc ý
        # Kiểm tra xem có phải đuôi rác (có double period hoặc từ như SEaS)
        tail = m_junk.group(0)
        if '..' in tail or '. .' in tail or 'SEaS' in tail or 'CỔ L4' in tail:
            last_line = last_line[:m_junk.start(1) + 1].strip()
            lines[-1] = last_line
    clean_bp_1 = '\n'.join(lines)

    print("--- CLEANED YÊU CẦU 3 ---")
    print(clean_yc_3)
    print("\n--- CLEANED BIỆN PHÁP 1 ---")
    print(clean_bp_1)

test_cleaning()
