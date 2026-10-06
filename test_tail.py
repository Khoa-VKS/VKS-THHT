# -*- coding: utf-8 -*-
import re

sample_bp1 = """1. Sau khi tách vụ án, người dùng không kiểm tra BPNC đã được tách cùng bị can hay chưa.
2. Đã thực hiện chuyển vụ án khi chưa kiểm tra phần BPNC trong vụ án sau khi tách.
3. Khu vực 17 đã hoàn thành luôn việc nhập vụ án. .SEaS mm CỔ L4 ki mkl A k"""

def clean_bottom_junk(text):
    # Tách các dòng
    lines = text.strip().split('\n')
    if not lines:
        return ""
    
    # Kiểm tra dòng cuối cùng
    last_line = lines[-1].strip()
    
    # Nếu dòng cuối có dấu chấm câu kết thúc tiếng Việt bình thường, nhưng theo sau là chuỗi rác
    # Ví dụ: "...nhập vụ án. .SEaS mm CỔ L4 ki mkl A k"
    # Tìm mẫu: sau dấu chấm (hoặc dấu chấm than, hỏi) là các ký tự rác hỗn tạp
    m = re.search(r'([\.\?\!])\s+([A-Za-z0-9\.\_\-\/\s]{4,})$', last_line)
    if m:
        tail = m.group(2).strip()
        # Kiểm tra xem tail có phải là chuỗi rác không (chứa chữ hoa lộn xộn, từ vô nghĩa không phải tiếng Việt bình thường)
        # Nếu có các từ như SEaS, CỔ, L4, ki, mkl
        words = tail.split()
        # Nếu có từ vô nghĩa hoặc ngắn
        last_line = last_line[:m.start(1) + 1].strip()
        lines[-1] = last_line
        
    return '\n'.join(lines).strip()

print("Original:")
print(sample_bp1)
print("\nCleaned:")
print(clean_bottom_junk(sample_bp1))
