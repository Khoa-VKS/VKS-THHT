# -*- coding: utf-8 -*-
import requests

class GoogleSheetService:
    @staticmethod
    def push_to_webhook(webhook_url, records):
        """
        Gửi dữ liệu lên Google Apps Script Webhook.
        Payload gồm: { "rows": [ {"name": "...", "yeucau": "...", "bienphap": "..."}, ... ] }
        """
        if not webhook_url:
            return False, "Chưa cung cấp Webhook URL"

        payload = {
            "title": "TỔNG HỢP LỖI THƯỜNG GẶP KHI THAO TÁC NỀN TẢNG QUẢN LÝ ÁN HÌNH SỰ",
            "rows": records
        }

        try:
            # Apps Script Webhook thường redirect 302 đến script.googleusercontent.com
            response = requests.post(webhook_url, json=payload, timeout=20, allow_redirects=True)
            if response.status_code in [200, 201, 302]:
                return True, "Thành công"
            else:
                return False, f"Google Sheet phản hồi mã lỗi: {response.status_code}"
        except Exception as e:
            return False, f"Lỗi kết nối: {str(e)}"
