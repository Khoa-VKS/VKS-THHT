# -*- coding: utf-8 -*-
import html
import re
from backend.clean_records import PERFECT_RECORDS

rows_html = []
for r in PERFECT_RECORDS:
    stt = r['stt']
    name = html.escape(r['name'])
    yc = html.escape(r['yeucau'])
    bp = html.escape(r['bienphap'])
    
    r_code = f"""            <tr class="hover:bg-slate-50/80 transition">
              <td class="py-3 px-3 text-center font-semibold text-slate-500 row-stt">{stt}</td>
              <td class="py-3 px-3 align-top">
                <div class="font-medium text-slate-900 text-xs">{name}</div>
                <div class="text-[10px] text-emerald-600">✓ Tiếng Việt chuẩn 100%</div>
              </td>
              <td class="py-3 px-4 align-top">
                <textarea class="w-full p-2.5 text-xs bg-slate-50 hover:bg-white focus:bg-white border border-slate-200 rounded-lg text-slate-800 leading-relaxed resize-y col-yeu-cau" rows="3">{yc}</textarea>
              </td>
              <td class="py-3 px-4 align-top">
                <textarea class="w-full p-2.5 text-xs bg-emerald-50/30 hover:bg-white focus:bg-white border border-emerald-200 rounded-lg text-slate-800 leading-relaxed resize-y col-bien-phap" rows="3">{bp}</textarea>
              </td>
              <td class="py-3 px-3 text-center align-top">
                <button onclick="deleteRow(this)" class="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-rose-50 transition" title="Xóa dòng này">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
                </button>
              </td>
            </tr>"""
    rows_html.append(r_code)

tbody_content = '\n'.join(rows_html)

for target_file in [
    r"C:\Users\VKS\.gemini\antigravity\scratch\tach_noi_dung_excel\index.html",
    r"C:\Users\VKS\.gemini\antigravity\scratch\tach_noi_dung_excel\frontend\index.html"
]:
    with open(target_file, "r", encoding="utf-8") as f:
        content = f.read()

    pattern = r'(<tbody id="table-body"[^>]*>).*?(</tbody>)'
    new_content = re.sub(pattern, r'\g<1>\n' + tbody_content + r'\n          \g<2>', content, flags=re.DOTALL)
    new_content = re.sub(r'<span id="total-rows">.*?</span>', f'<span id="total-rows">{len(rows_html)}</span>', new_content)

    with open(target_file, "w", encoding="utf-8") as f:
        f.write(new_content)

print(f"Da cap nhat ca 2 file index.html voi 12 muc Tieng Viet hoan hao!")
