"""
将桌面文本数据加载为JSON，供实验系统使用。
运行一次即可：python load_texts.py
"""
import openpyxl
import json
import os

SRC = r"C:\Users\xiaixin15\Desktop\文本.xlsx"
DST = os.path.join(os.path.dirname(__file__), "texts_data.json")

# 修正路径（用户桌面的实际路径）
SRC = r"C:\Users\xiaoxin15\Desktop\文本.xlsx"
DST = r"C:\Users\xiaoxin15\Desktop\文本审美对比实验系统\texts_data.json"

wb = openpyxl.load_workbook(SRC)
ws = wb.active

texts = {}
for row in ws.iter_rows(min_row=2, values_only=True):
    idx, level, char_len, content = row[0], row[1], row[2], row[3]
    texts[int(idx)] = {
        "id": int(idx),
        "level": int(level),
        "char_length": int(char_len),
        "text": str(content)
    }

with open(DST, "w", encoding="utf-8") as f:
    json.dump(texts, f, ensure_ascii=False, indent=2)

print(f"已生成 {len(texts)} 条文本数据 → {DST}")
