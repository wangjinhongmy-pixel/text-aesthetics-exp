import json
from collections import Counter

with open('results.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

names = [r['被试姓名'] for r in data]
print("Unique subjects:", set(names))
print("Count per subject:", Counter(names))
print("Total records:", len(data))
