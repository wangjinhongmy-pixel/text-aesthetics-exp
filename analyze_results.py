import json
from collections import Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import shutil
from datetime import datetime

# ========== 1. 读取数据 ==========
with open('results.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

subjects = list(set(r['被试姓名'] for r in data))
total_trials = len(data)
subject_name = subjects[0] if len(subjects) == 1 else ', '.join(subjects)

# ========== 2. 统计每个文本出现次数 ==========
all_ids = []
for r in data:
    all_ids.append(r['文本A ID'])
    all_ids.append(r['文本B ID'])
id_counts = Counter(all_ids)

# ========== 3. 计算各项指标 ==========

# 汇总统计
total_texts = 100
n_subjects = len(subjects)
avg_appearance = round(total_trials * 2 / total_texts, 2)
min_appearance = min(id_counts.values())
max_appearance = max(id_counts.values())

# 分档次分析：每次对比中，高档是"正确答案"
# 统计每个档次作为正确答案（高档）出现的次数，以及被选中的次数
level_stats = {lv: {'correct_appear': 0, 'selected': 0, 'total_pairs': 0} for lv in [1,2,3,4,5]}

for r in data:
    lv_a = int(r['文本A档次'].replace('第','').replace('档',''))
    lv_b = int(r['文本B档次'].replace('第','').replace('档',''))
    higher_lv = max(lv_a, lv_b)
    lower_lv = min(lv_a, lv_b)
    choice = r['被试选择']  # '文本A' or '文本B'
    choice_lv = lv_a if choice == '文本A' else lv_b

    # 高档是正确答案
    level_stats[higher_lv]['correct_appear'] += 1
    # 被试选中的档次
    level_stats[choice_lv]['selected'] += 1
    # 该对比组合的总次数
    level_stats[lower_lv]['total_pairs'] += 1
    level_stats[higher_lv]['total_pairs'] += 1

# 计算每档准确率：高档被选中的比例
for lv in [1,2,3,4,5]:
    ca = level_stats[lv]['correct_appear']
    sel = level_stats[lv]['selected']
    if ca > 0:
        level_stats[lv]['accuracy'] = round(sel / ca * 100, 2)
    else:
        level_stats[lv]['accuracy'] = None
    # 被选中比例
    total_appear = id_counts.get(lv, 0)
    level_stats[lv]['selected_ratio'] = round(sel / total_appear * 100, 2) if total_appear > 0 else 0

# 对比矩阵

# 正确的对比矩阵计算
combos = [
    (1,2),(1,3),(1,4),(1,5),
    (2,3),(2,4),(2,5),
    (3,4),(3,5),(4,5)
]

combo_results = {c: {'count': 0, 'choose_higher': 0} for c in combos}

for r in data:
    lv_a = int(r['文本A档次'].replace('第','').replace('档',''))
    lv_b = int(r['文本B档次'].replace('第','').replace('档',''))
    higher_lv = max(lv_a, lv_b)
    lower_lv = min(lv_a, lv_b)
    choice = r['被试选择']
    choice_lv = lv_a if choice == '文本A' else lv_b

    key = (lower_lv, higher_lv)
    combo_results[key]['count'] += 1
    if choice_lv == higher_lv:
        combo_results[key]['choose_higher'] += 1

# 每条文本统计
text_stats = {}
for i in range(1, 101):
    text_stats[i] = {
        'appearances': id_counts.get(i, 0),
        'chosen_as_A': 0,
        'chosen_as_B': 0,
        'chosen_as_correct': 0,
    }

for r in data:
    id_a = r['文本A ID']
    id_b = r['文本B ID']
    lv_a = int(r['文本A档次'].replace('第','').replace('档',''))
    lv_b = int(r['文本B档次'].replace('第','').replace('档',''))
    higher_lv = max(lv_a, lv_b)
    choice = r['被试选择']

    if choice == '文本A':
        text_stats[id_a]['chosen_as_A'] += 1
        if lv_a == higher_lv:
            text_stats[id_a]['chosen_as_correct'] += 1
        text_stats[id_b]['chosen_as_B'] += 1
        if lv_b == higher_lv:
            text_stats[id_b]['chosen_as_correct'] += 1
    else:
        text_stats[id_b]['chosen_as_B'] += 1
        if lv_b == higher_lv:
            text_stats[id_b]['chosen_as_correct'] += 1
        text_stats[id_a]['chosen_as_A'] += 1
        if lv_a == higher_lv:
            text_stats[id_a]['chosen_as_correct'] += 1

# 档次选择偏好
pref_stats = {lv: {'as_low_appear': 0, 'selected': 0} for lv in [1,2,3,4,5]}
for r in data:
    lv_a = int(r['文本A档次'].replace('第','').replace('档',''))
    lv_b = int(r['文本B档次'].replace('第','').replace('档',''))
    higher_lv = max(lv_a, lv_b)
    lower_lv = min(lv_a, lv_b)
    choice = r['被试选择']
    choice_lv = lv_a if choice == '文本A' else lv_b

    pref_stats[lower_lv]['as_low_appear'] += 1
    pref_stats[choice_lv]['selected'] += 1

# 总体准确率（选高档的比例）
total_correct = sum(combo_results[c]['choose_higher'] for c in combos)
total_comparisons = sum(combo_results[c]['count'] for c in combos)
overall_accuracy = round(total_correct / total_comparisons * 100, 2) if total_comparisons > 0 else 0

# ========== 4. 写入Excel ==========
template_path = r'C:\Users\xiaoxin15\Desktop\数据模板\文本对比实验结果模板.xlsx'
output_path = rf'C:\Users\xiaoxin15\Desktop\文本对比实验结果_{subject_name}_{datetime.now().strftime("%Y%m%d")}.xlsx'

shutil.copy(template_path, output_path)
wb = openpyxl.load_workbook(output_path)

# --- Sheet1: 原始数据 ---
ws1 = wb['原始数据']
for idx, r in enumerate(data, 1):
    row = [
        idx,
        r['被试姓名'],
        r['文本A ID'],
        r['文本A内容'],
        r['文本A档次'],
        r['文本B ID'],
        r['文本B内容'],
        r['文本B档次'],
        r['被试选择'],
        r['按键'],
        r['选择时间戳']
    ]
    ws1.append(row)

# --- Sheet2: 汇总统计 ---
ws2 = wb['汇总统计']
# 固定9行结构
summary_data = [
    ('被试姓名', subject_name, ''),
    ('实验日期', datetime.now().strftime('%Y-%m-%d'), ''),
    ('总对比次数', total_trials, ''),
    ('有效次数', total_trials, ''),
    ('无效次数', 0, ''),
    ('文本总数', total_texts, ''),
    ('参与被试人数', n_subjects, ''),
    ('每条文本平均出现次数', avg_appearance, ''),
    ('每条文本最少/最多出现次数', f'{min_appearance}/{max_appearance}', ''),
]
for i, (k, v, _) in enumerate(summary_data, 2):
    ws2.cell(row=i, column=1, value=k)
    ws2.cell(row=i, column=2, value=v)

# --- Sheet3: 分档次分析 ---
ws3 = wb['分档次分析']
for i, lv in enumerate([1,2,3,4,5], 2):
    ca = level_stats[lv]['correct_appear']
    sel = level_stats[lv]['selected']
    total_appear = id_counts.get(lv, 0)
    sel_ratio = round(sel / total_appear * 100, 2) if total_appear > 0 else 0
    theory_ratio = 80  # 高档作为正确答案的比例
    bias = round(sel_ratio - theory_ratio, 2)
    acc = level_stats[lv]['accuracy']

    ws3.cell(row=i, column=1, value=f'第{lv}档')
    ws3.cell(row=i, column=2, value=ca)
    ws3.cell(row=i, column=3, value=sel)
    ws3.cell(row=i, column=4, value=f'{sel_ratio}%')
    ws3.cell(row=i, column=5, value=f'{theory_ratio}%')
    ws3.cell(row=i, column=6, value=f'{bias:+.2f}%')
    ws3.cell(row=i, column=7, value=f'{acc}%' if acc is not None else 'N/A')

# --- Sheet4: 对比矩阵 ---
ws4 = wb['对比矩阵']
for i, c in enumerate(combos, 2):
    cnt = combo_results[c]['count']
    ch = combo_results[c]['choose_higher']
    acc_c = round(ch / cnt * 100, 2) if cnt > 0 else 0
    ws4.cell(row=i, column=1, value=f'第{c[0]}档 vs 第{c[1]}档')
    ws4.cell(row=i, column=2, value=cnt)
    ws4.cell(row=i, column=3, value=ch)
    ws4.cell(row=i, column=4, value=cnt - ch)
    ws4.cell(row=i, column=5, value=f'{acc_c}%')

# --- Sheet5: 每条文本统计 ---
ws5 = wb['每条文本统计']
# 先加载texts_data获取档次
with open('texts_data.json', 'r', encoding='utf-8') as f:
    texts_db = json.load(f)

for i in range(1, 101):
    lv = texts_db[str(i)]['level']
    stats = text_stats[i]
    total_app = stats['appearances']
    chosen_total = stats['chosen_as_A'] + stats['chosen_as_B']
    chosen_ratio = round(chosen_total / total_app * 100, 2) if total_app > 0 else 0

    ws5.cell(row=i+1, column=1, value=i)
    ws5.cell(row=i+1, column=2, value=f'第{lv}档')
    ws5.cell(row=i+1, column=3, value=total_app)
    ws5.cell(row=i+1, column=4, value=stats['chosen_as_A'])
    ws5.cell(row=i+1, column=5, value=stats['chosen_as_B'])
    ws5.cell(row=i+1, column=6, value=stats['chosen_as_correct'])
    ws5.cell(row=i+1, column=7, value=f'{chosen_ratio}%')

# --- Sheet6: 档次选择偏好 ---
ws6 = wb['档次选择偏好']
for i, lv in enumerate([1,2,3,4,5], 2):
    as_low = pref_stats[lv]['as_low_appear']
    selected = pref_stats[lv]['selected']
    total_app = id_counts.get(lv, 0)
    sel_rate = round(selected / total_app * 100, 2) if total_app > 0 else 0
    theory_prob = 20  # 各档均等概率
    bias = round(sel_rate - theory_prob, 2)

    ws6.cell(row=i, column=1, value=f'第{lv}档')
    ws6.cell(row=i, column=2, value=as_low)
    ws6.cell(row=i, column=3, value=selected)
    ws6.cell(row=i, column=4, value=f'{sel_rate}%')
    ws6.cell(row=i, column=5, value=f'{theory_prob}%')
    ws6.cell(row=i, column=6, value=f'{bias:+.2f}%')

wb.save(output_path)
print(f"\n[Saved] {output_path}")

# 打印关键统计
print(f"\n=== 关键统计 ===")
print(f"被试: {subject_name}")
print(f"总对比次数: {total_trials}")
print(f"总体准确率（选高档）: {overall_accuracy}%")
print(f"文本出现次数范围: {min_appearance}~{max_appearance}次")
