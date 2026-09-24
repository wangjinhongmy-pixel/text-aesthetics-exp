"""
文本审美对比实验系统
部署：Streamlit Cloud（免费）
"""

import streamlit as st
import json
import random
import openpyxl
from datetime import datetime
import os

# ============================================================
# 配置
# ============================================================
PAGE_CONFIG = {
    "page_title": "文本审美对比实验",
    "page_icon": "",
    "layout": "wide"
}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE_DIR, "results.json")
GLOBAL_COUNT_FILE = os.path.join(BASE_DIR, "global_counts.json")
SEQ_COUNTER_FILE = os.path.join(BASE_DIR, "seq_counter.json")
TEXTS_PATH = os.path.join(BASE_DIR, "texts_data.json")

TRIALS_PER_SUBJECT = 50
LEVELS = [1, 2, 3, 4, 5]

LEVEL_COMBO_WEIGHTS = {
    (1,2): 3, (2,3): 3, (3,4): 3, (4,5): 3,
    (1,3): 2, (2,4): 2, (3,5): 2,
    (1,4): 1, (2,5): 1, (1,5): 1,
}

# 蓝色主题色
BLUE = "#1a56db"
LIGHT_BLUE = "#e8f0fe"
WHITE = "#ffffff"

# ============================================================
# 全局样式
# ============================================================
st.markdown(f"""
<style>
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
.stDeployButton {{display: none;}}

/* 页面背景 */
.stApp {{background-color: #f0f4ff;}}

/* 标题居中 - 超大加粗 */
.center-title {{text-align: center !important; font-size: 4.5rem !important; font-weight: bold !important; color: {BLUE} !important; margin-top: 8rem !important; margin-bottom: 2.5rem !important;}}

/* 正文区域 */
.block-container {{padding-top: 2.5rem;}}

/* 文本框 - 蓝色边框白底，固定高度统一大小 */
.text-box {{
    background-color: {WHITE};
    border: 2px solid {BLUE};
    border-radius: 8px;
    padding: 1.2rem;
    height: 220px;
    display: flex;
    align-items: center;
    justify-content: center;
    text-align: center;
    font-size: 1.05rem;
    line-height: 2;
    color: #1a1a1a;
    word-break: break-word;
    overflow: hidden;
    width: 100%;
    box-sizing: border-box;
}}

/* 文本框内文字 */
.text-box-inner {{
    width: 100%;
}}

/* 主按钮 - 居中 */
.stButton > button {{
    background-color: {BLUE};
    color: white;
    border: none;
    border-radius: 6px;
    font-size: 1rem;
    padding: 0.6rem 2rem;
    width: 100%;
    max-width: 800px;
    margin: 0 auto;
    display: block;
}}
.stButton > button:hover {{background-color: #1342b8;}}

/* 表单提交按钮 */
.stFormSubmitButton > button {{
    background-color: {BLUE};
    color: white;
    border: none;
    border-radius: 6px;
    font-size: 1.05rem;
    padding: 0.6rem 2rem;
    width: 66%;
    margin: 0 auto;
    display: block;
}}
.stFormSubmitButton > button:hover {{background-color: #1342b8;}}

/* Radio 选项 */
.stRadio > div {{
    padding: 0.3rem 0;
}}

/* 进度条 */
.stProgress > div > div > div {{
    background-color: {BLUE};
}}

/* 成功/提示框 */
.stAlert {{
    border-left: 4px solid {BLUE};
}}

/* 实验页按钮统一位置 */
.fix-bottom {{
    margin-top: 1.5rem;
}}

/* 禁用所有输入框的浏览器自动填充 */
input {{
    autocomplete: off !important;
    -webkit-autofill: none !important;
    background-color: #fff !important;
}}

/* 年龄输入框：纯文本样式，无stepper */
input[type="number"] {{
    -moz-appearance: textfield !important;
}}
input[type="number"]::-webkit-outer-spin-button,
input[type="number"]::-webkit-inner-spin-button {{
    -webkit-appearance: none !important;
    margin: 0 !important;
}}
</style>
""", unsafe_allow_html=True)

# ============================================================
# 初始化
# ============================================================
def load_texts():
    with open(TEXTS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def init_session_state():
    defaults = {
        "phase": "consent",
        "subject_name": None,
        "answers": {},
        "trial_index": 0,
        "trials": [],
        "responses": [],
        "exp_start_time": None,
        "submitted": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def load_global_counts():
    if os.path.exists(GLOBAL_COUNT_FILE):
        with open(GLOBAL_COUNT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_global_counts(counts):
    with open(GLOBAL_COUNT_FILE, "w", encoding="utf-8") as f:
        json.dump(counts, f, ensure_ascii=False)

def load_seq_counter():
    """加载序号计数器，无则从1开始"""
    if os.path.exists(SEQ_COUNTER_FILE):
        try:
            with open(SEQ_COUNTER_FILE, "r", encoding="utf-8") as f:
                return json.load(f).get("next_seq", 1)
        except:
            return 1
    return 1

def save_seq_counter(next_seq):
    """保存序号计数器"""
    with open(SEQ_COUNTER_FILE, "w", encoding="utf-8") as f:
        json.dump({"next_seq": next_seq}, f, ensure_ascii=False)

def append_records_to_json(records):
    """
    将新记录追加写入 results.json。
    兼容手动删除部分数据后文件格式不完整的情况，
    自动检测并修复文件结构。
    """
    os.makedirs(os.path.dirname(JSON_PATH), exist_ok=True)

    if not os.path.exists(JSON_PATH):
        with open(JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        return

    # 文件存在，读取现有内容并追加
    try:
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            content = f.read().strip()

        if not content:
            # 空文件，直接写
            existing = []
        elif content.startswith("["):
            existing = json.loads(content)
        else:
            # 文件损坏（非JSON格式），备份后重新写
            backup_path = JSON_PATH + ".broken_" + datetime.now().strftime("%Y%m%d%H%M%S")
            with open(backup_path, "w", encoding="utf-8") as f:
                f.write(content)
            existing = []
    except json.JSONDecodeError:
        # JSON解析失败，备份损坏文件
        backup_path = JSON_PATH + ".broken_" + datetime.now().strftime("%Y%m%d%H%M%S")
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            with open(backup_path, "w", encoding="utf-8") as bf:
                bf.write(f.read())
        existing = []

    existing.extend(records)
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)


# ============================================================
# 试次生成
# ============================================================
def weighted_choice(items, counts_dict, power=1.5):
    """
    根据当前参与次数加权抽样：被选次数越少的文本，概率越高。
    items: ID列表
    counts_dict: {id: 当前参与次数}
    power: 权重幂次，越大越偏向少选的文本
    """
    counts = [counts_dict.get(i, 0) for i in items]
    # 参与次数越多，权重越低
    weights = [(max(counts) + 1 - c) ** power for c in counts]
    total = sum(weights)
    r = random.random() * total
    cumsum = 0
    for i, w in enumerate(weights):
        cumsum += w
        if cumsum >= r:
            return items[i]
    return items[-1]


def generate_trials(texts_db, n_trials=50, global_counts=None):
    """
    生成 n_trials 次成对比较试次。
    global_counts: {文本ID: 当前参与次数}，用于加权均衡每条文本的出现频率。
    参与次数越少的文本被选中的概率越高，实现 10~20 次的均匀分布（无硬上限）。
    """
    if global_counts is None:
        global_counts = {}

    by_level = {lv: [] for lv in LEVELS}
    for tid, t in texts_db.items():
        by_level[t["level"]].append(int(tid))

    combo_pool = []
    for (la, lb), w in LEVEL_COMBO_WEIGHTS.items():
        for _ in range(w):
            combo_pool.append((la, lb))

    trials = []
    attempts = 0

    while len(trials) < n_trials and attempts < 5000:
        attempts += 1
        random.shuffle(combo_pool)
        for la, lb in combo_pool:
            if len(trials) >= n_trials:
                break

            ca = by_level[la]
            cb = by_level[lb]

            # 加权抽样：被选次数少的文本优先
            ta = weighted_choice(ca, global_counts, power=1.5)
            tb = weighted_choice(cb, global_counts, power=1.5)

            # 避免选到同一文本
            if ta == tb:
                cb_others = [x for x in cb if x != ta]
                if not cb_others:
                    continue
                tb = weighted_choice(cb_others, global_counts, power=1.5)

            if random.random() < 0.5:
                left_id, right_id = tb, ta
                left_lv, right_lv = lb, la
            else:
                left_id, right_id = ta, tb
                left_lv, right_lv = la, lb

            trials.append({
                "trial_num": len(trials) + 1,
                "left_id": left_id,
                "right_id": right_id,
                "left_level": left_lv,
                "right_level": right_lv,
            })

    return trials

# ============================================================
# 数据写入
# ============================================================
def write_results(subject_name, answers, responses, trials, texts_db):
    """
    追加写入本次实验数据到 results.json。
    每条记录包含：
    - 第一阶段问卷全部答案（姓名、手机号、性别、年龄等）
    - 第二阶段本试次的对比选择结果
    同一被试的每个试次为一条独立记录，
    通过"被试姓名 + 实验开始时间"与手机号后四位关联前后两阶段数据。
    序号 = 现有最大序号 +1，自动连贯，不受手动删除影响。
    """
    gcounts = load_global_counts()

    # 读取已有数据，动态计算下一个序号（现有最大 +1）
    all_records = []
    if os.path.exists(JSON_PATH):
        try:
            with open(JSON_PATH, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    all_records = json.loads(content)
        except (json.JSONDecodeError, IOError):
            all_records = []

    # 自动编号：现有最大序号 +1，保证连贯
    if all_records:
        max_seq = max(r.get("序号", 0) for r in all_records)
        next_seq = max_seq + 1
    else:
        next_seq = 1

    exp_start = st.session_state.get("exp_start_time", "")

    for trial, resp in zip(trials, responses):
        lid = trial["left_id"]
        rid = trial["right_id"]
        lv_l = trial["left_level"]
        lv_r = trial["right_level"]
        chosen = resp["choice"]

        # 合并问卷答案 + 本试次结果
        record = {
            # 问卷答案（第一阶段）
            "序号": next_seq,
            "被试姓名": subject_name,
            "手机号后四位": answers.get("phone", "")[-4:] if answers.get("phone") else "",
            "实验开始时间": exp_start,
            "性别": answers.get("gender", ""),
            "年龄": answers.get("age", ""),
            "户籍类型": answers.get("hukou", ""),
            "在读教育阶段": answers.get("edu", ""),
            "专业大类": answers.get("major_cat", ""),
            "具体专业": answers.get("major_detail", "") or "（非其他）",
            "LLM使用频率": answers.get("llm_freq", ""),
            "读过AI文学": answers.get("ai_read", ""),
            "文学作品阅读频率": answers.get("read_freq", ""),
            "阅读最看重维度": answers.get("read_dim", ""),
            "阅读偏好": answers.get("aesthetic_pref", ""),
            # 审美对比结果（第二阶段）
            "试次编号": trial["trial_num"],
            "文本A ID": lid,
            "文本A内容": texts_db[str(lid)]["text"],
            "文本A档次": f"第{lv_l}档",
            "文本B ID": rid,
            "文本B内容": texts_db[str(rid)]["text"],
            "文本B档次": f"第{lv_r}档",
            "被试选择": f"文本{'A' if chosen == 'left' else 'B'}",
            "按键": "A" if chosen == "left" else "B",
            "选择时间戳": resp["timestamp"],
        }
        all_records.append(record)
        gcounts[str(lid)] = gcounts.get(str(lid), 0) + 1
        gcounts[str(rid)] = gcounts.get(str(rid), 0) + 1
        next_seq += 1

    # 追加写入
    append_records_to_json(all_records)
    save_global_counts(gcounts)
    return gcounts

# ============================================================
# 页面
# ============================================================
def render_consent():
    st.markdown('<p class="center-title">文本审美对比实验</p>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("""
    <div style="font-size:1.1rem; line-height:2; color:#222; max-width:800px; margin:0 auto; padding:1rem 0;">
    您好！非常感谢您参与本次学术实验研究。
    <br><br>
    本次实验整体分为两个阶段，请您根据真实情况与直观感受完成即可。
    <br><br>
    第一阶段为基础个人信息填写环节，预计耗时3–4分钟。问卷收集的人口学、阅读习惯信息，仅用于学术研究数据分析，所有数据做匿名化处理，不会对外泄露个人隐私。
    <br><br>
    完成信息填写后将进入第二阶段的审美对比实验。正式实验的具体操作规则、实验要求与评判标准，将在后续页面为您详细说明，请您认真阅读并按照提示完成任务。
    <br><br>
    请您在安静、不受干扰的环境下认真完成全部实验内容，作答结果对本研究具有重要价值。再次衷心感谢您的参与与配合！
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    col = st.columns([1])
    if col[0].button("继续", use_container_width=True):
        st.session_state["phase"] = "questionnaire"
        st.rerun()

def render_questionnaire():
    st.markdown('<p class="center-title">第一部分：信息收集</p>', unsafe_allow_html=True)
    st.markdown("---")

    # 初始化专业详情错误标志
    if "major_detail_error" not in st.session_state:
        st.session_state["major_detail_error"] = False

    with st.form("qf"):
        name = st.text_input("姓名：", placeholder="")
        phone = st.text_input("手机号码：", placeholder="11位手机号")
        gender = st.radio("性别：", ["男", "女"], index=None, horizontal=True)
        age_raw = st.text_input("年龄（周岁）：", placeholder="请输入年龄", key="age_input")
        hukou = st.radio("户籍类型：", ["城市户籍", "农村户籍"], index=None, horizontal=True)
        edu = st.radio("当前在读教育阶段：", [
            "本科在读", "硕士研究生在读", "博士研究生及以上在读", "已毕业（非在读）"
        ], index=None, horizontal=False)

        major_cat = st.radio("所学专业大类：", [
            "人文类（中文、外语、历史、哲学等）",
            "社会科学类（心理学、法学、社会学等）",
            "理工科", "经管类", "艺术类", "医学类", "农学类", "其他",
        ], index=None, horizontal=False)

        # 其他专业时立即显示填空框
        major_detail_val = None
        if major_cat == "其他":
            major_detail_val = st.text_input("您的具体专业是：", placeholder="")

        llm_freq = st.radio("使用大语言模型（ChatGPT、豆包等）的频率：", [
            "几乎不用", "每月数次（1-4次）", "每周1-2次", "每周3-5次", "几乎每天都用",
        ], index=None, horizontal=False)

        ai_read = st.radio("是否读过AI生成的文学内容：", [
            "从未读过", "偶尔读过1-3次", "经常主动阅读",
        ], index=None, horizontal=True)

        read_freq = st.radio("平时阅读文学作品的频率：", [
            "几乎不读", "每月1-2本/篇", "每周1-2次", "几乎每天",
        ], index=None, horizontal=True)

        read_dim = st.radio("阅读时最看重哪个维度：", [
            "文字文笔、语言美感", "故事情节、叙事逻辑", "思想内涵、情感感染力",
            "想象力、创意设定", "人物塑造",
        ], index=None, horizontal=False)

        aesthetic_pref = st.radio("阅读偏好：", [
            "传统经典文学风格", "现代通俗网络文学风格", "两者都可以接受", "无明确偏好",
        ], index=None, horizontal=True)

        st.markdown("---")

        # 如果曾经选"其他"但没填详情，显示错误提示
        if st.session_state["major_detail_error"]:
            st.error("请填写具体专业")

        sub = st.form_submit_button("进入正式实验", use_container_width=True)

        if sub:
            # 如果选了"其他"但没填专业详情，不让进入，记录错误
            if major_cat == "其他" and (not major_detail_val or not major_detail_val.strip()):
                st.session_state["major_detail_error"] = True
                st.rerun()
                return

            # 填了详情就清除错误标志
            st.session_state["major_detail_error"] = False

            if not name:
                st.error("请填写姓名"); return
            if not phone or len(phone) != 11 or not phone.isdigit():
                st.error("请填写正确的11位手机号"); return
            if not all([gender, hukou, edu, major_cat, llm_freq, ai_read, read_freq, read_dim, aesthetic_pref]):
                st.error("请填写所有必填项"); return

            st.session_state["subject_name"] = name
            # 验证年龄为有效数字
            try:
                age = int(age_raw)
                if age < 1 or age > 120:
                    st.error("请填写正确的年龄"); return
            except ValueError:
                st.error("请填写正确的年龄（数字）"); return

            st.session_state["answers"] = {
                "phone": phone, "gender": gender, "age": age, "hukou": hukou,
                "edu": edu, "major_cat": major_cat,
                "major_detail": major_detail_val,
                "llm_freq": llm_freq,
                "ai_read": ai_read, "read_freq": read_freq,
                "read_dim": read_dim, "aesthetic_pref": aesthetic_pref,
            }
            st.session_state["phase"] = "instruction"
            st.rerun()

def render_instruction():
    texts_db = load_texts()
    st.markdown('<p class="center-title">文本审美对比实验说明</p>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown(f"""
    <div style="font-size:1.1rem; line-height:2; color:#222; max-width:800px; margin:0 auto; padding:0.5rem 0; text-align:center;">
    欢迎来到文本审美对比实验。在接下来的任务中，屏幕会依次呈现共 50 对文本段落，每段文本长度在 150‑250 字左右。
    <br><br>
    每一轮界面会同时展示左右两段不同的文本，请您结合自己的主观感受，对比两段文本，选出您认为综合表现更好的一段进行按键操作：
    如果您认为"左侧文本A更好"，请点击左侧按键；如果您认为"右侧文本B更好"，请点击右侧按键。
    <br><br>
    本次实验没有规定统一的评判标准，你可以综合从辞藻、语句通顺度、情感表达、意象氛围、叙事感受等任意你看重的维度，做出整体综合判断。本任务不存在标准答案，没有正确或错误之分，完全依据您个人的阅读感受做出选择即可。
    <br><br>
    实验大概需要15-20分钟。所有实验数据仅用于学术研究，全部结果将做匿名保密处理，不会泄露您的个人信息。完成一次选择点击按键后，页面会自动跳转到下一组文本对。
    <br><br>
    准备好之后，请按 "继续" 键开始实验。
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    col = st.columns([1])
    if col[0].button("继续", use_container_width=True):
        trials = generate_trials(texts_db, n_trials=TRIALS_PER_SUBJECT, global_counts=load_global_counts())
        st.session_state["trials"] = trials
        st.session_state["responses"] = []
        st.session_state["trial_index"] = 0
        st.session_state["exp_start_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st.session_state["phase"] = "experiment"
        st.rerun()

def render_experiment():
    texts_db = load_texts()
    idx = st.session_state["trial_index"]
    trials = st.session_state["trials"]

    if idx >= len(trials):
        st.session_state["phase"] = "thanks"
        st.rerun()
        return

    trial = trials[idx]
    progress = idx / len(trials)
    st.progress(progress, text=f"进度：{idx + 1} / {len(trials)}")

    left_id = trial["left_id"]
    right_id = trial["right_id"]
    left_text = texts_db[str(left_id)]["text"]
    right_text = texts_db[str(right_id)]["text"]

    st.markdown("""<p style="text-align:center; font-size:1.15rem; color:#1a1a1a; margin-bottom:1rem;">请阅读以下两段文本，判断哪一段更美</p>""", unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    # 左文本框 - 蓝色边框样式，统一大尺寸
    with c1:
        st.markdown(f"""
        <div style="background-color:{WHITE}; border:2px solid {BLUE}; border-radius:8px;
                    padding:1.5rem; height:340px; display:flex; align-items:center;
                    justify-content:center; text-align:center; font-size:1.35rem;
                    line-height:2; color:#1a1a1a; word-break:break-word;
                    overflow:hidden; box-sizing:border-box;">
            {left_text}
        </div>
        """, unsafe_allow_html=True)
        st.markdown("""<p style="text-align:center; color:#888; font-size:0.9rem; margin-top:0.3rem;">文本 A</p>""", unsafe_allow_html=True)

    # 右文本框 - 统一大尺寸
    with c2:
        st.markdown(f"""
        <div style="background-color:{WHITE}; border:2px solid {BLUE}; border-radius:8px;
                    padding:1.5rem; height:340px; display:flex; align-items:center;
                    justify-content:center; text-align:center; font-size:1.35rem;
                    line-height:2; color:#1a1a1a; word-break:break-word;
                    overflow:hidden; box-sizing:border-box;">
            {right_text}
        </div>
        """, unsafe_allow_html=True)
        st.markdown("""<p style="text-align:center; color:#888; font-size:0.9rem; margin-top:0.3rem;">文本 B</p>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""<p style="text-align:center; font-size:1.05rem; color:#1a1a1a; margin-top:1.2rem;">您认为哪段文本更美？</p>""", unsafe_allow_html=True)

    bc1, bc2 = st.columns([1, 1])
    with bc1:
        f_ok = st.button("← A更好", key=f"f_{idx}", use_container_width=True)
    with bc2:
        j_ok = st.button("B更好 →", key=f"j_{idx}", use_container_width=True)

    if f_ok or j_ok:
        choice = "left" if f_ok else "right"
        st.session_state["responses"].append({
            "choice": choice,
            "left_id": left_id,
            "right_id": right_id,
            "left_level": trial["left_level"],
            "right_level": trial["right_level"],
            "key": "F" if f_ok else "J",
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        st.session_state["trial_index"] += 1
        st.rerun()

def render_thanks():
    st.markdown('<p class="center-title">实验完成</p>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("""
    <div style="font-size:1.1rem; line-height:2; color:#222; max-width:700px; margin:0 auto; padding:1rem; text-align:center;">
    非常感谢您抽出宝贵时间参与本次实验！
    <br><br>
    您本次的全部作答数据将会进行匿名化处理，仅用于学术研究分析，不会泄露任何个人相关信息。
    <br><br>
    实验到此已经全部结束，祝您生活愉快！
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    if not st.session_state.get("submitted", False):
        try:
            name = st.session_state.get("subject_name", "未知被试")
            answers = st.session_state.get("answers", {})
            resp = st.session_state.get("responses", [])
            trials = st.session_state.get("trials", [])
            texts_db = load_texts()
            write_results(name, answers, resp, trials, texts_db)
            st.session_state["submitted"] = True
            st.success(f"数据已保存（被试：{name}，共{len(resp)}条记录）", icon=None)
        except Exception as e:
            st.error(f"保存失败：{e}")

    if os.path.exists(JSON_PATH):
        with open(JSON_PATH, "rb") as f:
            st.download_button(
                label="下载全部结果（JSON）",
                data=f,
                file_name="文本对比实验结果.json",
                mime="application/json"
            )

# ============================================================
# 主程序
# ============================================================
def main():
    st.set_page_config(**PAGE_CONFIG)
    init_session_state()

    phase = st.session_state["phase"]
    if phase == "consent":
        render_consent()
    elif phase == "questionnaire":
        render_questionnaire()
    elif phase == "instruction":
        render_instruction()
    elif phase == "experiment":
        render_experiment()
    elif phase == "thanks":
        render_thanks()

if __name__ == "__main__":
    main()
