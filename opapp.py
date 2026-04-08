import os
import subprocess
from datetime import datetime

import streamlit as st

st.set_page_config(
    page_title="營運報表控制台",
    page_icon="📊",
    layout="wide",
)

# ===== 基本設定 =====
BASE_DIR = "/Users/jenny/lemon"
LOG_DIR = os.path.join(BASE_DIR, "logs")

os.makedirs(LOG_DIR, exist_ok=True)

# ===== 頁面樣式 =====
st.markdown("""
<style>
.main {
    background-color: #f7f9fc;
}
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
}
.title {
    font-size: 30px;
    font-weight: 800;
    margin-bottom: 4px;
}
.subtitle {
    color: #6b7280;
    margin-bottom: 18px;
}
.section-card {
    background: white;
    padding: 18px 18px 12px 18px;
    border-radius: 14px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    margin-bottom: 18px;
}
.section-title {
    font-size: 20px;
    font-weight: 700;
    margin-bottom: 10px;
}
.stButton > button {
    width: 100%;
    height: 44px;
    border-radius: 10px;
    border: 0;
    background: #4f46e5;
    color: white;
    font-weight: 700;
}
.stButton > button:hover {
    background: #4338ca;
    color: white;
}
.small-note {
    color: #6b7280;
    font-size: 13px;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="title">📊 營運報表控制台</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">手動執行日常報表、月報表、儲值金相關報表</div>', unsafe_allow_html=True)

# ===== 工具函式 =====
def run_python_script(script_name: str, args=None):
    if args is None:
        args = []

    script_path = os.path.join(BASE_DIR, script_name)

    if not os.path.exists(script_path):
        return {
            "ok": False,
            "stdout": "",
            "stderr": f"找不到檔案：{script_path}",
            "cmd": "",
        }

    cmd = ["python3", script_path] + [str(a) for a in args]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
        )
        return {
            "ok": result.returncode == 0,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "cmd": " ".join(cmd),
        }
    except Exception as e:
        return {
            "ok": False,
            "stdout": "",
            "stderr": str(e),
            "cmd": " ".join(cmd),
        }


def show_result(title: str, result: dict):
    st.markdown(f"### {title}")
    st.code(result.get("cmd", ""), language="bash")

    if result["ok"]:
        st.success("執行完成")
        if result["stdout"].strip():
            st.code(result["stdout"])
        else:
            st.info("沒有 stdout 輸出")
    else:
        st.error("執行失敗")
        if result["stderr"].strip():
            st.code(result["stderr"])
        elif result["stdout"].strip():
            st.code(result["stdout"])
        else:
            st.info("沒有錯誤輸出")


# ===== 共用執行區 =====
result_placeholder = st.empty()

# ===== 每日報表 =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">📅 每日報表</div>', unsafe_allow_html=True)
st.markdown('<div class="small-note">通常對應你原本 01:00 / 02:00 / 08:00 / 18:00 的排程</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)

with c1:
    if st.button("排班統計表"):
        result = run_python_script("排班統計表.py")
        with result_placeholder.container():
            show_result("排班統計表", result)

    if st.button("專員個資"):
        result = run_python_script("專員系統個資.py")
        with result_placeholder.container():
            show_result("專員個資", result)

with c2:
    if st.button("專員班表"):
        result = run_python_script("專員班表.py")
        with result_placeholder.container():
            show_result("專員班表", result)

    if st.button("當月次月訂單"):
        result = run_python_script("當月次月訂單.py")
        with result_placeholder.container():
            show_result("當月次月訂單", result)

with c3:
    if st.button("業績報表 08:00"):
        result = run_python_script("業績報表.py", ["0800"])
        with result_placeholder.container():
            show_result("業績報表 08:00", result)

    if st.button("業績報表 18:00"):
        result = run_python_script("業績報表.py", ["1800"])
        with result_placeholder.container():
            show_result("業績報表 18:00", result)

st.markdown('</div>', unsafe_allow_html=True)

# ===== 月報表 =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">🗓️ 月報表 / 區間報表</div>', unsafe_allow_html=True)
st.markdown('<div class="small-note">上下半月、退款、預收等</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)

with c1:
    if st.button("上下半月訂單（上半月）"):
        result = run_python_script("上下半月訂單.py", ["1"])
        with result_placeholder.container():
            show_result("上下半月訂單（上半月）", result)

with c2:
    if st.button("上下半月訂單（下半月）"):
        result = run_python_script("上下半月訂單.py", ["2"])
        with result_placeholder.container():
            show_result("上下半月訂單（下半月）", result)

with c3:
    if st.button("已退款"):
        result = run_python_script("已退款.py")
        with result_placeholder.container():
            show_result("已退款", result)

st.markdown('</div>', unsafe_allow_html=True)

# ===== 儲值金 =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">💰 儲值金</div>', unsafe_allow_html=True)
st.markdown('<div class="small-note">每月結算與預收資料</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)

with c1:
    if st.button("儲值金結算"):
        result = run_python_script("儲值金結算.py")
        with result_placeholder.container():
            show_result("儲值金結算", result)

with c2:
    if st.button("儲值金預收"):
        result = run_python_script("儲值金預收.py")
        with result_placeholder.container():
            show_result("儲值金預收", result)

with c3:
    if st.button("預收"):
        result = run_python_script("預收.py")
        with result_placeholder.container():
            show_result("預收", result)

st.markdown('</div>', unsafe_allow_html=True)

# ===== 一鍵執行 =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">🚀 一鍵執行</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)

with c1:
    if st.button("執行每日報表"):
        outputs = []
        jobs = [
            ("排班統計表", "排班統計表.py", []),
            ("專員班表", "專員班表.py", []),
            ("當月次月訂單", "當月次月訂單.py", []),
            ("專員個資", "專員系統個資.py", []),
            ("業績報表 08:00", "業績報表.py", ["0800"]),
            ("業績報表 18:00", "業績報表.py", ["1800"]),
        ]
        for label, script, args in jobs:
            outputs.append((label, run_python_script(script, args)))

        with result_placeholder.container():
            st.markdown("### 每日報表執行結果")
            for label, result in outputs:
                st.markdown(f"#### {label}")
                if result["ok"]:
                    st.success("完成")
                else:
                    st.error("失敗")
                if result["stdout"].strip():
                    st.code(result["stdout"])
                if result["stderr"].strip():
                    st.code(result["stderr"])

with c2:
    if st.button("執行月報表"):
        outputs = []
        jobs = [
            ("上下半月訂單（上半月）", "上下半月訂單.py", ["1"]),
            ("上下半月訂單（下半月）", "上下半月訂單.py", ["2"]),
            ("已退款", "已退款.py", []),
            ("預收", "預收.py", []),
            ("儲值金結算", "儲值金結算.py", []),
            ("儲值金預收", "儲值金預收.py", []),
        ]
        for label, script, args in jobs:
            outputs.append((label, run_python_script(script, args)))

        with result_placeholder.container():
            st.markdown("### 月報表執行結果")
            for label, result in outputs:
                st.markdown(f"#### {label}")
                if result["ok"]:
                    st.success("完成")
                else:
                    st.error("失敗")
                if result["stdout"].strip():
                    st.code(result["stdout"])
                if result["stderr"].strip():
                    st.code(result["stderr"])

st.markdown('</div>', unsafe_allow_html=True)

st.caption(f"目前時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
