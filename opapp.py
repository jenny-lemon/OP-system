"""
opapp.py  ──  營運報表控制台 v2
Enhanced with:
  • 排程執行狀態主控表（每日 / 每月）
  • 輸出報表檔案瀏覽（資料夾 + 狀態）
  • 線上新增 / 修改 / 刪除腳本設定
  • 手動觸發執行
"""

from pathlib import Path
import subprocess
import sys
import json
import os
from datetime import datetime, date, timedelta

import streamlit as st

# ──────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="營運報表控制台",
    page_icon="📊",
    layout="wide",
)

# ──────────────────────────────────────────────
# Paths
# ──────────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
LOG_DIR    = BASE_DIR / "logs"
OUTPUT_DIR = BASE_DIR / "output"        # 輸出報表根目錄
CONFIG_F   = BASE_DIR / "schedule_config.json"
RUNLOG_F   = BASE_DIR / "run_log.json"  # 執行歷史記錄

LOG_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

PYTHON_BIN = sys.executable or "python3"

# ──────────────────────────────────────────────
# Default schedule config
# ──────────────────────────────────────────────
DEFAULT_CONFIG = {
    "daily": [
        {"id": "d1", "label": "排班統計表",   "script": "排班統計表.py",    "args": [],      "schedule": "01:00"},
        {"id": "d2", "label": "專員班表",     "script": "專員班表.py",      "args": [],      "schedule": "02:00"},
        {"id": "d3", "label": "專員個資",     "script": "專員系統個資.py",  "args": [],      "schedule": "02:30"},
        {"id": "d4", "label": "當月次月訂單", "script": "當月次月訂單.py",  "args": [],      "schedule": "08:00"},
        {"id": "d5", "label": "業績報表 08",  "script": "業績報表.py",      "args": ["0800"],"schedule": "08:00"},
        {"id": "d6", "label": "業績報表 18",  "script": "業績報表.py",      "args": ["1800"],"schedule": "18:00"},
    ],
    "monthly": [
        {"id": "m1", "label": "上半月訂單",   "script": "上下半月訂單.py",  "args": ["1"],   "schedule": "月初 01 日"},
        {"id": "m2", "label": "下半月訂單",   "script": "上下半月訂單.py",  "args": ["2"],   "schedule": "月中 16 日"},
        {"id": "m3", "label": "已退款",       "script": "已退款.py",        "args": [],      "schedule": "月底"},
        {"id": "m4", "label": "預收",         "script": "預收.py",          "args": [],      "schedule": "月底"},
        {"id": "m5", "label": "儲值金結算",   "script": "儲值金結算.py",    "args": [],      "schedule": "月底"},
        {"id": "m6", "label": "儲值金預收",   "script": "儲值金預收.py",    "args": [],      "schedule": "月底"},
    ],
}

# ──────────────────────────────────────────────
# Config helpers
# ──────────────────────────────────────────────
def load_config() -> dict:
    if CONFIG_F.exists():
        try:
            return json.loads(CONFIG_F.read_text(encoding="utf-8"))
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    CONFIG_F.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

# ──────────────────────────────────────────────
# Run-log helpers
# ──────────────────────────────────────────────
def load_runlog() -> dict:
    if RUNLOG_F.exists():
        try:
            return json.loads(RUNLOG_F.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def save_runlog(log: dict):
    RUNLOG_F.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")

def record_run(job_id: str, ok: bool, stdout: str, stderr: str):
    log = load_runlog()
    log[job_id] = {
        "last_run": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ok": ok,
        "stdout": stdout[-800:] if stdout else "",
        "stderr": stderr[-800:] if stderr else "",
    }
    save_runlog(log)

# ──────────────────────────────────────────────
# Script runner
# ──────────────────────────────────────────────
def run_script(script_name: str, args=None) -> dict:
    args = args or []
    path = BASE_DIR / script_name
    if not path.exists():
        return {"ok": False, "stdout": "", "stderr": f"找不到：{path}", "cmd": ""}
    cmd = [PYTHON_BIN, str(path)] + [str(a) for a in args]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=str(BASE_DIR))
        return {"ok": r.returncode == 0, "stdout": r.stdout, "stderr": r.stderr, "cmd": " ".join(cmd)}
    except Exception as e:
        return {"ok": False, "stdout": "", "stderr": str(e), "cmd": " ".join(cmd)}

# ──────────────────────────────────────────────
# Unique ID generator
# ──────────────────────────────────────────────
def new_id(prefix="x") -> str:
    return f"{prefix}{datetime.now().strftime('%H%M%S%f')[-8:]}"

# ──────────────────────────────────────────────
# Output file scanner
# ──────────────────────────────────────────────
def scan_output_dir() -> list[dict]:
    """回傳 OUTPUT_DIR 及子目錄中的所有檔案，附帶 mtime / size。"""
    results = []
    for p in sorted(OUTPUT_DIR.rglob("*")):
        if p.is_file():
            stat = p.stat()
            results.append({
                "path": str(p),
                "name": p.name,
                "rel":  str(p.relative_to(OUTPUT_DIR)),
                "folder": str(p.parent.relative_to(OUTPUT_DIR)) if p.parent != OUTPUT_DIR else "（根目錄）",
                "size_kb": round(stat.st_size / 1024, 1),
                "mtime": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M"),
            })
    return results

# ──────────────────────────────────────────────
# CSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600;800&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Noto Sans TC', sans-serif;
}

/* Background */
.stApp {
    background: #0f1117;
    color: #e8eaf0;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #161b27;
    border-right: 1px solid #252d3d;
}
section[data-testid="stSidebar"] * { color: #c9d1e0 !important; }

/* Block container */
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}

/* ── Cards ── */
.card {
    background: #161b27;
    border: 1px solid #252d3d;
    border-radius: 14px;
    padding: 20px 22px 16px;
    margin-bottom: 18px;
}
.card-title {
    font-size: 17px;
    font-weight: 800;
    margin-bottom: 14px;
    letter-spacing: .02em;
    color: #a5b4fc;
    text-transform: uppercase;
    display: flex;
    align-items: center;
    gap: 8px;
}

/* ── Page header ── */
.page-header {
    display: flex;
    align-items: baseline;
    gap: 14px;
    margin-bottom: 6px;
}
.page-title {
    font-size: 28px;
    font-weight: 800;
    color: #e8eaf0;
    letter-spacing: -.01em;
}
.page-sub {
    color: #6b7280;
    font-size: 14px;
}

/* ── Status badges ── */
.badge {
    display: inline-block;
    padding: 2px 9px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 700;
    font-family: 'JetBrains Mono', monospace;
}
.badge-ok   { background:#14532d; color:#86efac; border:1px solid #166534; }
.badge-fail { background:#450a0a; color:#fca5a5; border:1px solid #7f1d1d; }
.badge-wait { background:#1c1917; color:#a8a29e; border:1px solid #292524; }

/* ── Table ── */
.sched-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13.5px;
}
.sched-table th {
    background: #1e2536;
    color: #a5b4fc;
    font-weight: 700;
    padding: 8px 12px;
    text-align: left;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: .06em;
}
.sched-table td {
    padding: 9px 12px;
    border-bottom: 1px solid #1e2536;
    color: #d1d5db;
    vertical-align: middle;
}
.sched-table tr:last-child td { border-bottom: none; }
.sched-table tr:hover td { background: #1a2035; }

/* ── Buttons ── */
.stButton > button {
    background: #312e81;
    color: #c7d2fe;
    border: 1px solid #3730a3;
    border-radius: 8px;
    font-weight: 700;
    font-size: 13px;
    padding: 6px 14px;
    width: 100%;
    transition: all .15s;
}
.stButton > button:hover {
    background: #4338ca;
    color: #fff;
    border-color: #6366f1;
}

/* ── Text inputs ── */
.stTextInput input, .stTextArea textarea, .stSelectbox select {
    background: #1e2536 !important;
    color: #e8eaf0 !important;
    border: 1px solid #252d3d !important;
    border-radius: 8px !important;
}

/* ── Expander ── */
details summary {
    font-size: 13px;
    color: #6b7280;
    cursor: pointer;
}

/* ── Code ── */
code, .stCode {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 12px;
}

/* ── Divider ── */
hr { border-color: #252d3d; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# Header
# ──────────────────────────────────────────────
st.markdown(f"""
<div class="page-header">
  <div class="page-title">📊 營運報表控制台</div>
  <div class="page-sub">專案目錄：{BASE_DIR} &nbsp;｜&nbsp; {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
</div>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# Sidebar navigation
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🗂️ 功能選單")
    page = st.radio(
        "",
        ["📋 排程主控表", "▶️ 手動執行", "📁 輸出報表", "⚙️ 腳本管理"],
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.caption("v2.0 · lemon ops")

cfg     = load_config()
runlog  = load_runlog()

# ══════════════════════════════════════════════════════════════
# PAGE 1 ── 排程主控表
# ══════════════════════════════════════════════════════════════
if page == "📋 排程主控表":

    def sched_table(jobs: list, title: str):
        st.markdown(f'<div class="card"><div class="card-title">{title}</div>', unsafe_allow_html=True)
        rows = ""
        for j in jobs:
            entry  = runlog.get(j["id"], {})
            last   = entry.get("last_run", "—")
            ok     = entry.get("ok", None)
            script_exists = (BASE_DIR / j["script"]).exists()

            if ok is True:
                badge = '<span class="badge badge-ok">✓ 成功</span>'
            elif ok is False:
                badge = '<span class="badge badge-fail">✗ 失敗</span>'
            else:
                badge = '<span class="badge badge-wait">— 尚未執行</span>'

            args_str = " ".join(j.get("args", [])) or "—"
            exist_icon = "🟢" if script_exists else "🔴"

            rows += f"""
            <tr>
              <td><b>{j['label']}</b></td>
              <td style="font-family:monospace;color:#a5b4fc">{j['schedule']}</td>
              <td>{exist_icon} {j['script']}</td>
              <td style="color:#78716c">{args_str}</td>
              <td>{badge}</td>
              <td style="font-family:monospace;font-size:12px;color:#6b7280">{last}</td>
            </tr>"""

        st.markdown(f"""
        <table class="sched-table">
          <thead><tr>
            <th>名稱</th><th>排程時間</th><th>腳本</th><th>參數</th><th>狀態</th><th>最後執行</th>
          </tr></thead>
          <tbody>{rows}</tbody>
        </table>
        </div>""", unsafe_allow_html=True)

    sched_table(cfg["daily"],   "📅 每日排程")
    sched_table(cfg["monthly"], "🗓️ 月排程")

    # 上次失敗詳情
    failures = [(jid, v) for jid, v in runlog.items() if not v.get("ok", True)]
    if failures:
        st.markdown('<div class="card"><div class="card-title">⚠️ 失敗詳情</div>', unsafe_allow_html=True)
        for jid, v in failures:
            label = next(
                (j["label"] for s in ["daily","monthly"] for j in cfg[s] if j["id"] == jid),
                jid
            )
            with st.expander(f"❌ {label} — {v['last_run']}"):
                if v["stderr"]:
                    st.code(v["stderr"], language="bash")
                if v["stdout"]:
                    st.code(v["stdout"])
        st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🔄 重新整理"):
        st.rerun()


# ══════════════════════════════════════════════════════════════
# PAGE 2 ── 手動執行
# ══════════════════════════════════════════════════════════════
elif page == "▶️ 手動執行":

    result_area = st.empty()

    def run_and_show(job: dict):
        with st.spinner(f"執行 {job['label']} …"):
            res = run_script(job["script"], job.get("args", []))
        record_run(job["id"], res["ok"], res["stdout"], res["stderr"])
        with result_area.container():
            if res["ok"]:
                st.success(f"✅ {job['label']} 完成")
            else:
                st.error(f"❌ {job['label']} 失敗")
            st.code(res["cmd"], language="bash")
            if res["stdout"].strip():
                st.code(res["stdout"])
            if res["stderr"].strip():
                with st.expander("stderr"):
                    st.code(res["stderr"])

    # ── Daily ──
    st.markdown('<div class="card"><div class="card-title">📅 每日報表</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, job in enumerate(cfg["daily"]):
        with cols[i % 3]:
            if st.button(job["label"], key=f"run_{job['id']}"):
                run_and_show(job)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Monthly ──
    st.markdown('<div class="card"><div class="card-title">🗓️ 月報表</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, job in enumerate(cfg["monthly"]):
        with cols[i % 3]:
            if st.button(job["label"], key=f"run_{job['id']}"):
                run_and_show(job)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Batch ──
    st.markdown('<div class="card"><div class="card-title">🚀 批次執行</div>', unsafe_allow_html=True)
    bc1, bc2 = st.columns(2)

    with bc1:
        if st.button("▶ 執行全部每日報表", key="batch_daily"):
            summary = []
            for job in cfg["daily"]:
                with st.spinner(f"執行 {job['label']} …"):
                    res = run_script(job["script"], job.get("args", []))
                record_run(job["id"], res["ok"], res["stdout"], res["stderr"])
                summary.append((job["label"], res["ok"]))
            with result_area.container():
                st.markdown("### 批次執行結果")
                for label, ok in summary:
                    if ok:
                        st.success(f"✅ {label}")
                    else:
                        st.error(f"❌ {label}")

    with bc2:
        if st.button("▶ 執行全部月報表", key="batch_monthly"):
            summary = []
            for job in cfg["monthly"]:
                with st.spinner(f"執行 {job['label']} …"):
                    res = run_script(job["script"], job.get("args", []))
                record_run(job["id"], res["ok"], res["stdout"], res["stderr"])
                summary.append((job["label"], res["ok"]))
            with result_area.container():
                st.markdown("### 批次執行結果")
                for label, ok in summary:
                    if ok:
                        st.success(f"✅ {label}")
                    else:
                        st.error(f"❌ {label}")

    st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# PAGE 3 ── 輸出報表
# ══════════════════════════════════════════════════════════════
elif page == "📁 輸出報表":

    files = scan_output_dir()

    # ── Summary metrics ──
    m1, m2, m3 = st.columns(3)
    m1.metric("📄 輸出檔案總數", len(files))
    total_kb = sum(f["size_kb"] for f in files)
    m2.metric("💾 總大小", f"{total_kb:.1f} KB")
    folders = {f["folder"] for f in files}
    m3.metric("📂 子資料夾數", len(folders))

    # ── Folder filter ──
    st.markdown('<div class="card"><div class="card-title">📂 輸出目錄瀏覽</div>', unsafe_allow_html=True)

    all_folders = sorted({"（全部）"} | folders)
    sel_folder  = st.selectbox("篩選資料夾", all_folders)

    filtered = files if sel_folder == "（全部）" else [f for f in files if f["folder"] == sel_folder]

    if not filtered:
        st.info(f"輸出目錄 `{OUTPUT_DIR}` 目前沒有檔案。執行報表後檔案將顯示於此。")
    else:
        rows = ""
        for f in sorted(filtered, key=lambda x: x["mtime"], reverse=True):
            rows += f"""
            <tr>
              <td>📄 {f['name']}</td>
              <td style="color:#6b7280">{f['folder']}</td>
              <td style="font-family:monospace;font-size:12px">{f['mtime']}</td>
              <td style="text-align:right;color:#a5b4fc">{f['size_kb']} KB</td>
            </tr>"""

        st.markdown(f"""
        <table class="sched-table">
          <thead><tr>
            <th>檔名</th><th>資料夾</th><th>修改時間</th><th style="text-align:right">大小</th>
          </tr></thead>
          <tbody>{rows}</tbody>
        </table>""", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Output dir path setting ──
    st.markdown('<div class="card"><div class="card-title">⚙️ 輸出目錄設定</div>', unsafe_allow_html=True)
    st.code(str(OUTPUT_DIR))
    st.caption("如需更改輸出目錄，請修改 opapp.py 中的 OUTPUT_DIR 變數。")
    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🔄 重新掃描"):
        st.rerun()


# ══════════════════════════════════════════════════════════════
# PAGE 4 ── 腳本管理
# ══════════════════════════════════════════════════════════════
elif page == "⚙️ 腳本管理":

    tab_daily, tab_monthly = st.tabs(["📅 每日排程腳本", "🗓️ 月排程腳本"])

    def script_editor(group_key: str, tab):
        with tab:
            jobs = cfg[group_key]

            # ── Existing jobs ──
            st.markdown('<div class="card"><div class="card-title">現有腳本</div>', unsafe_allow_html=True)

            for idx, job in enumerate(jobs):
                script_exists = (BASE_DIR / job["script"]).exists()
                exist_label   = "🟢 存在" if script_exists else "🔴 找不到檔案"

                with st.expander(f"{job['label']}  ·  {exist_label}  ·  `{job['script']}`"):
                    new_label    = st.text_input("顯示名稱",  job["label"],    key=f"lbl_{job['id']}")
                    new_script   = st.text_input("腳本檔名",  job["script"],   key=f"scr_{job['id']}")
                    new_args_str = st.text_input("執行參數（空格分隔）", " ".join(job.get("args", [])), key=f"arg_{job['id']}")
                    new_schedule = st.text_input("排程時間說明",          job["schedule"],   key=f"sch_{job['id']}")

                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button("💾 儲存修改", key=f"save_{job['id']}"):
                            cfg[group_key][idx]["label"]    = new_label
                            cfg[group_key][idx]["script"]   = new_script
                            cfg[group_key][idx]["args"]     = new_args_str.split() if new_args_str.strip() else []
                            cfg[group_key][idx]["schedule"] = new_schedule
                            save_config(cfg)
                            st.success("已儲存")
                            st.rerun()
                    with c2:
                        if st.button("🗑️ 刪除此項目", key=f"del_{job['id']}"):
                            cfg[group_key].pop(idx)
                            save_config(cfg)
                            st.warning(f"已刪除：{job['label']}")
                            st.rerun()

                    # Preview run
                    if st.button("▶ 測試執行", key=f"test_{job['id']}"):
                        with st.spinner("執行中…"):
                            res = run_script(job["script"], job.get("args", []))
                        record_run(job["id"], res["ok"], res["stdout"], res["stderr"])
                        if res["ok"]:
                            st.success("執行成功")
                        else:
                            st.error("執行失敗")
                        st.code(res.get("stderr") or res.get("stdout") or "（無輸出）")

            st.markdown("</div>", unsafe_allow_html=True)

            # ── Add new ──
            st.markdown('<div class="card"><div class="card-title">➕ 新增腳本</div>', unsafe_allow_html=True)

            with st.form(key=f"add_form_{group_key}"):
                a1, a2 = st.columns(2)
                with a1:
                    new_label    = st.text_input("顯示名稱",       placeholder="例：新報表")
                    new_script   = st.text_input("腳本檔名",       placeholder="例：新報表.py")
                with a2:
                    new_args_str = st.text_input("執行參數（空格分隔）", placeholder="例：0800")
                    new_schedule = st.text_input("排程時間說明",   placeholder="例：09:00")

                submitted = st.form_submit_button("➕ 新增")
                if submitted:
                    if not new_label or not new_script:
                        st.error("顯示名稱與腳本檔名為必填")
                    else:
                        prefix = group_key[0]
                        cfg[group_key].append({
                            "id":       new_id(prefix),
                            "label":    new_label,
                            "script":   new_script,
                            "args":     new_args_str.split() if new_args_str.strip() else [],
                            "schedule": new_schedule or "—",
                        })
                        save_config(cfg)
                        st.success(f"已新增：{new_label}")
                        st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)

            # ── Inline script editor ──
            st.markdown('<div class="card"><div class="card-title">📝 線上編輯腳本內容</div>', unsafe_allow_html=True)
            all_scripts = sorted({j["script"] for j in cfg[group_key]})
            sel_script  = st.selectbox("選擇腳本", ["（請選擇）"] + all_scripts, key=f"edit_sel_{group_key}")

            if sel_script and sel_script != "（請選擇）":
                spath = BASE_DIR / sel_script
                existing_code = spath.read_text(encoding="utf-8") if spath.exists() else "# 新檔案\n"
                edited = st.text_area("腳本內容", existing_code, height=350, key=f"editor_{group_key}")
                if st.button("💾 寫入腳本", key=f"write_{group_key}"):
                    spath.write_text(edited, encoding="utf-8")
                    st.success(f"已儲存 {sel_script}")

            st.markdown("</div>", unsafe_allow_html=True)

    script_editor("daily",   tab_daily)
    script_editor("monthly", tab_monthly)

    # ── Reset to defaults ──
    with st.expander("⚠️ 重置為預設設定"):
        if st.button("🔄 重置（會清除所有自訂設定）"):
            save_config(DEFAULT_CONFIG)
            st.warning("已重置為預設值，請重新整理頁面。")
            st.rerun()
