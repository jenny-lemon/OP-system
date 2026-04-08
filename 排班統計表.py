"""
opapp.py  ──  營運報表控制台 v3
Features:
  • 固定頂部導覽列（捲動後仍固定）
  • 地區切換，帳密從 accounts.py 讀取（本機，不上傳雲端）
  • 排程主控表：每日 / 月排程執行狀態一覽
  • 手動執行：單項 / 批次，帶入地區帳密
  • 輸出報表：掃描 output/ 目錄，顯示檔案狀態
  • 腳本管理：線上新增 / 修改 / 刪除 / 編輯腳本內容

帳密安全做法：
  accounts.py 加入 .gitignore → 永遠不上傳雲端
  opapp.py 只做 from accounts import ACCOUNTS，可以安全上傳
"""

from pathlib import Path
import subprocess
import sys
import json
import os
from datetime import datetime

import streamlit as st

# ── 帳密：從本機 accounts.py 讀取 ──────────────────────────────
try:
    from accounts import ACCOUNTS  # type: ignore
except ImportError:
    ACCOUNTS = {}   # accounts.py 不存在時不崩潰

# ──────────────────────────────────────────────
# Paths
# ──────────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
CONFIG_F   = BASE_DIR / "schedule_config.json"
RUNLOG_F   = BASE_DIR / "run_log.json"
OUTPUT_DIR.mkdir(exist_ok=True)

PYTHON_BIN = sys.executable or "python3"

# ──────────────────────────────────────────────
# Page config（必須是第一個 st 呼叫）
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="營運報表控制台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────
# CSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600;800&family=IBM+Plex+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Noto Sans TC', sans-serif !important;
    background: #0d1117 !important;
    color: #c9d1d9 !important;
}
#MainMenu, footer, header { visibility: hidden; }
section[data-testid="stSidebar"] { display: none !important; }

/* ── 固定頂部 NavBar ── */
.topnav {
    position: fixed;
    top: 0; left: 0; right: 0;
    z-index: 9999;
    background: #161b22;
    border-bottom: 1px solid #30363d;
    height: 52px;
    display: flex;
    align-items: center;
    padding: 0 24px;
    box-shadow: 0 2px 16px rgba(0,0,0,.5);
}
.nav-brand {
    font-size: 15px;
    font-weight: 800;
    color: #58a6ff;
    letter-spacing: .04em;
    margin-right: 28px;
    white-space: nowrap;
}
.nav-links { display: flex; gap: 2px; flex: 1; }
.nav-link {
    color: #8b949e;
    font-size: 13px;
    font-weight: 600;
    padding: 6px 14px;
    border-radius: 6px;
    white-space: nowrap;
}
.nav-link.active { background: #1f6feb22; color: #58a6ff; }
.nav-right {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 12px;
    font-size: 12px;
    color: #6e7681;
}
.region-chip {
    background: #1f6feb18;
    border: 1px solid #1f6feb44;
    color: #58a6ff;
    padding: 3px 12px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 700;
}
.region-chip.none {
    background: #21262d;
    border-color: #30363d;
    color: #6e7681;
}

/* ── Body offset ── */
.block-container {
    padding-top: 68px !important;
    padding-bottom: 3rem;
    max-width: 1400px;
}

/* ── Cards ── */
.card {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 12px;
    padding: 18px 22px 16px;
    margin-bottom: 16px;
}
.card-title {
    font-size: 11px;
    font-weight: 800;
    color: #58a6ff;
    text-transform: uppercase;
    letter-spacing: .12em;
    margin-bottom: 14px;
}

/* ── Table ── */
.sched-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.sched-table th {
    background: #1c2128;
    color: #6e7681;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .08em;
    padding: 8px 14px;
    text-align: left;
}
.sched-table td {
    padding: 10px 14px;
    border-bottom: 1px solid #21262d;
    vertical-align: middle;
}
.sched-table tr:last-child td { border-bottom: none; }
.sched-table tr:hover td { background: #1c2128; }

/* ── Badges ── */
.badge {
    display: inline-block;
    padding: 2px 9px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    font-family: 'IBM Plex Mono', monospace;
}
.badge-ok   { background:#0d4429; color:#3fb950; border:1px solid #196c2e; }
.badge-fail { background:#3d1a1a; color:#f85149; border:1px solid #6e2020; }
.badge-wait { background:#1c2128; color:#6e7681; border:1px solid #30363d; }

/* ── Buttons ── */
.stButton > button {
    background: #21262d !important;
    color: #c9d1d9 !important;
    border: 1px solid #30363d !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    width: 100%;
    transition: all .15s;
}
.stButton > button:hover {
    background: #1f6feb22 !important;
    color: #58a6ff !important;
    border-color: #1f6feb !important;
}

/* ── Metrics ── */
[data-testid="metric-container"] {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 10px;
    padding: 14px 18px;
}
[data-testid="stMetricLabel"] { color: #6e7681 !important; font-size: 12px !important; }
[data-testid="stMetricValue"] { color: #e6edf3 !important; font-size: 22px !important; font-weight: 800 !important; }

/* ── Inputs ── */
.stTextInput input, .stTextArea textarea {
    background: #1c2128 !important;
    color: #c9d1d9 !important;
    border: 1px solid #30363d !important;
    border-radius: 8px !important;
    font-size: 13px !important;
}
.stSelectbox > div > div {
    background: #1c2128 !important;
    border: 1px solid #30363d !important;
    color: #c9d1d9 !important;
}
label { color: #8b949e !important; font-size: 12px !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: transparent;
    border-bottom: 1px solid #30363d;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #6e7681 !important;
    font-size: 13px;
    font-weight: 700;
    padding: 8px 20px;
    border-radius: 0 !important;
}
.stTabs [aria-selected="true"] {
    color: #58a6ff !important;
    border-bottom: 2px solid #58a6ff !important;
}

/* ── Misc ── */
hr { border-color: #21262d !important; margin: 8px 0 16px !important; }
code {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important;
    background: #1c2128 !important;
    padding: 1px 5px !important;
    border-radius: 4px !important;
}
details summary { color: #6e7681; font-size: 13px; cursor: pointer; }
[data-testid="stAlert"] { border-radius: 8px !important; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# Session state defaults
# ──────────────────────────────────────────────
for k, v in {"page": "排程主控表", "region": None}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ──────────────────────────────────────────────
# Config / RunLog helpers
# ──────────────────────────────────────────────
DEFAULT_CONFIG = {
    "daily": [
        {"id": "d1", "label": "排班統計表",   "script": "排班統計表.py",   "args": [],       "schedule": "01:00", "all_regions": False},
        {"id": "d2", "label": "專員班表",     "script": "專員班表.py",     "args": [],       "schedule": "02:00", "all_regions": False},
        {"id": "d3", "label": "專員個資",     "script": "專員系統個資.py", "args": [],       "schedule": "02:30", "all_regions": False},
        {"id": "d4", "label": "當月次月訂單", "script": "當月次月訂單.py", "args": [],       "schedule": "08:00", "all_regions": True},
        {"id": "d5", "label": "業績報表 08",  "script": "業績報表.py",     "args": ["0800"], "schedule": "08:00", "all_regions": True},
        {"id": "d6", "label": "業績報表 18",  "script": "業績報表.py",     "args": ["1800"], "schedule": "18:00", "all_regions": True},
    ],
    "monthly": [
        {"id": "m1", "label": "上半月訂單",  "script": "上下半月訂單.py", "args": ["1"], "schedule": "月初01日", "all_regions": True},
        {"id": "m2", "label": "下半月訂單",  "script": "上下半月訂單.py", "args": ["2"], "schedule": "月中16日", "all_regions": True},
        {"id": "m3", "label": "已退款",      "script": "已退款.py",       "args": [],    "schedule": "月底",    "all_regions": True},
        {"id": "m4", "label": "預收",        "script": "預收.py",         "args": [],    "schedule": "月底",    "all_regions": False},
        {"id": "m5", "label": "儲值金結算",  "script": "儲值金結算.py",   "args": [],    "schedule": "月底",    "all_regions": False},
        {"id": "m6", "label": "儲值金預收",  "script": "儲值金預收.py",   "args": [],    "schedule": "月底",    "all_regions": False},
    ],
}

def load_config():
    if CONFIG_F.exists():
        try:
            return json.loads(CONFIG_F.read_text(encoding="utf-8"))
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    CONFIG_F.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

def load_runlog():
    if RUNLOG_F.exists():
        try:
            return json.loads(RUNLOG_F.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def save_runlog(log):
    RUNLOG_F.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")

def rkey(job_id, region=None):
    """執行記錄的唯一 key（job + 地區）。"""
    return f"{region}__{job_id}" if region else job_id

def record_run(key, ok, stdout, stderr):
    log = load_runlog()
    log[key] = {
        "last_run": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ok": ok,
        "stdout": (stdout or "")[-1000:],
        "stderr": (stderr or "")[-1000:],
    }
    save_runlog(log)

# ──────────────────────────────────────────────
# Script runner（帳密注入為環境變數傳入子程序）
# ──────────────────────────────────────────────
def run_script(script_name, args=None, region=None):
    args = args or []
    path = BASE_DIR / script_name
    if not path.exists():
        return {"ok": False, "stdout": "", "stderr": f"找不到腳本：{path}", "cmd": ""}

    env = os.environ.copy()
    if region and region in ACCOUNTS:
        acct = ACCOUNTS[region]
        env["REGION_EMAIL"]    = acct.get("email", "")
        env["REGION_PASSWORD"] = acct.get("password", "")
        env["REGION_NAME"]     = region

    cmd = [PYTHON_BIN, str(path)] + [str(a) for a in args]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=str(BASE_DIR), env=env)
        return {"ok": r.returncode == 0, "stdout": r.stdout, "stderr": r.stderr, "cmd": " ".join(cmd)}
    except Exception as e:
        return {"ok": False, "stdout": "", "stderr": str(e), "cmd": " ".join(cmd)}

def scan_output():
    files = []
    for p in sorted(OUTPUT_DIR.rglob("*")):
        if p.is_file():
            s = p.stat()
            files.append({
                "name":    p.name,
                "folder":  str(p.parent.relative_to(OUTPUT_DIR)) if p.parent != OUTPUT_DIR else "根目錄",
                "size_kb": round(s.st_size / 1024, 1),
                "mtime":   datetime.fromtimestamp(s.st_mtime).strftime("%Y-%m-%d %H:%M"),
            })
    return sorted(files, key=lambda x: x["mtime"], reverse=True)

def new_id(prefix="x"):
    return f"{prefix}{datetime.now().strftime('%f')}"

# ──────────────────────────────────────────────
# Load
# ──────────────────────────────────────────────
cfg    = load_config()
runlog = load_runlog()
REGIONS = sorted(ACCOUNTS.keys()) if ACCOUNTS else []

# ══════════════════════════════════════════════════════════════
# Fixed Top NavBar
# ══════════════════════════════════════════════════════════════
PAGES = ["排程主控表", "手動執行", "輸出報表", "腳本管理"]
cur   = st.session_state.page
rgn   = st.session_state.region

nav_html = "".join(
    f'<span class="nav-link{"  active" if p == cur else ""}">{p}</span>'
    for p in PAGES
)
chip_cls  = "region-chip" if rgn else "region-chip none"
chip_text = f"📍 {rgn}" if rgn else "📍 未選地區"

st.markdown(f"""
<div class="topnav">
  <div class="nav-brand">📊 營運報表</div>
  <div class="nav-links">{nav_html}</div>
  <div class="nav-right">
    <span class="{chip_cls}">{chip_text}</span>
    <span>{datetime.now().strftime('%m/%d %H:%M')}</span>
  </div>
</div>
""", unsafe_allow_html=True)

# 實際可點擊的導覽按鈕（排在 nav 下方，視覺上與 nav 對齊）
nav_cols = st.columns(len(PAGES) + 3)
for i, p in enumerate(PAGES):
    with nav_cols[i]:
        if st.button(p, key=f"nav_{p}"):
            st.session_state.page = p
            st.rerun()

st.markdown("<hr>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 共用：地區選擇器
# ══════════════════════════════════════════════════════════════
REGION_ALL = "【全部地區依序執行】"

def region_selector(allow_all=False):
    if not ACCOUNTS:
        st.error("⚠️ 找不到 accounts.py，請確認檔案位於專案目錄並包含 ACCOUNTS 字典。", icon="🚫")
        return None

    opts = ["（不指定地區）"]
    if allow_all:
        opts.append(REGION_ALL)
    opts += REGIONS

    cur_rgn = st.session_state.region
    idx = opts.index(cur_rgn) if cur_rgn in opts else 0

    col1, col2 = st.columns([2, 5])
    with col1:
        sel = st.selectbox("📍 操作地區", opts, index=idx, key="rgn_widget",
                           help="帳密由 accounts.py 讀取，以環境變數 REGION_EMAIL / REGION_PASSWORD 傳入腳本")
    with col2:
        if sel not in ("（不指定地區）", REGION_ALL):
            if sel in ACCOUNTS:
                st.info(f"✅ **{sel}** 帳號：{ACCOUNTS[sel].get('email','—')}　（密碼已遮蔽）", icon="🔑")
            else:
                st.warning(f"accounts.py 中找不到「{sel}」", icon="⚠️")
        elif sel == REGION_ALL:
            st.info(f"將依序執行所有地區：{', '.join(REGIONS)}", icon="🌐")

    new_val = None if sel == "（不指定地區）" else sel
    if new_val != st.session_state.region:
        st.session_state.region = new_val
        st.rerun()
    return new_val

# ══════════════════════════════════════════════════════════════
# PAGE 1 ── 排程主控表
# ══════════════════════════════════════════════════════════════
if st.session_state.page == "排程主控表":

    region = region_selector()

    all_jobs   = cfg["daily"] + cfg["monthly"]
    ok_count   = sum(1 for j in all_jobs if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is True)
    fail_count = sum(1 for j in all_jobs if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is False)
    wait_count = len(all_jobs) - ok_count - fail_count

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📋 排程總數", len(all_jobs))
    m2.metric("✅ 成功",    ok_count)
    m3.metric("❌ 失敗",    fail_count)
    m4.metric("⏳ 待執行",  wait_count)

    def render_table(jobs, title):
        st.markdown(f'<div class="card"><div class="card-title">{title}</div>', unsafe_allow_html=True)
        rows = ""
        for j in jobs:
            k      = rkey(j["id"], region)
            entry  = runlog.get(k) or runlog.get(j["id"], {})
            ok     = entry.get("ok", None)
            last   = entry.get("last_run", "—")
            exists = (BASE_DIR / j["script"]).exists()
            multi  = "🌐 全區" if j.get("all_regions") else "單區"

            badge = (
                '<span class="badge badge-ok">✓ 成功</span>'    if ok is True  else
                '<span class="badge badge-fail">✗ 失敗</span>'   if ok is False else
                '<span class="badge badge-wait">— 待執行</span>'
            )
            rows += f"""<tr>
              <td><b>{j['label']}</b></td>
              <td style="color:#8b949e;font-size:12px">{j['schedule']}</td>
              <td>{'🟢' if exists else '🔴'} <code>{j['script']}</code></td>
              <td style="color:#8b949e">{' '.join(j.get('args',[])) or '—'}</td>
              <td style="color:#8b949e;font-size:12px">{multi}</td>
              <td>{badge}</td>
              <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:#6e7681">{last}</td>
            </tr>"""

        st.markdown(f"""
        <table class="sched-table">
          <thead><tr>
            <th>名稱</th><th>排程</th><th>腳本</th><th>參數</th><th>地區</th><th>狀態</th><th>最後執行</th>
          </tr></thead>
          <tbody>{rows}</tbody>
        </table></div>""", unsafe_allow_html=True)

    render_table(cfg["daily"],   "📅 每日排程")
    render_table(cfg["monthly"], "🗓️ 月排程")

    # 失敗詳情
    failures = [
        (j, runlog.get(rkey(j["id"], region)) or runlog.get(j["id"]))
        for j in all_jobs
        if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is False
    ]
    if failures:
        st.markdown('<div class="card"><div class="card-title">⚠️ 失敗詳情</div>', unsafe_allow_html=True)
        for job, entry in failures:
            with st.expander(f"❌ {job['label']}  ·  {entry['last_run']}"):
                if entry.get("stderr"):
                    st.code(entry["stderr"], language="bash")
                if entry.get("stdout"):
                    st.code(entry["stdout"])
        st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🔄 重新整理"):
        st.rerun()


# ══════════════════════════════════════════════════════════════
# PAGE 2 ── 手動執行
# ══════════════════════════════════════════════════════════════
elif st.session_state.page == "手動執行":

    region = region_selector(allow_all=True)
    result_box = st.empty()

    def do_run(job, rgn):
        """執行一個 job，支援全區依序。回傳 [(地區, result), ...]。"""
        targets = REGIONS if rgn == REGION_ALL else [rgn]
        out = []
        for r in targets:
            res = run_script(job["script"], job.get("args", []), region=r)
            record_run(rkey(job["id"], r), res["ok"], res["stdout"], res["stderr"])
            out.append((r or "—", res))
        return out

    def show_results(label, pairs):
        with result_box.container():
            st.markdown(f"### {label}")
            for rgn_name, res in pairs:
                prefix = f"`{rgn_name}` " if rgn_name != "—" else ""
                if res["ok"]:
                    st.success(f"✅ {prefix}完成")
                else:
                    st.error(f"❌ {prefix}失敗")
                st.code(res["cmd"], language="bash")
                if res["stdout"].strip():
                    st.code(res["stdout"])
                if res["stderr"].strip():
                    with st.expander("stderr"):
                        st.code(res["stderr"])

    # ── 每日 ──
    st.markdown('<div class="card"><div class="card-title">📅 每日報表</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, job in enumerate(cfg["daily"]):
        with cols[i % 3]:
            if st.button(job["label"], key=f"run_{job['id']}"):
                with st.spinner(f"執行 {job['label']} …"):
                    show_results(job["label"], do_run(job, region))
    st.markdown("</div>", unsafe_allow_html=True)

    # ── 月報 ──
    st.markdown('<div class="card"><div class="card-title">🗓️ 月報表</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, job in enumerate(cfg["monthly"]):
        with cols[i % 3]:
            if st.button(job["label"], key=f"run_{job['id']}"):
                with st.spinner(f"執行 {job['label']} …"):
                    show_results(job["label"], do_run(job, region))
    st.markdown("</div>", unsafe_allow_html=True)

    # ── 批次 ──
    st.markdown('<div class="card"><div class="card-title">🚀 批次執行</div>', unsafe_allow_html=True)
    bc1, bc2 = st.columns(2)

    def batch_run(jobs, rgn):
        summary = []
        for job in jobs:
            with st.spinner(f"執行 {job['label']} …"):
                pairs = do_run(job, rgn)
            summary.append((job["label"], pairs))
        return summary

    def show_batch(title, summary):
        with result_box.container():
            st.markdown(f"### {title}")
            for label, pairs in summary:
                for rgn_name, res in pairs:
                    prefix = f"`{rgn_name}` " if rgn_name != "—" else ""
                    if res["ok"]:
                        st.success(f"✅ {prefix}{label}")
                    else:
                        st.error(f"❌ {prefix}{label}")
                        if res["stderr"]:
                            st.code(res["stderr"])

    with bc1:
        if st.button("▶ 全部每日報表", key="batch_d"):
            show_batch("每日批次結果", batch_run(cfg["daily"], region))

    with bc2:
        if st.button("▶ 全部月報表", key="batch_m"):
            show_batch("月報批次結果", batch_run(cfg["monthly"], region))

    st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# PAGE 3 ── 輸出報表
# ══════════════════════════════════════════════════════════════
elif st.session_state.page == "輸出報表":

    files   = scan_output()
    folders = sorted({f["folder"] for f in files})

    m1, m2, m3 = st.columns(3)
    m1.metric("📄 檔案總數", len(files))
    m2.metric("💾 總大小", f"{sum(f['size_kb'] for f in files):.1f} KB")
    m3.metric("📂 資料夾數", len(folders))

    st.markdown('<div class="card"><div class="card-title">📂 輸出目錄瀏覽</div>', unsafe_allow_html=True)
    st.caption(f"掃描路徑：`{OUTPUT_DIR}`")

    sel_folder = st.selectbox("篩選資料夾", ["（全部）"] + folders)
    filtered   = files if sel_folder == "（全部）" else [f for f in files if f["folder"] == sel_folder]

    if not filtered:
        st.info("output/ 目錄目前沒有檔案，執行報表後檔案將顯示於此。")
    else:
        rows = "".join(f"""<tr>
          <td>📄 <b>{f['name']}</b></td>
          <td style="color:#8b949e;font-size:12px">{f['folder']}</td>
          <td style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:#6e7681">{f['mtime']}</td>
          <td style="text-align:right;color:#58a6ff;font-family:'IBM Plex Mono',monospace;font-size:12px">{f['size_kb']} KB</td>
        </tr>""" for f in filtered)

        st.markdown(f"""
        <table class="sched-table">
          <thead><tr>
            <th>檔名</th><th>資料夾</th><th>修改時間</th><th style="text-align:right">大小</th>
          </tr></thead>
          <tbody>{rows}</tbody>
        </table>""", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🔄 重新掃描"):
        st.rerun()


# ══════════════════════════════════════════════════════════════
# PAGE 4 ── 腳本管理
# ══════════════════════════════════════════════════════════════
elif st.session_state.page == "腳本管理":

    tab_d, tab_m, tab_edit = st.tabs(["📅 每日排程腳本", "🗓️ 月排程腳本", "📝 線上編輯腳本"])

    def job_editor(group_key, tab):
        with tab:
            jobs = cfg[group_key]

            st.markdown('<div class="card"><div class="card-title">現有腳本</div>', unsafe_allow_html=True)
            for idx, job in enumerate(jobs):
                exists = (BASE_DIR / job["script"]).exists()
                icon   = "🟢" if exists else "🔴"
                with st.expander(f"{icon} {job['label']}  ·  `{job['script']}`"):
                    c1, c2 = st.columns(2)
                    with c1:
                        new_label  = st.text_input("顯示名稱",           job["label"],                  key=f"lbl_{job['id']}")
                        new_script = st.text_input("腳本檔名",            job["script"],                 key=f"scr_{job['id']}")
                    with c2:
                        new_args   = st.text_input("參數（空格分隔）",    " ".join(job.get("args",[])),  key=f"arg_{job['id']}")
                        new_sched  = st.text_input("排程說明",            job.get("schedule",""),        key=f"sch_{job['id']}")
                    new_multi = st.checkbox("需跑全部地區", value=job.get("all_regions", False), key=f"mul_{job['id']}")

                    a1, a2, a3 = st.columns(3)
                    with a1:
                        if st.button("💾 儲存", key=f"save_{job['id']}"):
                            cfg[group_key][idx].update({
                                "label": new_label, "script": new_script,
                                "args": new_args.split() if new_args.strip() else [],
                                "schedule": new_sched, "all_regions": new_multi,
                            })
                            save_config(cfg)
                            st.success("已儲存")
                            st.rerun()
                    with a2:
                        if st.button("▶ 測試", key=f"test_{job['id']}"):
                            r = run_script(job["script"], job.get("args",[]), region=st.session_state.region)
                            record_run(rkey(job["id"], st.session_state.region), r["ok"], r["stdout"], r["stderr"])
                            st.success("完成") if r["ok"] else st.error("失敗")
                            st.code(r.get("stderr") or r.get("stdout") or "（無輸出）")
                    with a3:
                        if st.button("🗑️ 刪除", key=f"del_{job['id']}"):
                            cfg[group_key].pop(idx)
                            save_config(cfg)
                            st.warning(f"已刪除：{job['label']}")
                            st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

            # ── 新增 ──
            st.markdown('<div class="card"><div class="card-title">➕ 新增腳本</div>', unsafe_allow_html=True)
            with st.form(key=f"add_{group_key}"):
                c1, c2 = st.columns(2)
                with c1:
                    f_label  = st.text_input("顯示名稱",        placeholder="例：新報表")
                    f_script = st.text_input("腳本檔名",        placeholder="例：新報表.py")
                with c2:
                    f_args   = st.text_input("參數（空格分隔）", placeholder="例：0800")
                    f_sched  = st.text_input("排程時間說明",     placeholder="例：09:00")
                f_multi = st.checkbox("需跑全部地區")
                if st.form_submit_button("➕ 新增"):
                    if not f_label or not f_script:
                        st.error("名稱與腳本為必填")
                    else:
                        cfg[group_key].append({
                            "id": new_id(group_key[0]),
                            "label": f_label, "script": f_script,
                            "args": f_args.split() if f_args.strip() else [],
                            "schedule": f_sched or "—",
                            "all_regions": f_multi,
                        })
                        save_config(cfg)
                        st.success(f"已新增：{f_label}")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    job_editor("daily",   tab_d)
    job_editor("monthly", tab_m)

    with tab_edit:
        st.markdown('<div class="card"><div class="card-title">📝 直接編輯腳本內容</div>', unsafe_allow_html=True)
        all_scripts = sorted({j["script"] for grp in ["daily","monthly"] for j in cfg[grp]})
        sel = st.selectbox("選擇腳本", ["（請選擇）"] + all_scripts)
        if sel and sel != "（請選擇）":
            spath  = BASE_DIR / sel
            code   = spath.read_text(encoding="utf-8") if spath.exists() else f"# {sel} 尚未建立\n"
            edited = st.text_area("腳本內容", code, height=420)
            if st.button("💾 寫入儲存"):
                spath.write_text(edited, encoding="utf-8")
                st.success(f"已寫入 {sel}")
        st.markdown("</div>", unsafe_allow_html=True)

        with st.expander("⚠️ 重置所有排程設定為預設值"):
            if st.button("🔄 確認重置"):
                save_config(DEFAULT_CONFIG)
                st.warning("已重置，請重新整理頁面")
                st.rerun()
