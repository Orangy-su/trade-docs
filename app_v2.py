"""
app_v2.py — 外贸单据自动生成系统 Web版 v2
新增：主数据手册固化存储 + 公网部署支持
"""
import os, sys, json, uuid, threading, traceback, hashlib
from flask import Flask, request, jsonify, send_file, render_template_string
from datetime import datetime
from werkzeug.utils import secure_filename   

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
app.secret_key = os.environ.get('SECRET_KEY', 'qf-trade-system-2026')

UPLOAD_DIR  = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR  = os.path.join(BASE_DIR, 'output')
MASTER_DIR  = os.path.join(BASE_DIR, 'master_store')  # 固化存储目录
MASTER_META = os.path.join(MASTER_DIR, 'meta.json')

for d in [UPLOAD_DIR, OUTPUT_DIR, MASTER_DIR]:
    os.makedirs(d, exist_ok=True)

tasks = {}


def load_master_meta():
    if os.path.exists(MASTER_META):
        with open(MASTER_META) as f:
            return json.load(f)
    return None

def save_master_meta(meta):
    with open(MASTER_META, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


HTML_V2 = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>外贸单据系统 · 擎烽电气</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;700;900&family=JetBrains+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
:root{
  --navy:#1F3864;--blue:#2E75B6;--lblue:#D6E4F0;--teal:#0D7377;
  --orange:#E67E22;--green:#27AE60;--red:#C0392B;--purple:#6C3483;
  --bg:#EEF2F7;--card:#fff;--border:#DDE3EC;--text:#2C3E50;
  --gray:#7F8C8D;--lightgray:#F5F7FA;
  --mono:'JetBrains Mono',monospace;--sans:'Noto Sans SC',sans-serif;
  --radius:14px;--shadow:0 2px 16px rgba(31,56,100,0.08);
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh}

/* ── Topbar ── */
.topbar{
  background:linear-gradient(135deg,var(--navy) 0%,#2a4a82 100%);
  padding:0 2rem;height:60px;display:flex;align-items:center;
  justify-content:space-between;position:sticky;top:0;z-index:200;
  box-shadow:0 2px 20px rgba(31,56,100,0.25);
}
.brand{display:flex;align-items:center;gap:14px;text-decoration:none}
.brand-logo{
  width:38px;height:38px;background:var(--teal);border-radius:10px;
  display:flex;align-items:center;justify-content:center;font-size:20px;
  box-shadow:0 2px 8px rgba(0,0,0,0.2);
}
.brand-text .name{font-size:15px;font-weight:700;color:#fff;letter-spacing:.3px}
.brand-text .sub{font-size:11px;color:rgba(255,255,255,.6);margin-top:1px}
.topbar-right{display:flex;align-items:center;gap:12px}
.master-status{
  display:flex;align-items:center;gap:8px;
  background:rgba(255,255,255,.1);border-radius:20px;
  padding:5px 12px;font-size:12px;color:rgba(255,255,255,.85);cursor:pointer;
  border:1px solid rgba(255,255,255,.15);transition:.2s;
}
.master-status:hover{background:rgba(255,255,255,.18)}
.master-dot{
  width:7px;height:7px;border-radius:50%;
  background:#6E7681;flex-shrink:0;transition:.3s;
}
.master-dot.ok{background:#3FB950}
.master-dot.none{background:#D29922;animation:blink 1.5s infinite}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.4}}
.ver-badge{
  font-family:var(--mono);font-size:11px;color:rgba(255,255,255,.5);
  background:rgba(255,255,255,.07);padding:3px 8px;border-radius:4px;
}

/* ── Layout ── */
.main{max-width:1000px;margin:0 auto;padding:2rem 1.5rem}
@media(max-width:640px){.main{padding:1rem}}

/* ── Section header ── */
.section-hdr{
  display:flex;align-items:center;gap:10px;margin-bottom:1rem;
  padding-bottom:.75rem;border-bottom:2px solid var(--lblue);
}
.section-num{
  width:26px;height:26px;border-radius:50%;background:var(--blue);
  color:#fff;font-size:13px;font-weight:700;
  display:flex;align-items:center;justify-content:center;flex-shrink:0;
}
.section-title{font-size:16px;font-weight:700;color:var(--navy)}
.section-sub{font-size:12px;color:var(--gray);margin-left:auto}

/* ── Cards ── */
.card{
  background:var(--card);border-radius:var(--radius);
  border:1px solid var(--border);padding:1.5rem;
  margin-bottom:1.25rem;box-shadow:var(--shadow);
}
.card.master-panel{border-left:4px solid var(--teal)}
.card.upload-panel{border-left:4px solid var(--orange)}
.card.action-panel{border-left:4px solid var(--blue)}
.card.progress-panel{display:none}
.card.progress-panel.show{display:block}
.card.result-panel{display:none}
.card.result-panel.show{display:block}

/* ── Master固化状态 ── */
.master-info-box{
  display:flex;align-items:center;justify-content:space-between;
  padding:1rem 1.25rem;border-radius:10px;
  border:1.5px solid var(--border);flex-wrap:wrap;gap:12px;
}
.master-info-box.loaded{
  background:linear-gradient(135deg,#F0FBF4,#E8F7EE);
  border-color:#A3D9B1;
}
.master-info-box.empty{
  background:#FAFBFC;border-style:dashed;
}
.master-file-info{display:flex;align-items:center;gap:12px}
.master-file-icon{font-size:2rem}
.master-file-name{font-size:14px;font-weight:600;color:var(--navy)}
.master-file-meta{font-size:11px;color:var(--gray);margin-top:2px}
.master-actions{display:flex;gap:8px;flex-shrink:0}

/* ── Upload zone ── */
.upload-zone{
  border:2px dashed var(--border);border-radius:10px;padding:2rem;
  text-align:center;cursor:pointer;transition:.2s;background:#FAFCFF;
  position:relative;
}
.upload-zone:hover,.upload-zone.drag-over{border-color:var(--blue);background:#EBF3FB}
.upload-zone.has-file{border-color:var(--green);background:#F0FBF4;border-style:solid}
.upload-zone input[type=file]{
  position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;
}
.upload-icon{font-size:2.5rem;margin-bottom:.5rem}
.upload-label{font-size:14px;color:var(--gray)}
.upload-label strong{color:var(--blue)}
.upload-ok{font-size:14px;color:var(--green);font-weight:600;
  display:flex;align-items:center;justify-content:center;gap:8px}

/* ── Batch input ── */
.batch-row{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap}
.field-group{flex:1;min-width:220px}
.field-label{font-size:12px;font-weight:600;color:var(--navy);margin-bottom:6px;display:block}
.field-input{
  width:100%;padding:11px 14px;border:1.5px solid var(--border);
  border-radius:8px;font-family:var(--mono);font-size:13px;
  color:var(--text);background:#fff;outline:none;transition:.2s;
}
.field-input:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(46,117,182,.1)}
.field-hint{font-size:11px;color:var(--gray);margin-top:5px}

/* ── Buttons ── */
.btn{
  padding:10px 20px;border-radius:8px;font-family:var(--sans);
  font-size:14px;font-weight:600;cursor:pointer;border:none;
  transition:all .2s;display:inline-flex;align-items:center;gap:7px;
  white-space:nowrap;
}
.btn-primary{background:var(--blue);color:#fff}
.btn-primary:hover:not(:disabled){background:var(--navy);transform:translateY(-1px);box-shadow:0 4px 12px rgba(46,117,182,.3)}
.btn-primary:disabled{background:#B0C4DE;cursor:not-allowed}
.btn-teal{background:var(--teal);color:#fff}
.btn-teal:hover:not(:disabled){background:#0A5F63;transform:translateY(-1px)}
.btn-teal:disabled{background:#aaa;cursor:not-allowed}
.btn-outline{background:#fff;color:var(--blue);border:1.5px solid var(--blue)}
.btn-outline:hover{background:var(--lblue)}
.btn-outline-red{background:#fff;color:var(--red);border:1.5px solid var(--red)}
.btn-outline-red:hover{background:#FCE4D6}
.btn-green{background:var(--green);color:#fff}
.btn-green:hover{background:#219A52}
.btn-sm{padding:7px 14px;font-size:12px;border-radius:6px}
.btn-lg{padding:13px 30px;font-size:15px;border-radius:10px}
.btn-block{width:100%;justify-content:center}

/* ── Log terminal ── */
.terminal{
  background:#0D1117;border-radius:10px;padding:1rem 1.25rem;
  font-family:var(--mono);font-size:12px;line-height:1.9;
  max-height:320px;overflow-y:auto;color:#C9D1D9;
}
.log-line{display:flex;gap:10px;align-items:baseline}
.lt{color:#484F58;flex-shrink:0;font-size:11px}
.lo{color:#3FB950}.lw{color:#D29922}.le{color:#F85149}
.li{color:#58A6FF}.ld{color:#6E7681}
.status-row{display:flex;align-items:center;gap:10px;margin-bottom:1rem}
.sdot{
  width:10px;height:10px;border-radius:50%;
  background:#6E7681;flex-shrink:0;transition:.3s;
}
.sdot.run{background:var(--orange);animation:pulse 1s infinite}
.sdot.ok {background:var(--green)}
.sdot.err{background:var(--red)}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(.8)}}
.status-title{font-size:15px;font-weight:700;color:var(--navy)}
.status-sub{font-size:12px;color:var(--gray)}

/* ── Results ── */
.result-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:1rem}
@media(max-width:560px){.result-grid{grid-template-columns:1fr}}
.result-file{
  border:1.5px solid var(--border);border-radius:10px;padding:1.1rem;
  display:flex;flex-direction:column;gap:8px;transition:.2s;
}
.result-file:hover{box-shadow:0 4px 16px rgba(0,0,0,.08)}
.result-file.s1{border-color:var(--blue);background:linear-gradient(135deg,#F5FAFF,#EBF3FB)}
.result-file.s2{border-color:var(--green);background:linear-gradient(135deg,#F5FDF7,#EAFAF1)}
.rf-tag{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px}
.s1 .rf-tag{color:var(--blue)}.s2 .rf-tag{color:var(--green)}
.rf-name{font-size:12px;color:var(--text);font-family:var(--mono);word-break:break-all}
.rf-sheets{font-size:11px;color:var(--gray)}

/* ── History ── */
.hist-list{display:flex;flex-direction:column;gap:8px}
.hist-item{
  display:flex;align-items:center;justify-content:space-between;
  padding:10px 14px;background:var(--lightgray);border-radius:8px;
  border:1px solid var(--border);gap:8px;flex-wrap:wrap;
}
.hist-batch{font-size:13px;font-weight:600;font-family:var(--mono);color:var(--navy)}
.hist-time{font-size:11px;color:var(--gray)}
.hist-actions{display:flex;gap:6px}

/* ── Modal ── */
.modal-overlay{
  display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);
  z-index:300;align-items:center;justify-content:center;
}
.modal-overlay.show{display:flex}
.modal{
  background:#fff;border-radius:16px;padding:2rem;width:100%;
  max-width:520px;margin:1rem;box-shadow:0 20px 60px rgba(0,0,0,.2);
  animation:slideUp .25s ease;
}
@keyframes slideUp{from{transform:translateY(20px);opacity:0}to{transform:none;opacity:1}}
.modal-title{font-size:17px;font-weight:700;color:var(--navy);margin-bottom:.5rem}
.modal-sub{font-size:13px;color:var(--gray);margin-bottom:1.5rem}
.modal-footer{display:flex;gap:10px;justify-content:flex-end;margin-top:1.5rem}
.divider{height:1px;background:var(--border);margin:1.25rem 0}
</style>
</head>
<body>

<!-- 顶部栏 -->
<nav class="topbar">
  <a class="brand" href="/">
    <div class="brand-logo">📦</div>
    <div class="brand-text">
      <div class="name">外贸单据自动生成系统</div>
      <div class="sub">Trade Document Automation · 广东擎烽电气科技有限公司</div>
    </div>
  </a>
  <div class="topbar-right">
    <div class="master-status" onclick="openMasterModal()">
      <div class="master-dot none" id="masterDot"></div>
      <span id="masterStatusText">主数据手册未设置</span>
    </div>
    <span class="ver-badge">v2.0</span>
  </div>
</nav>

<main class="main">

  <!-- ─── Part 1: 主数据手册（固化） ─── -->
  <div class="card master-panel" id="masterCard">
    <div class="section-hdr">
      <div class="section-num" style="background:var(--teal)">●</div>
      <div class="section-title">主数据管理手册</div>
      <div class="section-sub">上传一次，长期有效 · 含成品档案/物料价格/往来方/出货批次记录</div>
    </div>

    <!-- 未设置状态 -->
    <div class="master-info-box empty" id="masterEmpty">
      <div class="master-file-info">
        <div class="master-file-icon">📋</div>
        <div>
          <div class="master-file-name" style="color:var(--gray)">尚未上传主数据手册</div>
          <div class="master-file-meta">上传后系统会记住，下次打开无需重新上传</div>
        </div>
      </div>
      <button class="btn btn-teal" onclick="openMasterModal()">⬆ 上传主数据手册</button>
    </div>

    <!-- 已设置状态 -->
    <div class="master-info-box loaded" id="masterLoaded" style="display:none">
      <div class="master-file-info">
        <div class="master-file-icon">✅</div>
        <div>
          <div class="master-file-name" id="masterFileName">-</div>
          <div class="master-file-meta" id="masterFileMeta">-</div>
        </div>
      </div>
      <div class="master-actions">
        <button class="btn btn-outline btn-sm" onclick="openMasterModal()">🔄 重新上传</button>
        <button class="btn btn-outline-red btn-sm" onclick="clearMaster()">✕ 清除</button>
      </div>
    </div>
  </div>

  <!-- ─── Part 2: 上传装箱清单 ─── -->
  <div class="card upload-panel">
    <div class="section-hdr">
      <div class="section-num">1</div>
      <div class="section-title">上传装箱清单</div>
      <div class="section-sub">仓管员填写后提交的当票装箱清单</div>
    </div>
    <div class="upload-zone" id="packingZone">
      <input type="file" id="packingFile" accept=".xlsx,.xls" onchange="handlePacking(this)">
      <div class="upload-icon">🗂️</div>
      <div class="upload-label">拖拽文件到此处，或 <strong>点击选择文件</strong></div>
      <div style="font-size:11px;color:var(--gray);margin-top:4px">支持 .xlsx · .xls</div>
    </div>
  </div>

  <!-- ─── Part 3: 批次号 + 生成 ─── -->
  <div class="card action-panel">
    <div class="section-hdr">
      <div class="section-num">2</div>
      <div class="section-title">填写批次号，生成单据</div>
      <div class="section-sub">对应⑤出货批次记录主表的「批次号」列</div>
    </div>

    <div class="batch-row" style="margin-bottom:1.25rem">
      <div class="field-group">
        <label class="field-label">批次号 Batch No.</label>
        <input type="text" class="field-input" id="batchNo"
          placeholder="如：US13-20260601-01"
          oninput="updateBtns()">
        <div class="field-hint">对应主数据手册⑤出货批次记录 → 主表 → 批次号列</div>
      </div>
    </div>

    <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:center">
      <button class="btn btn-primary btn-lg" id="genBtn" onclick="generate()" disabled>
        ⚡ 生成单据
      </button>
      <button class="btn btn-teal btn-lg" id="genAllBtn" onclick="generateAll()" disabled>
        📋 生成全部批次
      </button>
      <div style="font-size:12px;color:var(--gray);flex:1;min-width:160px">
        「全部批次」会处理主数据⑤里所有批次
      </div>
    </div>
  </div>

  <!-- ─── 进度区 ─── -->
  <div class="card progress-panel" id="progressCard">
    <div class="status-row">
      <div class="sdot run" id="sdot"></div>
      <div>
        <div class="status-title" id="progTitle">正在生成…</div>
        <div class="status-sub" id="progSub">请稍候，不要关闭页面</div>
      </div>
    </div>
    <div class="terminal" id="termLog"></div>
  </div>

  <!-- ─── 结果区 ─── -->
  <div class="card result-panel" id="resultCard">
    <div class="section-hdr">
      <span style="font-size:22px">🎉</span>
      <div>
        <div class="section-title">单据生成完成</div>
        <div class="section-sub" id="resultSub"></div>
      </div>
    </div>
    <div class="result-grid" id="resultGrid"></div>
  </div>

  <!-- ─── 历史记录 ─── -->
  <div class="card" id="histCard" style="display:none">
    <div class="section-hdr">
      <span style="font-size:18px">🕐</span>
      <div class="section-title">本次会话记录</div>
    </div>
    <div class="hist-list" id="histList"></div>
  </div>

</main>

<!-- ─── 主数据上传 Modal ─── -->
<div class="modal-overlay" id="masterModal">
  <div class="modal">
    <div class="modal-title">上传主数据管理手册</div>
    <div class="modal-sub">上传一次后系统会记住，后续打开页面无需重新上传。如果数据有调整，重新上传即可覆盖。</div>

    <div class="upload-zone" id="masterZone" style="padding:1.5rem">
      <input type="file" id="masterFile" accept=".xlsx,.xls" onchange="handleMaster(this)">
      <div class="upload-icon" id="masterZoneIcon">📄</div>
      <div id="masterZoneText">
        <div class="upload-label">拖拽主数据管理手册到此处<br>或 <strong>点击选择文件</strong></div>
        <div style="font-size:11px;color:var(--gray);margin-top:4px">外贸主数据管理手册_v3.1_单证员版.xlsx</div>
      </div>
    </div>

    <div class="modal-footer">
      <button class="btn btn-outline" onclick="closeMasterModal()">取消</button>
    </div>
  </div>
</div>

<script>
// ── 状态 ──────────────────────────────────────────────────────────────────
const S = {
  masterPath: null,
  masterName: null,
  masterMeta: null,
  packingPath: null,
  history: [],
};

// ── 初始化：检查服务端是否已有主数据 ─────────────────────────────────────
async function init() {
  try {
    const r = await fetch('/master/info');
    const d = await r.json();
    if (d.ok && d.meta) {
      S.masterPath = d.meta.path;
      S.masterName = d.meta.name;
      S.masterMeta = d.meta;
      updateMasterUI(true);
    }
  } catch(e) {}
}
init();

// ── 主数据上传 Modal ──────────────────────────────────────────────────────
function openMasterModal()  { document.getElementById('masterModal').classList.add('show'); }
function closeMasterModal() { document.getElementById('masterModal').classList.remove('show'); }
document.getElementById('masterModal').addEventListener('click', e => {
  if (e.target === e.currentTarget) closeMasterModal();
});

async function handleMaster(input) {
  const file = input.files[0];
  if (!file) return;
  const zone = document.getElementById('masterZone');
  document.getElementById('masterZoneIcon').textContent = '⏳';
  document.getElementById('masterZoneText').innerHTML = `<div class="upload-label">正在上传 ${file.name}…</div>`;

  const fd = new FormData();
  fd.append('file', file);
  try {
    const r = await fetch('/master/upload', { method:'POST', body:fd });
    const d = await r.json();
    if (d.ok) {
      S.masterPath = d.path;
      S.masterName = d.name;
      S.masterMeta = d;
      updateMasterUI(true);
      closeMasterModal();
    } else {
      document.getElementById('masterZoneIcon').textContent = '❌';
      document.getElementById('masterZoneText').innerHTML = `<div style="color:var(--red);font-size:13px">上传失败：${d.error}</div>`;
    }
  } catch(e) {
    document.getElementById('masterZoneIcon').textContent = '❌';
    document.getElementById('masterZoneText').innerHTML = `<div style="color:var(--red);font-size:13px">网络错误，请重试</div>`;
  }
}

function updateMasterUI(loaded) {
  const dot  = document.getElementById('masterDot');
  const txt  = document.getElementById('masterStatusText');
  const empty  = document.getElementById('masterEmpty');
  const loadedEl = document.getElementById('masterLoaded');
  if (loaded && S.masterMeta) {
    dot.className = 'master-dot ok';
    txt.textContent = S.masterName || '主数据已就绪';
    empty.style.display = 'none';
    loadedEl.style.display = 'flex';
    document.getElementById('masterFileName').textContent = S.masterName;
    document.getElementById('masterFileMeta').textContent =
      `上传时间：${S.masterMeta.uploaded_at || '--'}  ·  ${S.masterMeta.size || ''}`;
  } else {
    dot.className = 'master-dot none';
    txt.textContent = '主数据手册未设置';
    empty.style.display = 'flex';
    loadedEl.style.display = 'none';
  }
  updateBtns();
}

async function clearMaster() {
  if (!confirm('确认清除已保存的主数据手册吗？')) return;
  await fetch('/master/clear', { method:'POST' });
  S.masterPath = S.masterName = S.masterMeta = null;
  updateMasterUI(false);
}

// ── 装箱清单上传 ──────────────────────────────────────────────────────────
async function handlePacking(input) {
  const file = input.files[0];
  if (!file) return;
  const zone = document.getElementById('packingZone');
  zone.innerHTML = `<div style="font-size:13px;color:var(--gray)">⏳ 正在上传 ${file.name}…</div>
    <input type="file" accept=".xlsx,.xls" onchange="handlePacking(this)"
      style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">`;

  const fd = new FormData();
  fd.append('file', file);
  fd.append('type', 'packing');
  try {
    const r = await fetch('/upload', { method:'POST', body:fd });
    const d = await r.json();
    if (d.ok) {
      S.packingPath = d.path;
      zone.className = 'upload-zone has-file';
      zone.innerHTML = `
        <div class="upload-ok">✅ ${file.name}</div>
        <div style="font-size:11px;color:var(--gray);margin-top:4px">${d.size}</div>
        <input type="file" accept=".xlsx,.xls" onchange="handlePacking(this)"
          style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">`;
    } else {
      zone.innerHTML = `<div style="color:var(--red);font-size:13px">❌ ${d.error}</div>
        <input type="file" accept=".xlsx,.xls" onchange="handlePacking(this)"
          style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">`;
    }
  } catch(e) {
    zone.innerHTML = `<div style="color:var(--red);font-size:13px">❌ 网络错误</div>
      <input type="file" accept=".xlsx,.xls" onchange="handlePacking(this)"
        style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">`;
  }
  updateBtns();
}

// ── 按钮状态 ──────────────────────────────────────────────────────────────
function updateBtns() {
  const ready = S.masterPath && S.packingPath;
  const hasBatch = document.getElementById('batchNo').value.trim();
  document.getElementById('genBtn').disabled    = !(ready && hasBatch);
  document.getElementById('genAllBtn').disabled = !ready;
}

// ── 生成 ──────────────────────────────────────────────────────────────────
async function generate()    { await startGen(document.getElementById('batchNo').value.trim(), false); }
async function generateAll() { await startGen('', true); }

async function startGen(batch, all) {
  const prog = document.getElementById('progressCard');
  const res  = document.getElementById('resultCard');
  prog.classList.add('show');
  res.classList.remove('show');
  document.getElementById('termLog').innerHTML = '';
  document.getElementById('sdot').className = 'sdot run';
  document.getElementById('progTitle').textContent = all ? '正在批量生成…' : `正在生成：${batch}`;
  document.getElementById('progSub').textContent = '请稍候，不要关闭页面';
  document.getElementById('genBtn').disabled = true;
  document.getElementById('genAllBtn').disabled = true;
  prog.scrollIntoView({behavior:'smooth',block:'nearest'});

  try {
    const r = await fetch('/generate', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({ packing_path:S.packingPath, batch_no:batch, all })
    });
    const d = await r.json();
    if (!d.ok) { addLog('e', d.error); finishTask('error'); return; }
    pollTask(d.task_id);
  } catch(e) {
    addLog('e', '请求失败：'+e.message);
    finishTask('error');
  }
}

async function pollTask(tid) {
  const poll = async () => {
    try {
      const r = await fetch(`/task/${tid}`);
      const d = await r.json();
      (d.logs||[]).forEach(l => {
        const cls = l.includes('✅')?'o':l.includes('⚠')?'w':l.includes('✗')?'e':
                    (l.includes('📦')||l.includes('📄')||l.includes('🔗'))?'i':'d';
        addLog(cls, l);
      });
      if (d.status === 'running') { setTimeout(poll, 800); return; }
      if (d.status === 'done') {
        document.getElementById('sdot').className = 'sdot ok';
        document.getElementById('progTitle').textContent = '✅ 生成完成';
        document.getElementById('progSub').textContent = `共 ${d.outputs.length} 套单据`;
        showResults(d.outputs);
      } else {
        finishTask('error');
      }
      updateBtns();
    } catch(e) { setTimeout(poll, 2000); }
  };
  poll();
}

function finishTask(status) {
  document.getElementById('sdot').className = `sdot ${status==='error'?'err':'ok'}`;
  if (status==='error') {
    document.getElementById('progTitle').textContent = '❌ 生成失败';
    document.getElementById('progSub').textContent = '请查看日志，确认数据填写是否正确';
  }
  updateBtns();
}

function showResults(outputs) {
  const rc = document.getElementById('resultCard');
  const rg = document.getElementById('resultGrid');
  rc.classList.add('show');
  document.getElementById('resultSub').textContent =
    `批次：${outputs.map(o=>o.batch).join('、')}  ·  ${new Date().toLocaleTimeString()}`;
  rg.innerHTML = '';
  outputs.forEach(out => {
    if (out.set1) rg.innerHTML += fCard('s1', out.set1, out.set1_name, '套一（中国出口）', 'Inv. · PL · 报关单 · 合同 · SI');
    if (out.set2) rg.innerHTML += fCard('s2', out.set2, out.set2_name, '套二（泰国清关）', 'Inv.(Parts) · PL(Parts)');
  });
  S.history.unshift({batch:outputs.map(o=>o.batch).join(', '), time:new Date().toLocaleTimeString(), outputs});
  renderHist();
  rc.scrollIntoView({behavior:'smooth',block:'nearest'});
}

function fCard(cls, fid, fname, tag, sheets) {
  return `<div class="result-file ${cls}">
    <div class="rf-tag">${tag}</div>
    <div class="rf-name">📄 ${fname}</div>
    <div class="rf-sheets">${sheets}</div>
    <button class="btn ${cls==='s1'?'btn-outline':'btn-green'} btn-block"
      onclick="dl('${fid}','${fname}')" style="margin-top:6px">⬇ 下载</button>
  </div>`;
}

async function dl(fid, fname) {
  const a = document.createElement('a');
  a.href = `/download/${fid}`;
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function renderHist() {
  if (!S.history.length) return;
  document.getElementById('histCard').style.display = 'block';
  document.getElementById('histList').innerHTML = S.history.slice(0,8).map(h=>`
    <div class="hist-item">
      <div><div class="hist-batch">${h.batch}</div><div class="hist-time">${h.time}</div></div>
      <div class="hist-actions">
        ${h.outputs.flatMap(o=>[
          o.set1?`<button class="btn btn-outline btn-sm" onclick="dl('${o.set1}','${o.set1_name}')">套一</button>`:'',
          o.set2?`<button class="btn btn-green btn-sm" onclick="dl('${o.set2}','${o.set2_name}')">套二</button>`:'',
        ]).join('')}
      </div>
    </div>`).join('');
}

function addLog(cls, text) {
  const log = document.getElementById('termLog');
  const t = new Date().toLocaleTimeString('zh-CN',{hour12:false});
  const d = document.createElement('div');
  d.className = 'log-line';
  d.innerHTML = `<span class="lt">${t}</span><span class="l${cls}">${esc(text)}</span>`;
  log.appendChild(d);
  log.scrollTop = log.scrollHeight;
}
function esc(t){return t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}

// ── 拖拽 ──────────────────────────────────────────────────────────────────
const packZone = document.getElementById('packingZone');
packZone.addEventListener('dragover', e=>{e.preventDefault();packZone.classList.add('drag-over')});
packZone.addEventListener('dragleave', ()=>packZone.classList.remove('drag-over'));
packZone.addEventListener('drop', e=>{
  e.preventDefault();packZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) {
    const inp = packZone.querySelector('input[type=file]');
    const dt = new DataTransfer();dt.items.add(f);inp.files=dt.files;
    handlePacking(inp);
  }
});
</script>
</body>
</html>"""


# ── API ───────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML_V2)


@app.route('/master/info')
def master_info():
    meta = load_master_meta()
    if meta and os.path.exists(meta.get('path', '')):
        return jsonify(ok=True, meta=meta)
    return jsonify(ok=False)


@app.route('/master/upload', methods=['POST'])
def master_upload():
    f = request.files.get('file')
    if not f:
        return jsonify(ok=False, error='没有收到文件')
    fname = 'master_' + uuid.uuid4().hex[:8] + os.path.splitext(f.filename)[1]
    path = os.path.join(MASTER_DIR, fname)
    f.save(path)
    size = os.path.getsize(path)
    size_str = f'{size/1024:.1f} KB' if size < 1048576 else f'{size/1048576:.1f} MB'
    meta = {
        'path': path,
        'name': f.filename,
        'size': size_str,
        'uploaded_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
    }
    save_master_meta(meta)
    return jsonify(ok=True, **meta)


@app.route('/master/clear', methods=['POST'])
def master_clear():
    if os.path.exists(MASTER_META):
        os.remove(MASTER_META)
    return jsonify(ok=True)


@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f:
        return jsonify(ok=False, error='没有文件')
    uid = uuid.uuid4().hex[:8]
    ext = os.path.splitext(f.filename)[1]
    path = os.path.join(UPLOAD_DIR, f'packing_{uid}{ext}')
    f.save(path)
    size = os.path.getsize(path)
    size_str = f'{size/1024:.1f} KB' if size < 1048576 else f'{size/1048576:.1f} MB'
    return jsonify(ok=True, path=path, size=size_str)


@app.route('/generate', methods=['POST'])
def generate_route():
    data = request.json
    meta = load_master_meta()
    if not meta or not os.path.exists(meta.get('path', '')):
        return jsonify(ok=False, error='请先上传主数据管理手册')

    packing_path = data.get('packing_path')
    batch_no     = data.get('batch_no', '')
    do_all       = data.get('all', False)

    if not packing_path or not os.path.exists(packing_path):
        return jsonify(ok=False, error='请先上传装箱清单')

    task_id = uuid.uuid4().hex[:12]
    tasks[task_id] = {'status':'running','logs':[],'outputs':[],'batch_no':batch_no,'_cur':0}
    threading.Thread(
        target=run_task,
        args=(task_id, meta['path'], packing_path, batch_no, do_all),
        daemon=True
    ).start()
    return jsonify(ok=True, task_id=task_id)


@app.route('/task/<tid>')
def task_status(tid):
    t = tasks.get(tid)
    if not t: return jsonify(status='not_found')
    cur = t['_cur']
    new_logs = t['logs'][cur:]
    t['_cur'] = len(t['logs'])
    return jsonify(status=t['status'], logs=new_logs, outputs=t['outputs'], batch_no=t['batch_no'])


@app.route('/download/<path:fname>')
def download_file(fname):
    safe = os.path.basename(fname)
    path = os.path.join(OUTPUT_DIR, safe)
    if not os.path.exists(path):
        return '文件不存在', 404
    return send_file(path, as_attachment=True, download_name=safe)


def run_task(tid, master_path, packing_path, batch_no, do_all):
    task = tasks[tid]
    def log(msg): task['logs'].append(msg)
    try:
        log('='*50); log('  外贸单据自动化系统 v2.0')
        log(f'  {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'); log('='*50)
        log('  📋 加载主数据管理手册…')
        from engine.config import MasterData
        master = MasterData(master_path)
        log(f'  ✅ 成品档案：{len(master.finished_goods)} 条')
        log(f'  ✅ 物料价格：{len(master.component_prices)} 条')
        log(f'  ✅ 往来方：{len(master.parties)} 个')
        log(f'  ✅ 批次主表：{len(master.shipment_mains)} 票')
        log(f'  ✅ 货柜子表：{len(master.shipment_containers)} 个')

        batches = list(master.shipment_mains.keys()) if do_all else ([batch_no] if batch_no else [])
        if not batches:
            log('  ✗ 未指定批次号'); task['status']='error'; return
        if do_all: log(f'  📋 批量模式：{len(batches)} 个批次')

        from engine.parser import parse_packing_list, build_document_bundle
        from engine.output import generate_document_set
        log('  📦 解析装箱清单…')
        packing = parse_packing_list(packing_path)
        log(f'     客户：{packing.header.customer_code}  日期：{packing.header.shipment_date}')
        log(f'     货柜：{len(packing.containers)} 个  明细：{len(packing.rows)} 行')

        outputs = []
        for bn in batches:
            log(f'\n  {'─'*46}'); log(f'  处理批次：{bn}')
            sm = master.shipment_mains.get(bn)
            if not sm: log(f'  ✗ 批次号未找到，跳过'); continue
            for prod in sm.product_codes:
                if prod not in master.finished_goods:
                    log(f'  ⚠  成品编码 {prod} 未在①成品档案找到')
            try:
                bundle = build_document_bundle(master, bn, packing)
                log(f'  ✅ Invoice No.: {bundle.invoice_no}')
                log(f'     货柜：{[c.container_no for c in bundle.containers]}')
                log(f'     套一：USD {bundle.set1_total_amount:,.2f}  ({sum(l.customs_suits for l in bundle.set1_lines)} 套)')
                log(f'     套二：USD {bundle.set2_total_amount:,.2f}  ({len(bundle.set2_lines)} 部件，{bundle.total_pkgs} 箱)')
                log(f'     GW {bundle.total_gw:.0f} kg  CBM {bundle.total_cbm:.2f} m³')
                log('  📄 生成 Excel…')
                paths = generate_document_set(bundle, OUTPUT_DIR, master_path=master_path)
                s1, s2 = os.path.basename(paths['set1']), os.path.basename(paths['set2'])
                log(f'  ✅ 套一：{s1}'); log(f'  ✅ 套二：{s2}')
                outputs.append({'batch':bn,'set1':s1,'set1_name':s1,'set2':s2,'set2_name':s2})
            except Exception as e:
                log(f'  ✗ 处理异常：{e}'); log(traceback.format_exc()[:300])
        log(f'\n{'='*50}'); log(f'  ✅ 完成！成功 {len(outputs)} 票'); log('='*50)
        task['outputs'] = outputs
        task['status'] = 'done' if outputs else 'error'
    except Exception as e:
        log(f'  ✗ 系统错误：{e}'); log(traceback.format_exc()[:500])
        task['status'] = 'error'


if __name__ == '__main__':
    print('\n' + '='*56)
    print('  外贸单据自动生成系统 v2.0  |  Web版')
    print('  主数据手册固化版 + 公网部署支持')
    print('='*56)
    print('  本地访问：http://localhost:5000')
    print('  公网部署：见 DEPLOY.md')
    print('  按 Ctrl+C 停止')
    print('='*56 + '\n')
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
