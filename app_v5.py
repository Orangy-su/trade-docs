"""
app_v5.py — 外贸单据系统 v5
修复：
- 主数据存储改为JSON（解析结果存文件，解决Render上传问题）
- 货柜信息补录：铅封号/SO号/ETD/过磅重量
- VGM弹窗：理论GW vs 过磅GW，用户选择后计算VGM
- SI完整输出
"""
import os, sys, json, uuid, threading, traceback, io, base64
from flask import Flask, request, jsonify, send_file, render_template_string
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
app.secret_key = os.environ.get('SECRET_KEY', 'qf-trade-v5-2026')

# 存储路径：优先 /data（Render Persistent Disk），其次 /tmp
DATA_DIR = '/data' if os.path.exists('/data') else '/tmp'
UPLOAD_DIR = os.path.join(DATA_DIR, 'uploads'); os.makedirs(UPLOAD_DIR, exist_ok=True)
OUTPUT_DIR = os.path.join(DATA_DIR, 'output');  os.makedirs(OUTPUT_DIR, exist_ok=True)
MASTER_JSON = os.path.join(DATA_DIR, 'master_data.json')  # 解析后的JSON
MASTER_FILE = os.path.join(DATA_DIR, 'master.xlsx')        # 原始文件

tasks = {}

# 柜型皮重映射
TARE_WEIGHTS = {'20GP': 2200, '40GP': 2300, '40HQ': 3900, '45HQ': 4200}

def save_master(data: dict, xlsx_bytes: bytes = None):
    with open(MASTER_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    if xlsx_bytes:
        with open(MASTER_FILE, 'wb') as f:
            f.write(xlsx_bytes)

def load_master() -> dict:
    if os.path.exists(MASTER_JSON):
        with open(MASTER_JSON, encoding='utf-8') as f:
            return json.load(f)
    return None

def get_master_xlsx_path():
    return MASTER_FILE if os.path.exists(MASTER_FILE) else None

def parse_master_xlsx(path: str) -> dict:
    """
    解析主数据管理手册 xlsx
    - 往来方 Sheet: 解析 BUYER / SELLER / FORWARDER 三种类型
    - 成品档案 Sheet: 解析成品编码列表
    逐行容错，单行解析失败不影响整体
    """
    import pandas as pd
    buyers, sellers, forwarders, products = [], [], [], []
    warnings = []

    def sv(row, idx, default=''):
        try:
            val = row.iloc[idx]
            if pd.isna(val): return default
            s = str(val).strip()
            return default if s in ('nan','NaN','None','') else s
        except:
            return default

    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        raise RuntimeError(f'无法打开文件（格式不支持或文件损坏）：{e}')

    for sh in xl.sheet_names:
        try:
            if '往来方' in sh:
                df = pd.read_excel(xl, sheet_name=sh, header=None, dtype=str)
                for _, row in df.iterrows():
                    try:
                        code_val = sv(row, 0)
                        ptype    = sv(row, 1).upper()
                        if not code_val or ptype not in ('SELLER','BUYER','FORWARDER'):
                            continue
                        # 跳过表头/说明行
                        if any(x in code_val for x in ['代码','★','（','说明','Code','类型']):
                            continue
                        name_en = sv(row, 2)
                        name_cn = sv(row, 3)
                        if ptype == 'BUYER':
                            buyers.append({
                                'code': code_val,
                                'name_en': name_en, 'name_cn': name_cn,
                                'address': sv(row, 4), 'address_en': sv(row, 4),
                                'payment_terms': sv(row, 8),
                                'incoterms': sv(row, 9, 'FOB'),
                                'currency': sv(row, 10, 'USD'),
                            })
                        elif ptype == 'SELLER':
                            sellers.append({
                                'code': code_val,
                                'name_en': name_en, 'name_cn': name_cn,
                                'address_en': sv(row, 4),
                                'tax_id': sv(row, 11),
                                'port_loading': sv(row, 12, 'NANSHA'),
                            })
                        elif ptype == 'FORWARDER':
                            forwarders.append({
                                'code': code_val,
                                'name_en': name_en,
                                'address_en': sv(row, 4),
                                'notify_name': sv(row, 5),
                                'notify_address': sv(row, 6),
                                'contact': sv(row, 7),
                            })
                    except Exception as row_e:
                        warnings.append(f'往来方行解析失败: {row_e}')
                        continue

            elif '成品档案' in sh:
                df = pd.read_excel(xl, sheet_name=sh, header=None, dtype=str)
                seen = set()
                for _, row in df.iterrows():
                    try:
                        code_val = sv(row, 0)
                        if not code_val:
                            continue
                        if any(x in code_val for x in [
                            '成品编码','★','（','待填','▌','Sheet',
                            '填写','说明','基础','产品类','Code','编码']):
                            continue
                        if len(code_val) < 5:
                            continue
                        if code_val in seen:
                            continue
                        seen.add(code_val)
                        products.append({
                            'code': code_val,
                            'name_cn': sv(row, 1),
                            'name_en': sv(row, 2),
                        })
                    except Exception as row_e:
                        warnings.append(f'成品档案行解析失败: {row_e}')
                        continue
        except Exception as sh_e:
            warnings.append(f'Sheet [{sh}] 解析失败: {sh_e}')
            continue

    if not buyers and not sellers:
        detail = '; '.join(warnings[:3]) if warnings else '请确认往来方Sheet存在且格式正确'
        raise RuntimeError(f'主数据解析后买方和卖方均为空。{detail}')

    return {
        'buyers': buyers,
        'sellers': sellers,
        'forwarders': forwarders,
        'products': products,
        '_warnings': warnings,
    }



HTML = r"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>外贸单据系统 · 中山汇图</title>
<style>
:root{
  --navy:#1F3864; --blue:#2E86C1; --teal:#117A65; --green:#1E8449;
  --warn:#D68910; --red:#C0392B; --gray:#595959; --lgray:#F8F9FA;
  --border:#DEE2E6; --card:#FFFFFF; --bg:#F2F4F7;
  --mono:'Courier New',monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Arial',sans-serif;background:var(--bg);color:#222;font-size:13px;line-height:1.5}

/* ── 导航栏 ── */
nav{background:var(--navy);padding:0 20px;height:52px;display:flex;align-items:center;
    justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.25)}
.brand{display:flex;align-items:center;gap:12px}
.brand-name{color:#fff;font-size:15px;font-weight:700;letter-spacing:.4px}
.brand-sub{color:rgba(255,255,255,.55);font-size:11px}
.nav-right{display:flex;align-items:center;gap:10px}
.master-pill{display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.12);
  border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:5px 12px;
  cursor:pointer;transition:.2s;user-select:none}
.master-pill:hover{background:rgba(255,255,255,.22)}
.master-pill span{color:#fff;font-size:11.5px;font-weight:600}
.mdot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.mdot.no{background:#F39C12}.mdot.ok{background:#27AE60}
.ver{color:rgba(255,255,255,.4);font-size:11px}

/* ── 主体布局 ── */
.main{max-width:860px;margin:0 auto;padding:20px 16px 60px}

/* ── 卡片 ── */
.card{background:var(--card);border-radius:12px;padding:18px 20px;
      margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.card-hdr{display:flex;align-items:center;gap:12px;margin-bottom:14px}
.card-icon{width:38px;height:38px;border-radius:10px;display:flex;align-items:center;
           justify-content:center;font-size:19px;flex-shrink:0}
.card-title{font-size:15px;font-weight:700;color:var(--navy)}
.card-sub{font-size:11.5px;color:var(--gray);margin-top:2px}

/* ── 上传区 ── */
.upload-zone{border:2px dashed var(--border);border-radius:10px;padding:1.6rem;
  position:relative;text-align:center;transition:.2s;background:#fff;cursor:pointer}
.upload-zone:hover,.upload-zone.drag{border-color:var(--blue);background:#EBF3FB}
.upload-zone.has{border-color:var(--green);background:#F0FBF4;border-style:solid}
.upload-zone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.up-icon{font-size:28px;margin-bottom:8px}
.up-label{font-size:13px;color:#444;font-weight:500}
.up-label strong{color:var(--blue)}
.up-hint{font-size:11px;color:var(--gray);margin-top:4px}

/* ── 表单 ── */
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.form-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
.field label{display:block;font-size:11px;font-weight:700;color:var(--gray);
             text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px}
.field input,.field select,.field textarea{width:100%;padding:7px 10px;border:1.5px solid var(--border);
  border-radius:7px;font-size:13px;background:#fff;font-family:inherit;transition:.15s}
.field input:focus,.field select:focus,.field textarea:focus{
  outline:none;border-color:var(--blue);box-shadow:0 0 0 3px rgba(46,134,193,.12)}
.field input.err,.field select.err{border-color:var(--red)}

/* ── Tab ── */
.tabs{display:flex;gap:0;border-bottom:2px solid var(--border);margin-bottom:16px}
.tab{padding:8px 16px;font-size:12.5px;font-weight:600;color:var(--gray);
     cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-2px;transition:.15s}
.tab.active{color:var(--navy);border-bottom-color:var(--navy)}
.tab:hover:not(.active){color:var(--navy)}
.tab-panel{display:none}.tab-panel.active{display:block}

/* ── 货柜表格 ── */
.ctnr-table{width:100%;border-collapse:collapse;font-size:12px}
.ctnr-table th{background:var(--navy);color:#fff;padding:8px 10px;text-align:left;
               font-size:11px;white-space:nowrap}
.ctnr-table td{padding:7px 8px;border-bottom:1px solid var(--border);vertical-align:middle}
.ctnr-table tr:last-child td{border-bottom:none}
.ctnr-table input{width:100%;padding:5px 7px;border:1.5px solid var(--border);
  border-radius:5px;font-size:12px;font-family:var(--mono)}
.ctnr-table input:focus{outline:none;border-color:var(--blue)}
.ctnr-no{font-family:var(--mono);font-size:11.5px;color:var(--navy);font-weight:600}
.grp-dot{width:22px;height:22px;border-radius:50%;cursor:pointer;border:2px solid #fff;
          box-shadow:0 0 0 1.5px rgba(0,0,0,.15);flex-shrink:0;transition:.1s}
.grp-dot:hover{transform:scale(1.15)}

/* ── 按钮 ── */
.btn{padding:8px 18px;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;
     border:none;transition:.15s;display:inline-flex;align-items:center;gap:6px}
.btn:disabled{opacity:.45;cursor:not-allowed}
.btn-blue{background:var(--blue);color:#fff}.btn-blue:hover:not(:disabled){background:#1A6FA8}
.btn-teal{background:var(--teal);color:#fff}.btn-teal:hover:not(:disabled){background:#0E6655}
.btn-ghost{background:transparent;color:var(--gray);border:1.5px solid var(--border)}
.btn-ghost:hover{background:var(--lgray)}
.btn-lg{padding:11px 28px;font-size:14px}
.btn-sm{padding:5px 12px;font-size:11.5px}

/* ── 进度/终端 ── */
.prog-card,.res-card{display:none}
.terminal{background:#111;color:#7CFC00;font-family:var(--mono);font-size:11.5px;
  padding:12px 14px;border-radius:8px;max-height:220px;overflow-y:auto;line-height:1.7;
  margin-top:8px}
.terminal .err{color:#FF6B6B}.terminal .ok{color:#7CFC00}
.sdot{width:12px;height:12px;border-radius:50%;flex-shrink:0}
.sdot.run{background:var(--warn);animation:pulse 1.2s infinite}
.sdot.done{background:var(--green)}.sdot.err{background:var(--red)}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}

/* ── 结果 ── */
.res-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:4px}
.res-item{border:1.5px solid var(--border);border-radius:10px;padding:14px 16px}
.res-item-title{font-size:13px;font-weight:700;color:var(--navy);margin-bottom:4px}
.res-item-sub{font-size:11px;color:var(--gray);margin-bottom:10px}
.dl-btn{display:flex;align-items:center;gap:6px;padding:7px 14px;border-radius:7px;
  font-size:12px;font-weight:600;text-decoration:none;border:none;cursor:pointer;
  background:var(--navy);color:#fff;transition:.15s;margin-bottom:6px}
.dl-btn:hover{background:var(--blue)}
.dl-btn.pdf{background:#C0392B}.dl-btn.pdf:hover{background:#A93226}

/* ── 提示框 ── */
.alert-box{padding:10px 14px;border-radius:8px;font-size:12px;margin-bottom:12px}
.alert-warn{background:#FEF9E7;border:1px solid #F9E79F;color:#7D6608}
.alert-info{background:#EBF5FB;border:1px solid #AED6F1;color:#1A5276}
.alert-ok{background:#E9F7EF;border:1px solid #A9DFBF;color:#1E8449}

/* ── 分隔线 ── */
.divider{border:none;border-top:1px solid var(--border);margin:16px 0}

/* ── Modal ── */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);
  z-index:1000;align-items:center;justify-content:center;padding:16px}
.modal-bg.open{display:flex}
.modal{background:#fff;border-radius:14px;padding:24px;width:100%;max-width:480px;
       max-height:90vh;overflow-y:auto;box-shadow:0 8px 40px rgba(0,0,0,.2)}
.modal-lg{max-width:680px}
.modal-title{font-size:17px;font-weight:700;color:var(--navy);margin-bottom:6px}
.modal-sub{font-size:12px;color:var(--gray);margin-bottom:16px;line-height:1.5}
.modal-footer{display:flex;justify-content:flex-end;gap:8px;margin-top:16px;padding-top:12px;
              border-top:1px solid var(--border)}
/* ── VGM表 ── */
.vgm-table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px}
.vgm-table th{background:var(--navy);color:#fff;padding:8px 10px;text-align:center}
.vgm-table td{padding:8px 10px;border-bottom:1px solid var(--border);text-align:center}
.vgm-table input{width:90px;padding:5px 7px;border:1.5px solid var(--border);
  border-radius:5px;text-align:center}
/* ── 历史 ── */
.hist-row{display:flex;align-items:center;gap:10px;padding:8px 0;
          border-bottom:1px solid var(--border);font-size:12px}
.hist-row:last-child{border-bottom:none}
.hist-inv{font-family:var(--mono);font-weight:600;color:var(--navy);min-width:160px}

@media(max-width:600px){
  .form-grid,.form-grid-3,.res-grid{grid-template-columns:1fr}
  .ctnr-table{font-size:11px}
  .modal{padding:16px}
}
</style>
</head>
<body>

<!-- 导航栏 —— 主数据状态整合在这里，去掉单独的状态卡 -->
<nav>
  <div class="brand">
    <div>
      <div class="brand-name">🗂 外贸单据系统</div>
      <div class="brand-sub">中山汇图电器有限公司</div>
    </div>
  </div>
  <div class="nav-right">
    <!-- 单一主数据入口，状态在这里直接显示 -->
    <div class="master-pill" onclick="openMasterModal()" title="点击上传或更新主数据">
      <div class="mdot no" id="mdot"></div>
      <span id="masterTxt">主数据未上传</span>
    </div>
    <span class="ver">v5.0</span>
  </div>
</nav>

<main class="main">

<!-- Step 1: 上传装箱清单 -->
<div class="card">
  <div class="card-hdr">
    <div class="card-icon" style="background:#FEF9E7">📦</div>
    <div>
      <div class="card-title">Step 1 · 上传装箱清单</div>
      <div class="card-sub">仓管员填写后提交，支持新旧格式 · 自动识别货柜与物料</div>
    </div>
  </div>
  <div class="upload-zone" id="packZone">
    <input type="file" id="packFile" accept=".xlsx,.xls" onchange="handlePack(this)">
    <div class="up-icon">🗂️</div>
    <div class="up-label">拖拽装箱清单到此处，或 <strong>点击选择</strong></div>
    <div class="up-hint">支持 .xlsx · .xls · 任意文件名</div>
  </div>
</div>

<!-- Step 2: 出货信息 -->
<div class="card" id="orderCard" style="display:none">
  <div class="card-hdr">
    <div class="card-icon" style="background:#EBF3FB">📋</div>
    <div>
      <div class="card-title">Step 2 · 填写出货信息</div>
      <div class="card-sub">选择客户和卖方主体，填写发票信息，配置货柜 SO 号</div>
    </div>
  </div>

  <div class="tabs">
    <div class="tab active" onclick="switchTab('basic')">基础信息</div>
    <div class="tab" onclick="switchTab('containers')">货柜配置</div>
    <div class="tab" onclick="switchTab('advanced')">高级设置</div>
  </div>

  <!-- Tab: 基础信息 -->
  <div class="tab-panel active" id="tab-basic">
    <div class="form-grid" style="margin-bottom:14px">
      <div class="field">
        <label>买方客户 ★</label>
        <select id="buyerSel" onchange="updateBuyer()">
          <option value="">— 选择客户 —</option>
        </select>
      </div>
      <div class="field">
        <label>卖方主体 ★</label>
        <select id="sellerSel" onchange="updateGenBtn()">
          <option value="">— 选择卖方 —</option>
        </select>
      </div>
    </div>
    <div class="form-grid" style="margin-bottom:14px">
      <div class="field" style="font-family:var(--mono)">
        <label>Invoice No. ★</label>
        <input type="text" id="invoiceNo" placeholder="如：HT-26-HT22060101A" oninput="updateGenBtn()">
      </div>
      <div class="field">
        <label>出货日期 ★ (Invoice Date)</label>
        <input type="date" id="shipDate" onchange="updateGenBtn()">
      </div>
    </div>
    <!-- 成品套数 -->
    <div id="productSuitsArea" style="display:none;margin-bottom:14px">
      <div style="font-size:12px;font-weight:600;color:var(--navy);margin-bottom:8px">
        报关整机套数 ★
        <span style="font-weight:400;color:var(--gray);font-size:11px">从装箱清单自动识别</span>
      </div>
      <div id="productSuitsGrid" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:10px"></div>
    </div>
    <div id="productSuitsEmpty" style="padding:10px 14px;background:var(--lgray);border-radius:8px;
      font-size:12px;color:var(--gray);border:1px dashed var(--border);margin-bottom:14px">
      ⬆ 请先上传装箱清单，系统自动识别成品编码
    </div>
    <div id="buyerInfoBar" style="display:none;padding:8px 14px;background:var(--lgray);
      border-radius:8px;font-size:12px;color:var(--gray);border:1px solid var(--border)"></div>
  </div>

  <!-- Tab: 货柜配置 -->
  <div class="tab-panel" id="tab-containers">
    <div class="alert-box alert-info" style="margin-bottom:12px">
      <strong>SO号 & 铅封号</strong>为 SI 必填项，ETD 为参考项。过磅重量用于 VGM 计算，留空则用装箱单理论重量。<br>
      <strong>单据分组：</strong>同色 = 合并出一套单据 · 点击色块切换分组
    </div>
    <div style="overflow-x:auto">
      <table class="ctnr-table" id="ctnrTable">
        <thead>
          <tr>
            <th style="width:44px">序号</th>
            <th style="width:130px">货柜号</th>
            <th style="width:100px">铅封号 ★</th>
            <th style="width:80px">柜型</th>
            <th style="width:140px">SO号 ★</th>
            <th style="width:88px">ETD</th>
            <th style="width:90px">过磅GW(kg)</th>
            <th style="width:90px">理论GW(kg)</th>
            <th style="width:85px">柜皮重(kg)</th>
            <th style="width:90px">单据分组</th>
          </tr>
        </thead>
        <tbody id="ctnrBody">
          <tr><td colspan="10" style="text-align:center;color:var(--gray);padding:2rem">
            请先上传装箱清单
          </td></tr>
        </tbody>
      </table>
    </div>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap;align-items:center">
      <span style="font-size:11px;color:var(--gray)">分组色标：</span>
      <span style="background:#EBF3FB;color:#1A5276;padding:2px 10px;border-radius:10px;font-size:11px;font-weight:700">● 分组A</span>
      <span style="background:#E9F7EF;color:#1E8449;padding:2px 10px;border-radius:10px;font-size:11px;font-weight:700">● 分组B</span>
      <span style="background:#FEF9E7;color:#7D6608;padding:2px 10px;border-radius:10px;font-size:11px;font-weight:700">● 分组C</span>
    </div>
  </div>

  <!-- Tab: 高级设置 -->
  <div class="tab-panel" id="tab-advanced">
    <div class="form-grid-3">
      <div class="field">
        <label>付款条件</label>
        <input type="text" id="payTerms" placeholder="T/T 45 days after B/L">
      </div>
      <div class="field">
        <label>贸易术语</label>
        <select id="incoterms"><option>FOB</option><option>CIF</option><option>EXW</option></select>
      </div>
      <div class="field">
        <label>货币</label>
        <select id="currency"><option>USD</option><option>CNY</option><option>EUR</option></select>
      </div>
    </div>
    <div class="form-grid" style="margin-top:14px">
      <div class="field">
        <label>合同日期偏移（天，负=早于发票）</label>
        <input type="number" id="contractOffset" value="-30">
      </div>
      <div class="field">
        <label>批次号（留空自动生成）</label>
        <input type="text" id="batchNo" placeholder="自动生成">
      </div>
    </div>
  </div>

  <hr class="divider">
  <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:center">
    <button class="btn btn-blue btn-lg" id="genBtn" onclick="preGenerate()" disabled>
      ⚡ 生成单据
    </button>
    <div id="readyHint" style="font-size:12px;color:var(--gray)">请完成上方必填项</div>
  </div>
</div>

<!-- 进度 -->
<div class="card prog-card" id="progCard">
  <div style="display:flex;align-items:center;gap:10px;margin-bottom:1rem">
    <div class="sdot run" id="sdot"></div>
    <div>
      <div style="font-size:15px;font-weight:700;color:var(--navy)" id="progTitle">正在生成…</div>
      <div style="font-size:12px;color:var(--gray)" id="progSub">请稍候</div>
    </div>
  </div>
  <div class="terminal" id="termLog"></div>
</div>

<!-- 结果 -->
<div class="card res-card" id="resCard">
  <div class="card-hdr">
    <span style="font-size:22px">🎉</span>
    <div>
      <div class="card-title">单据生成完成</div>
      <div class="card-sub" id="resSub"></div>
    </div>
  </div>
  <div class="alert-box alert-warn" id="gwReminder" style="display:none;margin-bottom:12px">
    ⚠️ <strong>注意：</strong>SI使用了过磅重量，与装箱单理论值有差异，请在下载后手动同步 PL 的 GW/NW 数值。
  </div>
  <div class="res-grid" id="resGrid"></div>
</div>

<!-- 历史 -->
<div class="card" id="histCard" style="display:none">
  <div class="card-hdr">
    <span style="font-size:18px">🕐</span>
    <div class="card-title">本次会话记录</div>
  </div>
  <div id="histList"></div>
</div>

</main>

<!-- ═══ 主数据 Modal（唯一入口）═══ -->
<div class="modal-bg" id="masterModal">
  <div class="modal">
    <div class="modal-title">📋 主数据管理手册</div>
    <div class="modal-sub">上传一次长期有效，服务器端保存解析结果，下次打开自动加载。如有更新重新上传即可覆盖。</div>
    <div class="upload-zone" id="masterZone" style="padding:1.4rem">
      <input type="file" id="masterFile" accept=".xlsx,.xls" onchange="handleMaster(this)">
      <div class="up-icon" id="mzIcon">📄</div>
      <div id="mzText">
        <div class="up-label">拖拽主数据管理手册到此<br>或 <strong>点击选择</strong></div>
        <div class="up-hint">任意文件名均可</div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeMasterModal()">关闭</button>
    </div>
  </div>
</div>

<!-- ═══ VGM 确认 Modal ═══ -->
<div class="modal-bg" id="vgmModal">
  <div class="modal modal-lg">
    <div class="modal-title">⚖️ 确认 SI 重量 & VGM</div>
    <div class="modal-sub">选择每个货柜 SI 使用的重量数据，VGM = 选定 GW + 柜皮重（自动带入）。</div>
    <table class="vgm-table">
      <thead>
        <tr>
          <th>货柜</th><th>理论GW(kg)<br>装箱单汇总</th><th>过磅GW(kg)<br>实测重量</th>
          <th>SI使用</th><th>柜皮重(kg)</th><th>VGM(kg)</th>
        </tr>
      </thead>
      <tbody id="vgmBody"></tbody>
    </table>
    <div class="alert-box alert-warn">
      ⚠️ 若使用过磅重量，与装箱单理论重量有差异，生成后请手动更新 PL 的 GW/NW 数据。
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeVgmModal()">取消</button>
      <button class="btn btn-blue" onclick="confirmVgm()">✅ 确认并生成单据</button>
    </div>
  </div>
</div>

<!-- ═══ 新增买方 Modal ═══ -->
<div class="modal-bg" id="newBuyerModal">
  <div class="modal">
    <div class="modal-title">新增买方客户</div>
    <div class="modal-sub">填写后自动保存到主数据，下次直接下拉选择。</div>
    <div class="form-grid" style="margin-bottom:12px">
      <div class="field">
        <label>客户代码 ★</label>
        <input type="text" id="nb_code" placeholder="如：US13">
      </div>
      <div class="field">
        <label>买方英文全称 ★</label>
        <input type="text" id="nb_name_en">
      </div>
    </div>
    <div class="field" style="margin-bottom:12px">
      <label>买方地址（英文）★</label>
      <textarea id="nb_addr" rows="2" style="resize:vertical"></textarea>
    </div>
    <div class="form-grid-3" style="margin-bottom:12px">
      <div class="field">
        <label>付款条件</label>
        <input type="text" id="nb_payment" placeholder="T/T 45 days after B/L">
      </div>
      <div class="field">
        <label>贸易术语</label>
        <input type="text" id="nb_inco" placeholder="FOB">
      </div>
      <div class="field">
        <label>货币</label>
        <input type="text" id="nb_currency" value="USD">
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeNewBuyerModal()">取消</button>
      <button class="btn btn-teal" onclick="submitNewBuyer()">✅ 保存并选中</button>
    </div>
  </div>
</div>

<script>
// ═══════════════════════════════════════════════════════
// 状态
// ═══════════════════════════════════════════════════════
const S = {
  masterLoaded: false, masterName: '', masterData: null,
  packLoaded: false, packData: null,
  products: [], containers: [], vgmData: [],
  currentBatch: null, useWeighed: {}
};
const GRP_COLORS = [
  {bg:'#EBF3FB',fg:'#1A5276',label:'A'},
  {bg:'#E9F7EF',fg:'#1E8449',label:'B'},
  {bg:'#FEF9E7',fg:'#7D6608',label:'C'},
];

// ═══════════════════════════════════════════════════════
// 初始化
// ═══════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', async () => {
  // 自动加载已保存的主数据
  try {
    const r = await fetch('/master/info');
    const d = await r.json();
    if(d.ok){ S.masterLoaded=true; S.masterName=d.name||'主数据已就绪'; S.masterData=d.data; updateMasterUI(true); populateSelects(); }
  } catch(e){}

  // 今天日期
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('shipDate').value = today;

  // 拖拽区
  setupDrop('packZone', f => { const i=document.getElementById('packFile'); i.files; handlePackFile(f); });
  setupDrop('masterZone', f => { handleMasterFile(f); });
});

function setupDrop(zoneId, cb){
  const z = document.getElementById(zoneId);
  if(!z) return;
  z.addEventListener('dragover', e=>{ e.preventDefault(); z.classList.add('drag'); });
  z.addEventListener('dragleave', ()=> z.classList.remove('drag'));
  z.addEventListener('drop', e=>{ e.preventDefault(); z.classList.remove('drag'); const f=e.dataTransfer.files[0]; if(f) cb(f); });
}

// ═══════════════════════════════════════════════════════
// 主数据（唯一入口：导航栏 master-pill）
// ═══════════════════════════════════════════════════════
function openMasterModal(){ document.getElementById('masterModal').classList.add('open'); }
function closeMasterModal(){ document.getElementById('masterModal').classList.remove('open'); }

function handleMaster(input){ if(input.files[0]) handleMasterFile(input.files[0]); }

async function handleMasterFile(file){
  const zone = document.getElementById('masterZone');
  const icon = document.getElementById('mzIcon');
  const txt  = document.getElementById('mzText');
  icon.textContent = '⏳'; txt.innerHTML = '<div class="up-label">解析中，请稍候…</div>';
  zone.classList.remove('has');
  const fd = new FormData(); fd.append('file', file);
  try {
    const r = await fetch('/master/upload', {method:'POST', body:fd});
    const d = await r.json();
    if(d.ok){
      S.masterLoaded=true; S.masterName=d.name||file.name; S.masterData=d.data;
      updateMasterUI(true); populateSelects();
      icon.textContent='✅'; txt.innerHTML=`<div class="up-label" style="color:var(--green)">主数据已加载：${file.name}</div>`;
      zone.classList.add('has');
      setTimeout(()=>closeMasterModal(), 1200);
    } else {
      icon.textContent='❌'; txt.innerHTML=`<div class="up-label" style="color:var(--red)">${d.error||'解析失败'}</div>`;
    }
  } catch(e){ icon.textContent='❌'; txt.innerHTML='<div class="up-label" style="color:var(--red)">上传失败，请重试</div>'; }
}

function updateMasterUI(loaded){
  const dot = document.getElementById('mdot');
  const txt = document.getElementById('masterTxt');
  dot.className = 'mdot ' + (loaded ? 'ok' : 'no');
  txt.textContent = loaded ? S.masterName : '主数据未上传';
}

function populateSelects(){
  if(!S.masterData) return;
  const buyers = S.masterData.buyers||[];
  const sellers = S.masterData.sellers||[];
  const bs = document.getElementById('buyerSel');
  const ss = document.getElementById('sellerSel');
  const cur = bs.value;
  bs.innerHTML = '<option value="">— 选择客户 —</option>';
  buyers.forEach(b=>{ const o=document.createElement('option'); o.value=b.code; o.textContent=`${b.code} · ${b.name_en}`; bs.appendChild(o); });
  // 新增客户入口
  const add = document.createElement('option'); add.value='__new__'; add.textContent='+ 新增买方客户…'; bs.appendChild(add);
  if(cur) bs.value=cur;

  ss.innerHTML='<option value="">— 选择卖方 —</option>';
  sellers.forEach(s=>{ const o=document.createElement('option'); o.value=s.code; o.textContent=`${s.code} · ${s.name_en||s.name_cn}`; ss.appendChild(o); });
  if(sellers.length===1) ss.value=sellers[0].code;
}

function updateBuyer(){
  const v = document.getElementById('buyerSel').value;
  if(v==='__new__'){ document.getElementById('buyerSel').value=''; openNewBuyerModal(); return; }
  const bar = document.getElementById('buyerInfoBar');
  if(S.masterData && v){
    const b = (S.masterData.buyers||[]).find(x=>x.code===v);
    if(b){ bar.style.display='block'; bar.innerHTML=`<strong>${b.name_en}</strong><br>${b.address||''}<br>付款：${b.payment_terms||'—'} · 贸易术语：${b.incoterms||'—'}`; }
    // 自动填付款条件
    if(b&&b.payment_terms) document.getElementById('payTerms').value=b.payment_terms;
    if(b&&b.incoterms){ const sel=document.getElementById('incoterms'); for(let i=0;i<sel.options.length;i++) if(sel.options[i].value===b.incoterms){sel.selectedIndex=i;break;} }
  } else { bar.style.display='none'; }
  updateGenBtn();
}

// ═══════════════════════════════════════════════════════
// 装箱清单
// ═══════════════════════════════════════════════════════
function handlePack(input){ if(input.files[0]) handlePackFile(input.files[0]); }

async function handlePackFile(file){
  const zone = document.getElementById('packZone');
  zone.innerHTML = `<div class="up-icon">⏳</div><div class="up-label">解析中…</div>`;
  const fd = new FormData(); fd.append('file', file);
  try {
    const r = await fetch('/upload', {method:'POST', body:fd});
    const d = await r.json();
    if(d.ok){
      S.packLoaded=true; S.packData=d;
      S.products=d.products||[]; S.containers=d.containers||[];
      zone.className='upload-zone has';
      zone.innerHTML=`<div class="up-icon">✅</div>
        <div class="up-label" style="color:var(--green)">${file.name}</div>
        <div class="up-hint">${d.containers?.length||0} 个货柜 · ${d.rows||0} 行物料</div>
        <input type="file" accept=".xlsx,.xls" onchange="handlePack(this)" style="position:absolute;inset:0;opacity:0;cursor:pointer">`;
      buildOrderCard();
      document.getElementById('orderCard').style.display='block';
    } else {
      zone.className='upload-zone'; zone.innerHTML=`<div class="up-icon">❌</div><div class="up-label" style="color:var(--red)">${d.error||'解析失败'}</div><input type="file" accept=".xlsx,.xls" onchange="handlePack(this)" style="position:absolute;inset:0;opacity:0;cursor:pointer">`;
    }
  } catch(e){
    zone.className='upload-zone'; zone.innerHTML=`<div class="up-icon">❌</div><div class="up-label" style="color:var(--red)">上传失败</div><input type="file" accept=".xlsx,.xls" onchange="handlePack(this)" style="position:absolute;inset:0;opacity:0;cursor:pointer">`;
  }
}

// ═══════════════════════════════════════════════════════
// 订单卡片构建
// ═══════════════════════════════════════════════════════
function buildOrderCard(){
  buildProductSuits();
  buildCtnrTable();
  updateGenBtn();
}

function buildProductSuits(){
  const area  = document.getElementById('productSuitsArea');
  const empty = document.getElementById('productSuitsEmpty');
  const grid  = document.getElementById('productSuitsGrid');
  if(!S.products.length){ area.style.display='none'; empty.style.display='block'; return; }
  area.style.display='block'; empty.style.display='none'; grid.innerHTML='';
  S.products.forEach(p=>{
    const div=document.createElement('div');
    div.style.cssText='background:var(--lgray);border:1.5px solid var(--border);border-radius:8px;padding:10px 12px;display:flex;align-items:center;gap:10px';
    div.innerHTML=`<div style="flex:1"><div style="font-size:11px;font-weight:700;color:var(--navy)">${p.code}</div><div style="font-size:11px;color:var(--gray);margin-top:2px">${p.name||''}</div></div><input type="number" min="1" value="${p.suits||486}" style="width:80px;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:13px;font-weight:700;text-align:center" data-code="${p.code}" oninput="updateGenBtn()">`;
    grid.appendChild(div);
  });
}

function buildCtnrTable(){
  const tbody=document.getElementById('ctnrBody'); tbody.innerHTML='';
  S.containers.forEach((c,i)=>{
    if(!c.grp) c.grp=0;
    const g=GRP_COLORS[c.grp%GRP_COLORS.length];
    const tr=document.createElement('tr');
    tr.innerHTML=`
      <td style="text-align:center;font-weight:700;color:var(--gray)">${i+1}</td>
      <td><div class="ctnr-no">${c.container_no||'—'}</div></td>
      <td><input value="${c.seal_no||''}" placeholder="铅封号" oninput="S.containers[${i}].seal_no=this.value"></td>
      <td>
        <select style="width:100%;padding:5px 4px;border:1.5px solid var(--border);border-radius:5px;font-size:12px" onchange="S.containers[${i}].container_size=this.value">
          ${['40HQ','45HQ','40GP','20GP'].map(sz=>`<option${sz===(c.container_size||'40HQ')?' selected':''}>${sz}</option>`).join('')}
        </select>
      </td>
      <td><input value="${c.so_no||''}" placeholder="SO号" style="font-family:var(--mono)" oninput="S.containers[${i}].so_no=this.value;updateGenBtn()"></td>
      <td><input type="date" value="${c.etd||''}" oninput="S.containers[${i}].etd=this.value"></td>
      <td><input type="number" value="${c.gw_weighed||''}" placeholder="实测" style="text-align:right" oninput="S.containers[${i}].gw_weighed=parseFloat(this.value)||0"></td>
      <td style="text-align:center;color:var(--gray);font-size:12px">${c.gw_theory ? c.gw_theory.toFixed(0) : '—'}</td>
      <td style="text-align:center">
        <select style="width:75px;padding:5px 4px;border:1.5px solid var(--border);border-radius:5px;font-size:12px" onchange="S.containers[${i}].tare=parseInt(this.value)">
          ${[[20,2200],[40,2300],[40.5,3900],[45,4200]].map(([l,t])=>`<option value="${t}"${(c.tare||3900)===t?' selected':''}>${t}</option>`).join('')}
        </select>
      </td>
      <td style="text-align:center">
        <div class="grp-dot" style="background:${g.bg};border-color:${g.fg};margin:0 auto" title="分组${g.label}" onclick="cycleGrp(${i},this)"></div>
      </td>`;
    tbody.appendChild(tr);
  });
}

function cycleGrp(idx, el){
  S.containers[idx].grp = ((S.containers[idx].grp||0)+1) % GRP_COLORS.length;
  const g=GRP_COLORS[S.containers[idx].grp]; el.style.background=g.bg; el.style.borderColor=g.fg;
}

// ═══════════════════════════════════════════════════════
// Tab 切换
// ═══════════════════════════════════════════════════════
function switchTab(id){
  document.querySelectorAll('.tab,.tab-panel').forEach(e=>e.classList.remove('active'));
  event.target.classList.add('active');
  document.getElementById('tab-'+id).classList.add('active');
}

// ═══════════════════════════════════════════════════════
// 生成按钮可用性检查
// ═══════════════════════════════════════════════════════
function updateGenBtn(){
  const buyer   = document.getElementById('buyerSel').value;
  const seller  = document.getElementById('sellerSel').value;
  const invoice = document.getElementById('invoiceNo').value.trim();
  const date    = document.getElementById('shipDate').value;
  const soOK    = S.containers.length===0 || S.containers.some(c=>c.so_no&&c.so_no.trim());
  const ok = buyer&&seller&&invoice&&date&&soOK&&!S.containers.every(c=>!(c.so_no&&c.so_no.trim()));
  const btn = document.getElementById('genBtn');
  const hint = document.getElementById('readyHint');
  btn.disabled = !ok;
  if(!buyer||!seller)      hint.textContent='请选择买方客户和卖方主体';
  else if(!invoice)        hint.textContent='请填写 Invoice No.';
  else if(!date)           hint.textContent='请选择出货日期';
  else if(!soOK)           hint.textContent='请在货柜配置中填写 SO 号';
  else                     hint.textContent='';
}

// ═══════════════════════════════════════════════════════
// VGM 弹窗
// ═══════════════════════════════════════════════════════
function preGenerate(){
  // 检查是否有过磅重量
  const hasWeighed = S.containers.some(c=>c.gw_weighed && c.gw_weighed>0);
  if(hasWeighed){ buildVgmModal(); openVgmModal(); }
  else generate({});
}

function buildVgmModal(){
  const tbody=document.getElementById('vgmBody'); tbody.innerHTML='';
  S.containers.forEach((c,i)=>{
    const tare=c.tare||3900;
    const hasW=c.gw_weighed&&c.gw_weighed>0;
    tbody.innerHTML+=`<tr>
      <td class="ctnr-no">${c.container_no||i+1}</td>
      <td>${c.gw_theory ? c.gw_theory.toFixed(0) : '—'}</td>
      <td>${hasW ? c.gw_weighed.toFixed(0) : '—'}</td>
      <td>
        <select id="vgm_src_${i}" onchange="updateVgmRow(${i},${tare})">
          <option value="theory">理论GW</option>
          ${hasW?'<option value="weighed">过磅GW</option>':''}
        </select>
      </td>
      <td>${tare}</td>
      <td id="vgm_result_${i}">${((c.gw_theory||0)+tare).toFixed(0)}</td>
    </tr>`;
  });
}

function updateVgmRow(idx,tare){
  const sel=document.getElementById(`vgm_src_${idx}`);
  const c=S.containers[idx];
  const gw = sel.value==='weighed' ? (c.gw_weighed||0) : (c.gw_theory||0);
  document.getElementById(`vgm_result_${idx}`).textContent=(gw+tare).toFixed(0);
}

function openVgmModal(){ document.getElementById('vgmModal').classList.add('open'); }
function closeVgmModal(){ document.getElementById('vgmModal').classList.remove('open'); }

function confirmVgm(){
  const useWeighed={};
  S.containers.forEach((_,i)=>{
    const sel=document.getElementById(`vgm_src_${i}`);
    if(sel) useWeighed[i]=(sel.value==='weighed');
  });
  closeVgmModal();
  generate(useWeighed);
}

// ═══════════════════════════════════════════════════════
// 生成单据
// ═══════════════════════════════════════════════════════
function generate(useWeighed){
  // 收集套数
  const suitsMap={};
  document.querySelectorAll('#productSuitsGrid input[data-code]').forEach(inp=>{
    suitsMap[inp.dataset.code]=parseInt(inp.value)||0;
  });

  // 货柜分组
  const grps={};
  S.containers.forEach((c,i)=>{
    const g=c.grp||0; if(!grps[g]) grps[g]=[];
    grps[g].push({seq:c.seq||i+1, ...c, use_weighed:useWeighed[i]||false});
  });

  // 构造 batches
  const buyer=document.getElementById('buyerSel').value;
  const seller=document.getElementById('sellerSel').value;
  const invoice=document.getElementById('invoiceNo').value.trim();
  const date=document.getElementById('shipDate').value;
  const payTerms=document.getElementById('payTerms').value.trim();
  const inco=document.getElementById('incoterms').value;
  const curr=document.getElementById('currency').value;
  const offset=parseInt(document.getElementById('contractOffset').value)||(-30);
  const batchNo=document.getElementById('batchNo').value.trim();
  const products=S.products.map(p=>({code:p.code,suits:suitsMap[p.code]||p.suits||0}));

  const batches=Object.values(grps).map((ctnrs,gi)=>{
    const seqs=ctnrs.map(c=>c.seq||0);
    return {
      batch_no: batchNo || `${invoice}-G${gi+1}`,
      customer_code: buyer, seller_code: seller,
      shipment_date: date, invoice_no: invoice,
      product_codes: products.map(p=>p.code),
      customs_suits: products.map(p=>p.suits),
      container_seqs: seqs,
      pay_terms: payTerms, incoterms: inco, currency: curr,
      contract_offset: offset,
      container_details: ctnrs.map(c=>({
        seq: c.seq, container_no: c.container_no,
        seal_no: c.seal_no||'', container_size: c.container_size||'40HQ',
        so_no: c.so_no||'', etd: c.etd||'',
        port_loading: 'NANSHA',
        vgm: c.use_weighed
              ? (c.gw_weighed||0)+(c.tare||3900)
              : (c.gw_theory||0)+(c.tare||3900),
        gw_final: c.use_weighed ? (c.gw_weighed||0) : 0,
      })),
    };
  });

  // 显示进度卡
  document.getElementById('progCard').style.display='block';
  document.getElementById('resCard').style.display='none';
  document.getElementById('progTitle').textContent='正在生成…';
  document.getElementById('progSub').textContent='请稍候';
  document.getElementById('sdot').className='sdot run';
  const log=document.getElementById('termLog'); log.innerHTML='';
  const addLog=(t)=>{ log.innerHTML+=`<div>${t.replace(/✅/g,'<span class=ok>✅</span>').replace(/✗/g,'<span class=err>✗</span>')}</div>`; log.scrollTop=log.scrollHeight; };

  // 提交
  fetch('/generate_v4',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({batches, pack_session:S.packData?.session})})
  .then(r=>r.json()).then(d=>{
    if(!d.ok){ addLog('✗ 提交失败：'+(d.error||'')); document.getElementById('sdot').className='sdot err'; return; }
    const tid=d.task_id; addLog('任务已提交，等待生成…');
    const poll=setInterval(async()=>{
      const tr=await fetch('/task/'+tid).then(r=>r.json());
      (tr.logs||[]).slice(log.children.length).forEach(l=>addLog(l));
      if(tr.status==='done'||tr.status==='error'){
        clearInterval(poll);
        document.getElementById('sdot').className='sdot '+(tr.status==='done'?'done':'err');
        document.getElementById('progTitle').textContent=tr.status==='done'?'✅ 生成完成':'❌ 生成失败';
        document.getElementById('progSub').textContent='';
        if(tr.status==='done') showResults(tr.outputs||[],useWeighed);
      }
    },1000);
  }).catch(e=>{ addLog('✗ 网络错误：'+e); document.getElementById('sdot').className='sdot err'; });
}

// ═══════════════════════════════════════════════════════
// 结果展示
// ═══════════════════════════════════════════════════════
function showResults(outputs, useWeighed){
  document.getElementById('resCard').style.display='block';
  const hasW=Object.values(useWeighed).some(v=>v);
  document.getElementById('gwReminder').style.display=hasW?'block':'none';
  document.getElementById('resSub').textContent=`共生成 ${outputs.length} 套单据`;
  const grid=document.getElementById('resGrid'); grid.innerHTML='';
  outputs.forEach(o=>{
    const div=document.createElement('div'); div.className='res-item';
    div.innerHTML=`
      <div class="res-item-title">📁 批次 ${o.batch}</div>
      <div class="res-item-sub">套一（整机发票/报关单）+ 套二（散件装箱单）</div>
      <a class="dl-btn" href="/download/${o.set1}" download="${o.set1_name||o.set1}">⬇ 套一 Excel</a>
      <a class="dl-btn" href="/download/${o.set2}" download="${o.set2_name||o.set2}">⬇ 套二 Excel</a>`;
    grid.appendChild(div);
  });

  // 历史记录
  const hist=document.getElementById('histCard');
  hist.style.display='block';
  const list=document.getElementById('histList');
  outputs.forEach(o=>{
    const row=document.createElement('div'); row.className='hist-row';
    const ts=new Date().toLocaleTimeString('zh-CN',{hour:'2-digit',minute:'2-digit'});
    row.innerHTML=`<span style="color:var(--gray);min-width:45px">${ts}</span>
      <span class="hist-inv">${o.batch}</span>
      <a href="/download/${o.set1}" style="color:var(--blue);font-size:12px" download>套一</a>
      <span style="color:var(--border)">·</span>
      <a href="/download/${o.set2}" style="color:var(--blue);font-size:12px" download>套二</a>`;
    list.prepend(row);
  });
}

// ═══════════════════════════════════════════════════════
// 新增买方 Modal
// ═══════════════════════════════════════════════════════
function openNewBuyerModal(){ document.getElementById('newBuyerModal').classList.add('open'); }
function closeNewBuyerModal(){ document.getElementById('newBuyerModal').classList.remove('open'); }

async function submitNewBuyer(){
  const payload={
    code: document.getElementById('nb_code').value.trim(),
    name_en: document.getElementById('nb_name_en').value.trim(),
    address: document.getElementById('nb_addr').value.trim(),
    payment_terms: document.getElementById('nb_payment').value.trim(),
    incoterms: document.getElementById('nb_inco').value.trim(),
    currency: document.getElementById('nb_currency').value.trim(),
  };
  if(!payload.code||!payload.name_en){ alert('客户代码和英文名称为必填项'); return; }
  const r=await fetch('/buyer/add',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const d=await r.json();
  if(d.ok){
    if(!S.masterData) S.masterData={buyers:[],sellers:[]};
    if(!S.masterData.buyers) S.masterData.buyers=[];
    S.masterData.buyers.push(payload);
    populateSelects();
    document.getElementById('buyerSel').value=payload.code;
    updateBuyer();
    closeNewBuyerModal();
    ['nb_code','nb_name_en','nb_addr','nb_payment','nb_inco'].forEach(id=>document.getElementById(id).value='');
    document.getElementById('nb_currency').value='USD';
  } else { alert(d.error||'保存失败'); }
}

// ═══════════════════════════════════════════════════════
// 点击背景关闭 Modal
// ═══════════════════════════════════════════════════════
document.addEventListener('click', e=>{
  ['masterModal','vgmModal','newBuyerModal'].forEach(id=>{
    const el=document.getElementById(id);
    if(el&&e.target===el) el.classList.remove('open');
  });
});
</script>
</body>
</html>"""






# ── API ───────────────────────────────────────────────────────────────────────

@app.route('/')
def index(): return render_template_string(HTML)

@app.route('/master/info')
def master_info():
    data = load_master()
    if data:
        name = data.get('_meta', {}).get('name', '主数据已就绪')
        return jsonify(ok=True, name=name, data={
            'buyers':   data.get('buyers', []),
            'sellers':  data.get('sellers', []),
            'products': data.get('products', []),
            'forwarders': data.get('forwarders', []),
        })
    return jsonify(ok=False)

@app.route('/master/upload', methods=['POST'])
def master_upload():
    f = request.files.get('file')
    if not f: return jsonify(ok=False, error='没有文件')
    try:
        xlsx_bytes = f.read()
        # 先保存原始文件
        with open(MASTER_FILE, 'wb') as out: out.write(xlsx_bytes)
        # 解析
        data = parse_master_xlsx(MASTER_FILE)
        data['_meta'] = {'name': f.filename,
                         'uploaded_at': datetime.now().strftime('%Y-%m-%d %H:%M')}
        save_master(data)
        return jsonify(ok=True, name=f.filename, data={
            'buyers': data.get('buyers', []),
            'sellers': data.get('sellers', []),
            'products': data.get('products', []),
            'forwarders': data.get('forwarders', []),
        })
    except Exception as e:
        return jsonify(ok=False, error=str(e))

@app.route('/buyer/add', methods=['POST'])
def buyer_add():
    req = request.json
    data = load_master()
    if not data: return jsonify(ok=False, error='请先上传主数据手册')
    data.setdefault('buyers', []).append({
        'code': req['code'], 'name_en': req['name_en'],
        'name_cn': req.get('name_cn', ''), 'address_en': req['address_en'],
        'payment_terms': req.get('payment_terms', ''),
        'incoterms': req.get('incoterms', 'FOB'),
        'currency': req.get('currency', 'USD'),
    })
    save_master(data)
    return jsonify(ok=True, data={'buyers': data['buyers'], 'sellers': data.get('sellers', []),
                                   'products': data.get('products', [])})

@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f: return jsonify(ok=False, error='没有文件')
    uid = uuid.uuid4().hex[:8]
    ext = os.path.splitext(f.filename)[1] or '.xlsx'
    path = os.path.join(UPLOAD_DIR, f'packing_{uid}{ext}')
    f.save(path)
    size = os.path.getsize(path)
    size_str = f'{size/1024:.1f} KB' if size < 1048576 else f'{size/1048576:.1f} MB'

    containers, product_codes, row_count = [], [], 0
    container_gw = {}
    try:
        from engine.parser_v2 import parse_packing_list_auto
        packing = parse_packing_list_auto(path)
        product_codes = packing.product_codes
        row_count = len(packing.rows)
        # 汇总每柜理论GW
        for r in packing.rows:
            if r.nw_per_box > 0:  # 用NW/pcs × qty
                container_gw[r.container_seq] = container_gw.get(r.container_seq, 0) + r.total_nw
        for seq, c in packing.containers.items():
            containers.append({'seq': seq, 'container_no': c.container_no,
                'seal_no': c.seal_no or '', 'container_size': c.container_size or '40HQ',
                'vgm_kg': c.vgm_kg, 'so_no': c.so_no or '',
                'etd': c.etd or '', 'port_loading': c.port_loading or 'NANSHA',
                'gw_theory': round(container_gw.get(seq, 0), 1),
                'tare': 3900, 'grp': 0})
    except Exception as e:
        pass
    return jsonify(ok=True, path=path, session=path, size=size_str,
                   containers=containers, container_count=len(containers),
                   row_count=row_count, product_codes=product_codes,
                   products=[{'code':c,'suits':0,'name':''} for c in product_codes],
                   rows=row_count,
                   container_gw=container_gw)

@app.route('/generate_v4', methods=['POST'])
def generate_v4():
    data = request.json
    # 兼容前端发来的 pack_session 或 packing_path
    packing_path = data.get('packing_path') or data.get('pack_session')
    batches = data.get('batches', [])
    master_path = get_master_xlsx_path()
    if not master_path: return jsonify(ok=False, error='请先上传主数据管理手册')
    if not packing_path or not os.path.exists(packing_path):
        return jsonify(ok=False, error='请先上传装箱清单（pack_session 无效）')
    task_id = uuid.uuid4().hex[:12]
    tasks[task_id] = {'status':'running','logs':[],'outputs':[],'_cur':0}
    threading.Thread(target=run_task_v4,
        args=(task_id, master_path, packing_path, batches), daemon=True).start()
    return jsonify(ok=True, task_id=task_id)

@app.route('/task/<tid>')
def task_status(tid):
    t = tasks.get(tid)
    if not t: return jsonify(status='not_found')
    cur = t['_cur']; new_logs = t['logs'][cur:]; t['_cur'] = len(t['logs'])
    return jsonify(status=t['status'], logs=new_logs, outputs=t['outputs'])

@app.route('/download/<path:fname>')
def download_file(fname):
    safe = os.path.basename(fname)
    path = os.path.join(OUTPUT_DIR, safe)
    if not os.path.exists(path): return '文件不存在', 404
    return send_file(path, as_attachment=True, download_name=safe)


def run_task_v4(tid, master_path, packing_path, batches):
    task = tasks[tid]
    def log(msg): task['logs'].append(msg)
    try:
        log('='*50); log('  外贸单据自动化系统 v4.0')
        log(f'  {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'); log('='*50)
        log('  📋 加载主数据管理手册…')
        from engine.config import MasterData, ShipmentMain, ShipmentContainer
        master = MasterData(master_path)
        log(f'  ✅ 成品档案：{len(master.finished_goods)} 条')
        log(f'  ✅ 物料价格：{len(master.component_prices)} 条')
        log(f'  ✅ 往来方：{len(master.parties)} 个')

        from engine.parser_v2 import parse_packing_list_auto
        from engine.parser import build_document_bundle
        from engine.build_v6 import generate_document_set
        log('  📦 解析装箱清单…')
        packing = parse_packing_list_auto(packing_path)
        log(f'     货柜：{len(packing.containers)} 个  明细：{len(packing.rows)} 行')

        outputs = []
        for batch_info in batches:
            bn = batch_info['batch_no']
            log(f'\n  ─── 批次：{bn}')
            seqs = batch_info.get('container_seqs', [])
            ctnr_details = {c['seq']: c for c in batch_info.get('container_details', [])}

            master.shipment_mains[bn] = ShipmentMain(
                batch_no=bn, customer_code=batch_info.get('customer_code',''),
                shipment_date=batch_info.get('shipment_date',''),
                invoice_no=batch_info.get('invoice_no',''),
                product_codes=batch_info.get('product_codes',[]),
                customs_suits=batch_info.get('customs_suits',[]),
                container_seq_str=','.join(str(s) for s in seqs),
                seller_code=batch_info.get('seller_code','QF-CN'), remark='',
            )
            for seq in seqs:
                cd = ctnr_details.get(seq, {})
                pc = packing.containers.get(seq)
                master.shipment_containers[seq] = ShipmentContainer(
                    seq=seq, batch_no=bn,
                    container_no=cd.get('container_no') or (pc.container_no if pc else ''),
                    seal_no=cd.get('seal_no',''),
                    container_size=cd.get('container_size','40HQ'),
                    vgm_kg=cd.get('vgm') or None,
                    so_no=cd.get('so_no',''),
                    etd=cd.get('etd',''),
                    port_loading=cd.get('port_loading','NANSHA') or 'NANSHA',
                    remark='',
                )
                # 更新packing里的GW（用网页选择的实际GW）
                gw_final = cd.get('gw_final', 0)
                if gw_final > 0:
                    rows_for_seq = [r for r in packing.rows if r.container_seq == seq]
                    if rows_for_seq:
                        # 按比例分配GW
                        total_nw = sum(r.total_nw for r in rows_for_seq)
                        if total_nw > 0:
                            for r in rows_for_seq:
                                ratio = r.total_nw / total_nw
                                r.gw_per_box = (gw_final * ratio / r.box_count) if r.box_count > 0 else 0

            try:
                bundle = build_document_bundle(master, bn, packing)
                log(f'  ✅ Invoice: {bundle.invoice_no}')
                log(f'     套一：USD {bundle.set1_total_amount:,.2f}')
                log(f'     套二：USD {bundle.set2_total_amount:,.2f}  ({len(bundle.set2_lines)} 部件)')
                log('  📄 生成 Excel…')
                paths = generate_document_set(bundle, OUTPUT_DIR, master_path=master_path)
                s1=os.path.basename(paths['set1']); s2=os.path.basename(paths['set2'])
                log(f'  ✅ 套一：{s1}'); log(f'  ✅ 套二：{s2}')
                outputs.append({'batch':bn,'set1':s1,'set1_name':s1,'set2':s2,'set2_name':s2})
            except Exception as e:
                log(f'  ✗ 异常：{e}'); log(traceback.format_exc()[:400])

        log(f'\n{"="*50}'); log(f'  ✅ 完成！{len(outputs)} 套'); log('='*50)
        task['outputs'] = outputs
        task['status'] = 'done' if outputs else 'error'
    except Exception as e:
        task['logs'].append(f'  ✗ 系统错误：{e}')
        task['logs'].append(traceback.format_exc()[:500])
        task['status'] = 'error'


if __name__ == '__main__':
    print('\n'+'='*56)
    print('  外贸单据自动生成系统 v4.0')
    print(f'  存储路径：{DATA_DIR}')
    print('='*56+'\n')
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
