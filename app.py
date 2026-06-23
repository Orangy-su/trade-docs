"""
app_v4.py — 外贸单据系统 v4
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
app.secret_key = os.environ.get('SECRET_KEY', 'qf-trade-v4-2026')

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
    import pandas as pd
    buyers, sellers, forwarders, products = [], [], [], []
    try:
        xl = pd.ExcelFile(path)
        for sh in xl.sheet_names:
            if '往来方' in sh:
                df = pd.read_excel(xl, sheet_name=sh, header=None)
                for _, row in df.iterrows():
                    def sv(i, d=''):
                        v = str(row[i]).strip() if i < len(row) and pd.notna(row[i]) else d
                        return d if v in ('nan','NaN') else v
                    code = sv(0); ptype = sv(1)
                    if not code or ptype not in ('SELLER','BUYER','FORWARDER'): continue
                    name_en = sv(2); name_cn = sv(3)
                    if '（' in name_en or not name_en: continue
                    if ptype == 'BUYER':
                        buyers.append({'code':code,'name_en':name_en,'name_cn':name_cn,
                            'address_en':sv(4),'payment_terms':sv(8),
                            'incoterms':sv(9,'FOB'),'currency':sv(10,'USD')})
                    elif ptype == 'SELLER':
                        sellers.append({'code':code,'name_en':name_en,'name_cn':name_cn,
                            'address_en':sv(4),'tax_id':sv(11),'port_loading':sv(12)})
                    elif ptype == 'FORWARDER':
                        forwarders.append({'code':code,'name_en':name_en,
                            'address_en':sv(4),'notify_name':sv(5),
                            'notify_address':sv(6),'contact':sv(7)})
            elif '成品档案' in sh:
                df = pd.read_excel(xl, sheet_name=sh, header=None)
                seen = set()
                for _, row in df.iterrows():
                    def sv(i, d=''):
                        v = str(row[i]).strip() if i < len(row) and pd.notna(row[i]) else d
                        return d if v in ('nan','NaN') else v
                    code = sv(0); name_cn = sv(1); name_en = sv(2)
                    # 过滤表头、说明行、占位符
                    if not code: continue
                    if any(x in code for x in ['成品编码','★','（','待填','▌','Sheet','填写','|','说明','基础信息','产品类']): continue
                    if len(code) < 5: continue  # 有效成品编码至少5位
                    if code in seen: continue
                    seen.add(code)
                    products.append({'code':code,'name_cn':name_cn,'name_en':name_en})
    except Exception as e:
        pass
    return {'buyers':buyers,'sellers':sellers,'forwarders':forwarders,'products':products}


HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>外贸单据系统 · 擎烽</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;700&family=JetBrains+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
:root{
  --navy:#1F3864;--blue:#2E75B6;--lblue:#D6E4F0;--teal:#0D7377;
  --orange:#E67E22;--green:#27AE60;--red:#C0392B;--purple:#6C3483;
  --bg:#EEF2F7;--card:#fff;--border:#DDE3EC;--text:#2C3E50;
  --gray:#7F8C8D;--lgray:#F5F7FA;
  --mono:'JetBrains Mono',monospace;--sans:'Noto Sans SC',sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh}
.topbar{background:linear-gradient(135deg,var(--navy),#2a4a82);padding:0 2rem;height:58px;
  display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:200;box-shadow:0 2px 20px rgba(31,56,100,.25)}
.brand{display:flex;align-items:center;gap:12px}
.brand-logo{width:36px;height:36px;background:var(--teal);border-radius:9px;
  display:flex;align-items:center;justify-content:center;font-size:18px}
.brand-name{font-size:15px;font-weight:700;color:#fff}
.brand-sub{font-size:11px;color:rgba(255,255,255,.55);margin-top:1px}
.topbar-right{display:flex;align-items:center;gap:10px}
.master-pill{display:flex;align-items:center;gap:7px;background:rgba(255,255,255,.1);
  border:1px solid rgba(255,255,255,.15);border-radius:20px;padding:5px 12px;
  font-size:12px;color:rgba(255,255,255,.85);cursor:pointer;transition:.2s}
.master-pill:hover{background:rgba(255,255,255,.18)}
.mdot{width:7px;height:7px;border-radius:50%;background:#6E7681}
.mdot.ok{background:#3FB950}
.mdot.no{background:#D29922;animation:blink 1.5s infinite}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.4}}
.ver{font-family:var(--mono);font-size:11px;color:rgba(255,255,255,.45);
  background:rgba(255,255,255,.07);padding:3px 8px;border-radius:4px}
.main{max-width:980px;margin:0 auto;padding:2rem 1.5rem}
@media(max-width:640px){.main{padding:1rem}}
.card{background:var(--card);border-radius:14px;border:1px solid var(--border);
  padding:1.5rem;margin-bottom:1.25rem;box-shadow:0 1px 6px rgba(0,0,0,.05)}
.card-hdr{display:flex;align-items:center;gap:10px;margin-bottom:1.1rem;
  padding-bottom:.8rem;border-bottom:1px solid var(--border)}
.card-icon{width:34px;height:34px;border-radius:8px;
  display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0}
.card-title{font-size:15px;font-weight:700;color:var(--navy)}
.card-sub{font-size:12px;color:var(--gray);margin-top:2px}
.upload-zone{border:2px dashed var(--border);border-radius:10px;padding:1.75rem;
  text-align:center;cursor:pointer;transition:.2s;background:#FAFCFF;position:relative}
.upload-zone:hover,.upload-zone.drag{border-color:var(--blue);background:#EBF3FB}
.upload-zone.has{border-color:var(--green);background:#F0FBF4;border-style:solid}
.upload-zone input[type=file]{position:absolute;inset:0;opacity:0;
  cursor:pointer;width:100%;height:100%}
.up-icon{font-size:2.2rem;margin-bottom:.5rem}
.up-label{font-size:13px;color:var(--gray)}.up-label strong{color:var(--blue)}
.up-ok{font-size:13px;color:var(--green);font-weight:600;
  display:flex;align-items:center;justify-content:center;gap:7px}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
@media(max-width:580px){.form-grid{grid-template-columns:1fr}}
.form-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
@media(max-width:640px){.form-grid-3{grid-template-columns:1fr 1fr}}
.field{display:flex;flex-direction:column;gap:5px}
.field label{font-size:12px;font-weight:600;color:var(--navy)}
.field input,.field select,.field textarea{
  padding:10px 13px;border:1.5px solid var(--border);border-radius:8px;
  font-family:var(--sans);font-size:13px;color:var(--text);
  background:#fff;outline:none;transition:.2s}
.field input:focus,.field select:focus{border-color:var(--blue);
  box-shadow:0 0 0 3px rgba(46,117,182,.1)}
.field .hint{font-size:11px;color:var(--gray)}
.btn{padding:10px 18px;border-radius:8px;font-family:var(--sans);font-size:13px;
  font-weight:600;cursor:pointer;border:none;transition:all .2s;
  display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn-blue{background:var(--blue);color:#fff}
.btn-blue:hover:not(:disabled){background:var(--navy);transform:translateY(-1px)}
.btn-blue:disabled{background:#B0C4DE;cursor:not-allowed}
.btn-teal{background:var(--teal);color:#fff}
.btn-teal:hover:not(:disabled){background:#0A5F63}
.btn-green{background:var(--green);color:#fff}
.btn-green:hover{background:#219A52}
.btn-outline{background:#fff;color:var(--blue);border:1.5px solid var(--blue)}
.btn-outline:hover{background:var(--lblue)}
.btn-outline-red{background:#fff;color:var(--red);border:1.5px solid var(--red)}
.btn-ghost{background:transparent;color:var(--gray);border:1.5px solid var(--border)}
.btn-ghost:hover{background:var(--lgray)}
.btn-sm{padding:6px 12px;font-size:12px;border-radius:6px}
.btn-lg{padding:13px 28px;font-size:15px;border-radius:10px}
.btn-block{width:100%;justify-content:center}
.tabs{display:flex;gap:0;border-bottom:2px solid var(--border);margin-bottom:1.25rem}
.tab{padding:9px 18px;font-size:13px;font-weight:600;color:var(--gray);
  cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-2px;transition:.2s}
.tab.active{color:var(--blue);border-bottom-color:var(--blue)}
.tab-panel{display:none}.tab-panel.active{display:block}
.divider{height:1px;background:var(--border);margin:1.25rem 0}
/* 货柜表格 */
.ctnr-table{width:100%;border-collapse:collapse;font-size:12px}
.ctnr-table th{background:var(--navy);color:#fff;padding:8px 10px;font-size:11px;font-weight:600;text-align:left}
.ctnr-table td{padding:6px 8px;border-bottom:1px solid var(--border);vertical-align:middle}
.ctnr-table tr:last-child td{border-bottom:none}
.ctnr-table tr:hover td{background:#F8FAFF}
.mini-input{padding:5px 8px;border:1.5px solid var(--border);border-radius:6px;
  font-size:12px;width:100%;outline:none;transition:.2s;background:#fff}
.mini-input:focus{border-color:var(--blue)}
.group-btn{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;
  border-radius:12px;font-size:11px;font-weight:700;cursor:pointer;border:none;transition:.2s}
.g0{background:#E8F4FD;color:#1A5276;border:1.5px solid #2E75B6}
.g1{background:#E9F7EF;color:#1E8449;border:1.5px solid #27AE60}
.g2{background:#FEF9E7;color:#7D6608;border:1.5px solid #D4AC0D}
.g3{background:#F5EEF8;color:#512E5F;border:1.5px solid #8E44AD}
.sn0{background:#2E75B6;color:#fff}.sn1{background:#27AE60;color:#fff}
.sn2{background:#D4AC0D;color:#fff}.sn3{background:#8E44AD;color:#fff}
.seq-num{width:26px;height:26px;border-radius:50%;display:flex;
  align-items:center;justify-content:center;font-weight:700;font-size:12px;flex-shrink:0}
/* 进度/结果 */
.prog-card{display:none}.prog-card.show{display:block}
.res-card{display:none}.res-card.show{display:block}
.terminal{background:#0D1117;border-radius:10px;padding:1rem 1.25rem;
  font-family:var(--mono);font-size:12px;line-height:1.9;
  max-height:280px;overflow-y:auto;color:#C9D1D9}
.ll{display:flex;gap:10px;align-items:baseline}
.lt{color:#484F58;flex-shrink:0;font-size:11px}
.lo{color:#3FB950}.lw{color:#D29922}.le{color:#F85149}.li{color:#58A6FF}.ld{color:#6E7681}
.sdot{width:9px;height:9px;border-radius:50%;background:#6E7681;flex-shrink:0}
.sdot.run{background:var(--orange);animation:pulse 1s infinite}
.sdot.ok{background:var(--green)}.sdot.err{background:var(--red)}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(.8)}}
.res-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:1rem}
@media(max-width:540px){.res-grid{grid-template-columns:1fr}}
.rf{border:1.5px solid var(--border);border-radius:10px;padding:1rem;
  display:flex;flex-direction:column;gap:7px}
.rf.s1{border-color:var(--blue);background:linear-gradient(135deg,#F5FAFF,#EBF3FB)}
.rf.s2{border-color:var(--green);background:linear-gradient(135deg,#F5FDF7,#EAFAF1)}
.rf-tag{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px}
.s1 .rf-tag{color:var(--blue)}.s2 .rf-tag{color:var(--green)}
.rf-name{font-size:12px;font-family:var(--mono);word-break:break-all;color:var(--text)}
/* Modal */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);
  z-index:300;align-items:center;justify-content:center}
.modal-bg.show{display:flex}
.modal{background:#fff;border-radius:16px;padding:2rem;width:100%;max-width:600px;
  margin:1rem;box-shadow:0 20px 60px rgba(0,0,0,.2);animation:slideUp .25s ease}
.modal-lg{max-width:760px}
@keyframes slideUp{from{transform:translateY(16px);opacity:0}to{transform:none;opacity:1}}
.modal-title{font-size:17px;font-weight:700;color:var(--navy);margin-bottom:.35rem}
.modal-sub{font-size:13px;color:var(--gray);margin-bottom:1.4rem}
.modal-footer{display:flex;gap:10px;justify-content:flex-end;margin-top:1.5rem}
/* VGM表格 */
.vgm-table{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:1rem}
.vgm-table th{background:var(--navy);color:#fff;padding:9px 12px;font-size:12px}
.vgm-table td{padding:8px 12px;border-bottom:1px solid var(--border)}
.vgm-table tr:last-child td{border-bottom:none}
.alert-box{padding:10px 14px;border-radius:8px;font-size:12px;margin-top:12px}
.alert-warn{background:#FFF3CD;border-left:3px solid var(--orange);color:#7D6608}
.alert-info{background:var(--lblue);border-left:3px solid var(--blue);color:var(--navy)}
/* 历史 */
.hist-item{display:flex;align-items:center;justify-content:space-between;
  padding:10px 14px;background:var(--lgray);border-radius:8px;
  border:1px solid var(--border);gap:8px;flex-wrap:wrap;margin-bottom:8px}
</style>
</head>
<body>
<nav class="topbar">
  <div class="brand">
    <div class="brand-logo">📦</div>
    <div>
      <div class="brand-name">外贸单据自动生成系统</div>
      <div class="brand-sub">广东擎烽电气科技有限公司 · Trade Document Automation</div>
    </div>
  </div>
  <div class="topbar-right">
    <div class="master-pill" onclick="openMasterModal()">
      <div class="mdot no" id="mdot"></div>
      <span id="masterTxt">主数据未上传</span>
    </div>
    <span class="ver">v4.0</span>
  </div>
</nav>

<main class="main">

<!-- 主数据状态卡 -->
<div class="card" id="masterCard">
  <div class="card-hdr">
    <div class="card-icon" style="background:#E8F5E9">📋</div>
    <div>
      <div class="card-title">主数据管理手册</div>
      <div class="card-sub">上传一次长期有效 · 含成品档案/物料价格/往来方</div>
    </div>
    <div style="margin-left:auto">
      <button class="btn btn-teal btn-sm" onclick="openMasterModal()">⬆ 上传/更新</button>
    </div>
  </div>
  <div id="masterStatus" style="padding:10px 14px;background:var(--lgray);border-radius:8px;
    font-size:12px;color:var(--gray);border:1px dashed var(--border)">
    尚未上传主数据手册，请点击右上角「上传/更新」
  </div>
</div>

<!-- Step 1: 上传装箱清单 -->
<div class="card">
  <div class="card-hdr">
    <div class="card-icon" style="background:#FEF9E7">📦</div>
    <div>
      <div class="card-title">Step 1 · 上传装箱清单</div>
      <div class="card-sub">仓管员填写后提交，支持新旧两种格式</div>
    </div>
  </div>
  <div class="upload-zone" id="packZone">
    <input type="file" id="packFile" accept=".xlsx,.xls" onchange="handlePack(this)">
    <div class="up-icon">🗂️</div>
    <div class="up-label">拖拽文件到此处，或 <strong>点击选择</strong></div>
    <div style="font-size:11px;color:var(--gray);margin-top:4px">支持 .xlsx · .xls · 任意文件名</div>
  </div>
</div>

<!-- Step 2: 出货信息 -->
<div class="card" id="orderCard" style="display:none">
  <div class="card-hdr">
    <div class="card-icon" style="background:#EBF3FB">📋</div>
    <div>
      <div class="card-title">Step 2 · 填写出货信息</div>
      <div class="card-sub">选择客户、填写发票信息、配置货柜</div>
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
    <!-- 成品套数动态区 -->
    <div id="productSuitsArea" style="display:none;margin-bottom:14px">
      <div style="font-size:12px;font-weight:600;color:var(--navy);margin-bottom:8px">
        报关整机套数 ★
        <span style="font-weight:400;color:var(--gray);font-size:11px">从装箱清单自动识别，每个成品填对应套数</span>
      </div>
      <div id="productSuitsGrid" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:10px"></div>
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
    <div class="alert-box alert-warn" style="margin-bottom:12px">
      <strong>货柜分组：</strong>同色=合并出一套单据 · 点色块切换分组<br>
      <strong>货柜信息：</strong>铅封号、SO号、ETD为SI必填项 · 过磅重量用于VGM计算
    </div>
    <div style="overflow-x:auto">
      <table class="ctnr-table" id="ctnrTable">
        <thead>
          <tr>
            <th style="width:44px">序号</th>
            <th style="width:130px">货柜号</th>
            <th style="width:100px">铅封号 ★</th>
            <th style="width:80px">柜型</th>
            <th style="width:130px">SO号 ★</th>
            <th style="width:90px">ETD</th>
            <th style="width:90px">过磅GW kg<br><span style="font-weight:400;font-size:10px">实测重量</span></th>
            <th style="width:90px">理论GW kg<br><span style="font-weight:400;font-size:10px">自动汇总</span></th>
            <th style="width:85px">柜皮重 kg<br><span style="font-weight:400;font-size:10px">柜门标注值</span></th>
            <th style="width:110px">单据分组</th>
          </tr>
        </thead>
        <tbody id="ctnrBody">
          <tr><td colspan="9" style="text-align:center;color:var(--gray);padding:2rem">
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

  <div class="divider"></div>
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
    ⚠️ <strong>注意：</strong>套二PL的毛重数据来自装箱清单理论重量。
    如果SI使用的是过磅重量，请在下载后手动更新PL对应行的GW/NW数值，以保持与SI一致。
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

<!-- ═══ 主数据 Modal ═══ -->
<div class="modal-bg" id="masterModal">
  <div class="modal">
    <div class="modal-title">上传主数据管理手册</div>
    <div class="modal-sub">系统解析后保存到服务器，后续打开页面自动加载，无需重新上传。如有更新重新上传即可覆盖。</div>
    <div class="upload-zone" id="masterZone" style="padding:1.4rem">
      <input type="file" id="masterFile" accept=".xlsx,.xls" onchange="handleMaster(this)">
      <div class="up-icon" id="mzIcon">📄</div>
      <div id="mzText">
        <div class="up-label">拖拽主数据管理手册到此<br>或 <strong>点击选择</strong></div>
        <div style="font-size:11px;color:var(--gray);margin-top:4px">任意文件名均可</div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeMasterModal()">关闭</button>
    </div>
  </div>
</div>

<!-- ═══ VGM确认 Modal ═══ -->
<div class="modal-bg" id="vgmModal">
  <div class="modal modal-lg">
    <div class="modal-title">⚖️ 确认SI重量 & VGM</div>
    <div class="modal-sub">请选择每个货柜SI使用的重量数据，VGM = 选定GW + 柜皮重（自动带入）。</div>
    <table class="vgm-table">
      <thead>
        <tr>
          <th>货柜</th>
          <th>理论GW(kg)<br>装箱单汇总</th>
          <th>过磅GW(kg)<br>实测重量</th>
          <th>SI使用</th>
          <th>柜皮重(kg)</th>
          <th>VGM(kg)</th>
        </tr>
      </thead>
      <tbody id="vgmBody"></tbody>
    </table>
    <div class="alert-box alert-warn">
      ⚠️ 若SI选用过磅重量，与装箱单理论重量有差异，生成后请手动更新PL的GW/NW数据。
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
    <div class="modal-sub">填写后自动保存到主数据手册，下次直接下拉选择。</div>
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
        <select id="nb_inco"><option>FOB</option><option>CIF</option></select>
      </div>
      <div class="field">
        <label>货币</label>
        <select id="nb_currency"><option>USD</option><option>CNY</option></select>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeNewBuyerModal()">取消</button>
      <button class="btn btn-teal" onclick="saveNewBuyer()">💾 保存</button>
    </div>
  </div>
</div>

<script>
const S = {
  masterLoaded:false, masterName:'', masterData:null,
  packingPath:null, containers:[], productCodes:[],
  containerGroups:{},  // seq -> groupIndex
  containerGwTheory:{}, // seq -> 理论GW
  vgmData:{},          // seq -> {useWeigh, gwFinal, vgm}
  history:[],
};
const GRP=['A','B','C','D'];
const GRP_CLS=['g0','g1','g2','g3'];
const SN_CLS=['sn0','sn1','sn2','sn3'];
const TARE={20:{GP:2200},40:{GP:2300,HQ:3900},45:{HQ:4200}};

function getTareDefault(sizeStr){
  // 返回默认柜皮重，用户可在表格里修改
  const m=sizeStr&&sizeStr.match(/(\d+)(GP|HQ)/i);
  if(!m)return 3900;
  const defaults={'20GP':2200,'40GP':2300,'40HQ':3900,'45HQ':4200};
  return defaults[(m[1]+m[2].toUpperCase())]||3900;
}
function getTare(sizeStr){
  const m=sizeStr&&sizeStr.match(/(\d+)(GP|HQ)/i);
  if(!m)return 3900;
  const sz=TARE[m[1]];
  return sz&&sz[m[2].toUpperCase()]||3900;
}

// ── 初始化 ──────────────────────────────────────────────────
async function init(){
  document.getElementById('shipDate').value=new Date().toISOString().split('T')[0];
  try{
    const r=await fetch('/master/info');
    const d=await r.json();
    if(d.ok){S.masterLoaded=true;S.masterName=d.name||'主数据已就绪';S.masterData=d.data;updateMasterUI(true);populateSelects();}
  }catch(e){}
}
init();

// ── 主数据 ──────────────────────────────────────────────────
function openMasterModal(){document.getElementById('masterModal').classList.add('show')}
function closeMasterModal(){document.getElementById('masterModal').classList.remove('show')}
document.getElementById('masterModal').addEventListener('click',e=>{if(e.target===e.currentTarget)closeMasterModal()});

async function handleMaster(input){
  const file=input.files[0];if(!file)return;
  document.getElementById('mzIcon').textContent='⏳';
  document.getElementById('mzText').innerHTML=`<div class="up-label">正在上传并解析 ${file.name}…</div>`;
  const fd=new FormData();fd.append('file',file);
  try{
    const r=await fetch('/master/upload',{method:'POST',body:fd});
    const d=await r.json();
    if(d.ok){
      S.masterLoaded=true;S.masterName=d.name;S.masterData=d.data;
      updateMasterUI(true);populateSelects();closeMasterModal();
    }else{
      document.getElementById('mzIcon').textContent='❌';
      document.getElementById('mzText').innerHTML=`<div style="color:var(--red);font-size:13px">解析失败：${d.error||'请检查文件格式'}</div>`;
    }
  }catch(e){
    document.getElementById('mzIcon').textContent='❌';
    document.getElementById('mzText').innerHTML=`<div style="color:var(--red)">上传失败，请重试</div>`;
  }
}

function updateMasterUI(loaded){
  document.getElementById('mdot').className='mdot '+(loaded?'ok':'no');
  document.getElementById('masterTxt').textContent=loaded?S.masterName:'主数据未上传';
  const st=document.getElementById('masterStatus');
  if(loaded){
    const d=S.masterData||{};
    st.innerHTML=`✅ 已加载 · 买方 ${(d.buyers||[]).length} 个 · 卖方 ${(d.sellers||[]).length} 个 · 成品 ${(d.products||[]).length} 个`;
    st.style.cssText='padding:10px 14px;background:#E9F7EF;border-radius:8px;font-size:12px;color:#1E8449;border:1px solid #A3D9B1';
  }else{
    st.innerHTML='尚未上传主数据手册，请点击右上角「上传/更新」';
    st.style.cssText='padding:10px 14px;background:var(--lgray);border-radius:8px;font-size:12px;color:var(--gray);border:1px dashed var(--border)';
  }
  updateGenBtn();
}

function populateSelects(){
  if(!S.masterData)return;
  const bs=document.getElementById('buyerSel');
  const ss=document.getElementById('sellerSel');
  const prev_b=bs.value,prev_s=ss.value;
  bs.innerHTML='<option value="">— 选择客户 —</option>';
  (S.masterData.buyers||[]).forEach(b=>{
    bs.innerHTML+=`<option value="${b.code}" data-payment="${b.payment_terms||''}"
      data-inco="${b.incoterms||'FOB'}" data-currency="${b.currency||'USD'}"
      data-addr="${esc(b.address_en||'')}">${b.code} · ${(b.name_en||'').substring(0,35)}</option>`;
  });
  bs.innerHTML+=`<option value="__new__">＋ 新增买方客户…</option>`;
  if(prev_b)bs.value=prev_b;
  ss.innerHTML='<option value="">— 选择卖方 —</option>';
  (S.masterData.sellers||[]).forEach(s=>{
    ss.innerHTML+=`<option value="${s.code}">${s.code} · ${s.name_cn||s.name_en||''}</option>`;
  });
  if(prev_s)ss.value=prev_s;
  else if((S.masterData.sellers||[]).length===1)ss.value=S.masterData.sellers[0].code;
}

function updateBuyer(){
  const sel=document.getElementById('buyerSel');
  if(sel.value==='__new__'){sel.value='';openNewBuyerModal();return;}
  const opt=sel.selectedOptions[0];
  if(!opt||!opt.value){document.getElementById('buyerInfoBar').style.display='none';return;}
  document.getElementById('payTerms').value=opt.dataset.payment||'';
  document.getElementById('incoterms').value=opt.dataset.inco||'FOB';
  document.getElementById('currency').value=opt.dataset.currency||'USD';
  const bar=document.getElementById('buyerInfoBar');
  bar.style.display='flex';bar.style.gap='16px';bar.style.flexWrap='wrap';
  bar.innerHTML=`<span>💳 ${opt.dataset.payment||'—'}</span><span>🚢 ${opt.dataset.inco||'FOB'}</span><span>💵 ${opt.dataset.currency||'USD'}</span>`;
  updateGenBtn();
}

// ── 装箱清单 ────────────────────────────────────────────────
async function handlePack(input){
  const file=input.files[0];if(!file)return;
  const zone=document.getElementById('packZone');
  zone.innerHTML=`<div style="font-size:13px;color:var(--gray)">⏳ 正在上传并解析 ${file.name}…</div>
    <input type="file" accept=".xlsx,.xls" onchange="handlePack(this)"
      style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">`;
  const fd=new FormData();fd.append('file',file);fd.append('type','packing');
  try{
    const r=await fetch('/upload',{method:'POST',body:fd});
    const d=await r.json();
    if(d.ok){
      S.packingPath=d.path;S.containers=d.containers||[];S.productCodes=d.product_codes||[];
      S.containerGwTheory=d.container_gw||{};
      zone.className='upload-zone has';
      zone.innerHTML=`<div class="up-ok">✅ ${file.name}</div>
        <div style="font-size:11px;color:var(--gray);margin-top:4px">
          ${d.size} · ${d.container_count} 个货柜 · ${d.row_count} 行明细
        </div>
        <input type="file" accept=".xlsx,.xls" onchange="handlePack(this)"
          style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">`;
      // 初始化货柜分组
      S.containers.forEach((c,i)=>{if(S.containerGroups[c.seq]===undefined)S.containerGroups[c.seq]=i;});
      renderContainerTable();renderProductSuits();
      document.getElementById('orderCard').style.display='block';
      document.getElementById('orderCard').scrollIntoView({behavior:'smooth',block:'nearest'});
    }else{
      zone.innerHTML=`<div style="color:var(--red);font-size:13px">❌ ${d.error||'解析失败'}</div>
        <input type="file" accept=".xlsx,.xls" onchange="handlePack(this)"
          style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">`;
    }
  }catch(e){
    zone.innerHTML=`<div style="color:var(--red);font-size:13px">❌ 网络错误</div>
      <input type="file" accept=".xlsx,.xls" onchange="handlePack(this)"
        style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">`;
  }
  updateGenBtn();
}

// ── 成品套数动态渲染 ─────────────────────────────────────────
function renderProductSuits(){
  const area=document.getElementById('productSuitsArea');
  const empty=document.getElementById('productSuitsEmpty');
  const grid=document.getElementById('productSuitsGrid');
  if(!S.productCodes||!S.productCodes.length){area.style.display='none';empty.style.display='block';return;}
  empty.style.display='none';area.style.display='block';
  grid.innerHTML=S.productCodes.map(pc=>{
    const prod=(S.masterData&&S.masterData.products||[]).find(p=>p.code===pc);
    const label=prod?`${pc}<br><span style="font-size:10px;color:var(--gray);font-weight:400">${prod.name_cn}</span>`:pc;
    return `<div class="field" style="background:#F8FAFF;padding:10px 12px;border-radius:8px;border:1px solid var(--border)">
      <label style="font-family:var(--mono);font-size:11px;line-height:1.4">${label}</label>
      <input type="number" class="field-input" id="suits_${pc}"
        placeholder="整机套数" min="0" style="margin-top:6px;padding:8px 12px;border:1.5px solid var(--border);border-radius:6px;width:100%;font-size:13px;outline:none"
        oninput="updateGenBtn()">
    </div>`;
  }).join('');
}

// ── 货柜表格 ─────────────────────────────────────────────────
function renderContainerTable(){
  const tbody=document.getElementById('ctnrBody');
  if(!S.containers.length){
    tbody.innerHTML='<tr><td colspan="9" style="text-align:center;color:var(--gray);padding:2rem">未解析到货柜信息</td></tr>';
    return;
  }
  tbody.innerHTML=S.containers.map(c=>{
    const gi=S.containerGroups[c.seq]||0;
    const gc=GRP_CLS[gi%GRP_CLS.length];
    const snc=SN_CLS[gi%SN_CLS.length];
    const gname=GRP[gi%GRP.length];
    const gwT=S.containerGwTheory[c.seq]||0;
    return `<tr>
      <td><div class="seq-num ${snc}">${c.seq}</div></td>
      <td><span style="font-family:var(--mono);font-size:11px">${c.container_no||'—'}</span></td>
      <td><input class="mini-input" id="seal_${c.seq}" value="${c.seal_no||''}" placeholder="铅封号"></td>
      <td>
        <select class="mini-input" id="csize_${c.seq}" style="padding:4px 6px">
          ${['20GP','40GP','40HQ','45HQ'].map(s=>`<option${s===(c.container_size||'40HQ')?' selected':''}>${s}</option>`).join('')}
        </select>
      </td>
      <td><input class="mini-input" id="so_${c.seq}" value="${c.so_no||''}" placeholder="SO号" style="font-family:var(--mono);font-size:11px"></td>
      <td><input class="mini-input" id="etd_${c.seq}" type="date" value="${c.etd||''}"></td>
      <td><input class="mini-input" id="weigh_${c.seq}" type="number" placeholder="过磅重量" style="width:85px"></td>
      <td style="text-align:center;font-family:var(--mono);font-size:12px;color:var(--gray)">
        ${gwT>0?gwT.toFixed(0):'—'}
      </td>
      <td><input class="mini-input" id="tare_${c.seq}" type="number"
        value="${getTareDefault(c.container_size||'40HQ')}"
        placeholder="柜皮重" style="width:80px"></td>
      <td>
        <button class="group-btn ${gc}" onclick="cycleGroup(${c.seq})" id="gbtn_${c.seq}">
          ● 分组${gname}
        </button>
      </td>
    </tr>`;
  }).join('');
}

function cycleGroup(seq){
  const allGs=new Set(Object.values(S.containerGroups));
  const cur=S.containerGroups[seq]||0;
  const maxG=Math.max(...allGs);
  S.containerGroups[seq]=cur>=maxG&&maxG<3?cur+1:(cur+1)%(maxG+2===1?2:maxG+2);
  if(S.containerGroups[seq]>3)S.containerGroups[seq]=0;
  renderContainerTable();
}

// ── 页签切换 ─────────────────────────────────────────────────
function switchTab(name){
  ['basic','containers','advanced'].forEach((n,i)=>{
    document.querySelectorAll('.tab')[i].className='tab'+(n===name?' active':'');
    document.getElementById('tab-'+n).className='tab-panel'+(n===name?' active':'');
  });
}

// ── 新增买方 ─────────────────────────────────────────────────
function openNewBuyerModal(){document.getElementById('newBuyerModal').classList.add('show')}
function closeNewBuyerModal(){document.getElementById('newBuyerModal').classList.remove('show')}
document.getElementById('newBuyerModal').addEventListener('click',e=>{if(e.target===e.currentTarget)closeNewBuyerModal()});
async function saveNewBuyer(){
  const data={code:document.getElementById('nb_code').value.trim(),
    name_en:document.getElementById('nb_name_en').value.trim(),
    address_en:document.getElementById('nb_addr').value.trim(),
    payment_terms:document.getElementById('nb_payment').value.trim(),
    incoterms:document.getElementById('nb_inco').value,
    currency:document.getElementById('nb_currency').value};
  if(!data.code||!data.name_en||!data.address_en){alert('客户代码、英文全称、地址为必填项');return;}
  const r=await fetch('/buyer/add',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  const d=await r.json();
  if(d.ok){if(d.data){S.masterData=d.data;populateSelects();}
    document.getElementById('buyerSel').value=data.code;updateBuyer();closeNewBuyerModal();
    ['nb_code','nb_name_en','nb_addr','nb_payment'].forEach(id=>document.getElementById(id).value='');
  }else alert('保存失败：'+d.error);
}

// ── 按钮状态 ─────────────────────────────────────────────────
function updateGenBtn(){
  const hasSuits=S.productCodes&&S.productCodes.length>0&&
    S.productCodes.every(pc=>{const el=document.getElementById('suits_'+pc);return el&&el.value.trim()!==''&&parseInt(el.value)>0;});
  const ready=S.masterLoaded&&S.packingPath&&
    document.getElementById('buyerSel').value&&
    document.getElementById('sellerSel').value&&
    document.getElementById('invoiceNo').value.trim()&&
    hasSuits;
  document.getElementById('genBtn').disabled=!ready;
  document.getElementById('readyHint').textContent=ready?'准备就绪，点击生成':'请完成上方必填项（★）';
  document.getElementById('readyHint').style.color=ready?'var(--green)':'var(--gray)';
}
['buyerSel','sellerSel','invoiceNo','shipDate'].forEach(id=>{
  const el=document.getElementById(id);if(el){el.addEventListener('input',updateGenBtn);el.addEventListener('change',updateGenBtn);}
});

// ── VGM确认弹窗 ──────────────────────────────────────────────
function preGenerate(){
  // 收集货柜信息
  const groups={};
  Object.entries(S.containerGroups).forEach(([seq,gi])=>{
    if(!groups[gi])groups[gi]=[];groups[gi].push(parseInt(seq));
  });
  // 检查是否有过磅重量填写
  const hasWeigh=S.containers.some(c=>{
    const el=document.getElementById('weigh_'+c.seq);
    return el&&el.value.trim()!=='';
  });
  if(hasWeigh){
    // 显示VGM弹窗
    buildVgmModal();
    document.getElementById('vgmModal').classList.add('show');
  }else{
    // 没有过磅重量，直接用理论GW
    S.containers.forEach(c=>{
      const gwT=S.containerGwTheory[c.seq]||0;
      const csizeP=document.getElementById('csize_'+c.seq)?.value||'40HQ';
      const tareElP=document.getElementById('tare_'+c.seq);
      const tare=tareElP&&tareElP.value?parseFloat(tareElP.value):getTareDefault(csizeP);
      S.vgmData[c.seq]={useWeigh:false,gwFinal:gwT,vgm:Math.round(gwT+tare)};
    });
    document.getElementById('gwReminder').style.display='none';
    doGenerate();
  }
}

function buildVgmModal(){
  const tbody=document.getElementById('vgmBody');
  tbody.innerHTML=S.containers.map(c=>{
    const gwT=S.containerGwTheory[c.seq]||0;
    const weighEl=document.getElementById('weigh_'+c.seq);
    const gwW=weighEl&&weighEl.value?parseFloat(weighEl.value):null;
    const csize=document.getElementById('csize_'+c.seq)?.value||'40HQ';
    // 从货柜配置表里读实际填写的柜皮重
    const tareEl=document.getElementById('tare_'+c.seq);
    const tare=tareEl&&tareEl.value?parseFloat(tareEl.value):getTareDefault(csize);
    const hasWeigh=gwW!==null&&gwW>0;
    return `<tr>
      <td style="font-family:var(--mono);font-size:12px">${c.container_no||'柜'+c.seq}</td>
      <td style="text-align:center">${gwT>0?gwT.toFixed(0):'—'}</td>
      <td style="text-align:center;font-weight:600;color:${hasWeigh?'var(--teal)':'var(--gray)'}">${hasWeigh?gwW.toFixed(0):'未填写'}</td>
      <td>
        <select id="vgm_choice_${c.seq}" class="mini-input" onchange="updateVgmCalc(${c.seq})" style="width:100%">
          ${gwT>0?`<option value="theory">理论GW（${gwT.toFixed(0)} kg）</option>`:''}
          ${hasWeigh?`<option value="weigh"${!gwT>0?' selected':''}>过磅GW（${gwW.toFixed(0)} kg）</option>`:''}
        </select>
        ${hasWeigh&&gwT>0?`<div style="font-size:10px;color:var(--orange);margin-top:3px">差额：${(gwW-gwT).toFixed(0)} kg</div>`:''}
      </td>
      <td style="text-align:center;color:var(--gray)">${tare}</td>
      <td style="text-align:center;font-weight:700;font-family:var(--mono)" id="vgm_result_${c.seq}">
        ${Math.round((hasWeigh?gwW:gwT)+tare)}
      </td>
    </tr>`;
  }).join('');
}

function updateVgmCalc(seq){
  const choice=document.getElementById('vgm_choice_'+seq)?.value;
  const gwT=S.containerGwTheory[seq]||0;
  const weighEl=document.getElementById('weigh_'+seq);
  const gwW=weighEl&&weighEl.value?parseFloat(weighEl.value):0;
  const csize=document.getElementById('csize_'+seq)?.value||'40HQ';
  const tareEl2=document.getElementById('tare_'+seq);
  const tare=tareEl2&&tareEl2.value?parseFloat(tareEl2.value):getTareDefault(csize);
  const gw=choice==='weigh'?gwW:gwT;
  const vgm=Math.round(gw+tare);
  const el=document.getElementById('vgm_result_'+seq);
  if(el)el.textContent=vgm;
}

function closeVgmModal(){document.getElementById('vgmModal').classList.remove('show')}
document.getElementById('vgmModal').addEventListener('click',e=>{if(e.target===e.currentTarget)closeVgmModal()});

function confirmVgm(){
  let usedWeigh=false;
  S.containers.forEach(c=>{
    const choice=document.getElementById('vgm_choice_'+c.seq)?.value||'theory';
    const gwT=S.containerGwTheory[c.seq]||0;
    const weighEl=document.getElementById('weigh_'+c.seq);
    const gwW=weighEl&&weighEl.value?parseFloat(weighEl.value):gwT;
    const gw=choice==='weigh'?gwW:gwT;
    const csize=document.getElementById('csize_'+c.seq)?.value||'40HQ';
    const tareEl2=document.getElementById('tare_'+seq);
  const tare=tareEl2&&tareEl2.value?parseFloat(tareEl2.value):getTareDefault(csize);
    S.vgmData[c.seq]={useWeigh:choice==='weigh',gwFinal:gw,vgm:Math.round(gw+tare)};
    if(choice==='weigh')usedWeigh=true;
  });
  document.getElementById('gwReminder').style.display=usedWeigh?'block':'none';
  closeVgmModal();
  doGenerate();
}

// ── 生成单据 ─────────────────────────────────────────────────
async function doGenerate(){
  const groups={};
  Object.entries(S.containerGroups).forEach(([seq,gi])=>{
    if(!groups[gi])groups[gi]=[];groups[gi].push(parseInt(seq));
  });
  const batchBase=document.getElementById('batchNo').value.trim()||
    `${document.getElementById('buyerSel').value}-${document.getElementById('shipDate').value.replace(/-/g,'').slice(2)}`;
  const productCodesArr=S.productCodes||[];
  const customsSuitsArr=productCodesArr.map(pc=>{
    const el=document.getElementById('suits_'+pc);return el?parseInt(el.value)||0:0;
  });

  // 收集货柜详情
  const containerDetails=S.containers.map(c=>({
    seq:c.seq, container_no:c.container_no,
    seal_no:document.getElementById('seal_'+c.seq)?.value||'',
    container_size:document.getElementById('csize_'+c.seq)?.value||'40HQ',
    so_no:document.getElementById('so_'+c.seq)?.value||'',
    etd:document.getElementById('etd_'+c.seq)?.value||'',
    gw_final:S.vgmData[c.seq]?.gwFinal||S.containerGwTheory[c.seq]||0,
    vgm:S.vgmData[c.seq]?.vgm||0,
  }));

  const batches=Object.entries(groups).map(([gi,seqs])=>({
    batch_no:`${batchBase}-${GRP[parseInt(gi)%GRP.length]}`,
    customer_code:document.getElementById('buyerSel').value,
    shipment_date:document.getElementById('shipDate').value,
    invoice_no:document.getElementById('invoiceNo').value.trim()+(Object.keys(groups).length>1?GRP[parseInt(gi)%GRP.length]:''),
    product_codes:productCodesArr,
    customs_suits:customsSuitsArr,
    container_seqs:seqs,
    seller_code:document.getElementById('sellerSel').value,
    payment_terms:document.getElementById('payTerms').value,
    incoterms:document.getElementById('incoterms').value,
    contract_offset:parseInt(document.getElementById('contractOffset').value)||0,
    container_details:containerDetails.filter(c=>seqs.includes(c.seq)),
  }));

  // UI
  const prog=document.getElementById('progCard');
  const res=document.getElementById('resCard');
  prog.classList.add('show');res.classList.remove('show');
  document.getElementById('termLog').innerHTML='';
  document.getElementById('sdot').className='sdot run';
  document.getElementById('progTitle').textContent='正在生成单据…';
  document.getElementById('progSub').textContent=`${batches.length} 套单据`;
  document.getElementById('genBtn').disabled=true;
  prog.scrollIntoView({behavior:'smooth',block:'nearest'});

  try{
    const r=await fetch('/generate_v4',{method:'POST',
      headers:{'Content-Type':'application/json'},body:JSON.stringify({packing_path:S.packingPath,batches})});
    const d=await r.json();
    if(!d.ok){addLog('e',d.error);document.getElementById('sdot').className='sdot err';document.getElementById('genBtn').disabled=false;return;}
    pollTask(d.task_id);
  }catch(e){addLog('e','请求失败：'+e.message);document.getElementById('sdot').className='sdot err';document.getElementById('genBtn').disabled=false;}
}

async function pollTask(tid){
  const poll=async()=>{
    try{
      const r=await fetch(`/task/${tid}`);const d=await r.json();
      (d.logs||[]).forEach(l=>{
        const c=l.includes('✅')?'o':l.includes('⚠')?'w':l.includes('✗')?'e':
                  (l.includes('📦')||l.includes('📄')||l.includes('🔗'))?'i':'d';
        addLog(c,l);
      });
      if(d.status==='running'){setTimeout(poll,800);return;}
      if(d.status==='done'){
        document.getElementById('sdot').className='sdot ok';
        document.getElementById('progTitle').textContent='✅ 生成完成';
        document.getElementById('progSub').textContent=`${d.outputs.length} 套单据`;
        showResults(d.outputs);
      }else{document.getElementById('sdot').className='sdot err';document.getElementById('progTitle').textContent='❌ 生成失败';}
      document.getElementById('genBtn').disabled=false;
    }catch(e){setTimeout(poll,2000);}
  };poll();
}

function showResults(outputs){
  const rc=document.getElementById('resCard');const rg=document.getElementById('resGrid');
  rc.classList.add('show');
  document.getElementById('resSub').textContent=`${outputs.length} 套 · ${new Date().toLocaleTimeString()}`;
  rg.innerHTML='';
  outputs.forEach(out=>{
    if(out.set1)rg.innerHTML+=fCard('s1',out.set1,out.set1_name,'套一（中国出口）','Inv.·PL·报关单·合同·SI');
    if(out.set2)rg.innerHTML+=fCard('s2',out.set2,out.set2_name,'套二（泰国清关）','Inv.(含Unit Code)·PL(含图片)');
  });
  S.history.unshift({label:outputs.map(o=>o.batch).join('+'),time:new Date().toLocaleTimeString(),outputs});
  renderHist();rc.scrollIntoView({behavior:'smooth',block:'nearest'});
}
function fCard(cls,fid,fname,tag,sheets){
  return `<div class="rf ${cls}">
    <div class="rf-tag">${tag}</div>
    <div class="rf-name">📄 ${fname}</div>
    <div style="font-size:11px;color:var(--gray)">${sheets}</div>
    <button class="btn ${cls==='s1'?'btn-outline':'btn-green'} btn-block"
      onclick="dl('${fid}','${fname}')" style="margin-top:6px">⬇ 下载</button>
  </div>`;
}
async function dl(fid,fname){const a=document.createElement('a');a.href=`/download/${fid}`;a.download=fname;document.body.appendChild(a);a.click();document.body.removeChild(a);}
function renderHist(){
  if(!S.history.length)return;
  document.getElementById('histCard').style.display='block';
  document.getElementById('histList').innerHTML=S.history.slice(0,8).map(h=>`
    <div class="hist-item">
      <div><div style="font-size:13px;font-weight:600;font-family:var(--mono);color:var(--navy)">${h.label}</div>
        <div style="font-size:11px;color:var(--gray)">${h.time}</div></div>
      <div style="display:flex;gap:6px">
        ${h.outputs.flatMap(o=>[
          o.set1?`<button class="btn btn-outline btn-sm" onclick="dl('${o.set1}','${o.set1_name}')">套一</button>`:'',
          o.set2?`<button class="btn btn-green btn-sm" onclick="dl('${o.set2}','${o.set2_name}')">套二</button>`:'',
        ]).join('')}
      </div>
    </div>`).join('');
}
function addLog(cls,text){
  const log=document.getElementById('termLog');
  const t=new Date().toLocaleTimeString('zh-CN',{hour12:false});
  const d=document.createElement('div');d.className='ll';
  d.innerHTML=`<span class="lt">${t}</span><span class="l${cls}">${esc(text)}</span>`;
  log.appendChild(d);log.scrollTop=log.scrollHeight;
}
function esc(t){return String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}

// 拖拽
const pz=document.getElementById('packZone');
pz.addEventListener('dragover',e=>{e.preventDefault();pz.classList.add('drag')});
pz.addEventListener('dragleave',()=>pz.classList.remove('drag'));
pz.addEventListener('drop',e=>{e.preventDefault();pz.classList.remove('drag');
  const f=e.dataTransfer.files[0];if(f){const i=pz.querySelector('input');
  const dt=new DataTransfer();dt.items.add(f);i.files=dt.files;handlePack(i);}});
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
                'seal_no': c.seal_no, 'container_size': c.container_size,
                'vgm_kg': c.vgm_kg, 'so_no': c.so_no,
                'etd': c.etd, 'port_loading': c.port_loading})
    except Exception as e:
        pass
    return jsonify(ok=True, path=path, size=size_str,
                   containers=containers, container_count=len(containers),
                   row_count=row_count, product_codes=product_codes,
                   container_gw=container_gw)

@app.route('/generate_v4', methods=['POST'])
def generate_v4():
    data = request.json
    packing_path = data.get('packing_path')
    batches = data.get('batches', [])
    master_path = get_master_xlsx_path()
    if not master_path: return jsonify(ok=False, error='请先上传主数据管理手册')
    if not packing_path or not os.path.exists(packing_path):
        return jsonify(ok=False, error='请先上传装箱清单')
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
