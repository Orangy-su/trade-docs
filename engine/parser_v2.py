"""
parser_v2.py — 新版装箱清单解析器（Sheet: 最终模板输出）
列映射：
  col0=序号, col1=整机编码, col2=工厂型号, col3=物料编码
  col4=BOM配件名, col6=中文名(ET), col7=英文名(ET)
  col11=外箱尺寸, col12=NW/pcs, col13=每箱数量
  col14=是否报关, col15=装柜位置, col16=装柜箱数
  col17=BOM理论用量, col18=本次计划装柜数量
  第0行 col20..col29=货柜号(柜1..柜10)
"""
import pandas as pd, re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


@dataclass
class PackingRow:
    seq_no: int
    product_code: str
    factory_model: str
    material_code: str
    name_cn: str
    name_en: str
    container_seq: int
    container_label: str
    box_count: float
    qty_per_box: float
    total_qty: float
    box_size_str: str
    box_l_mm: float
    box_w_mm: float
    box_h_mm: float
    gw_per_box: float
    nw_per_box: float
    origin: str = 'China'
    is_customs: bool = True
    remark: str = ''

    @property
    def cbm_per_box(self):
        if self.box_l_mm > 0 and self.box_w_mm > 0 and self.box_h_mm > 0:
            return self.box_l_mm/1000 * self.box_w_mm/1000 * self.box_h_mm/1000
        return 0.0

    @property
    def total_cbm(self): return round(self.cbm_per_box * self.box_count, 4)
    @property
    def total_gw(self): return round(self.gw_per_box * self.box_count, 2)
    @property
    def total_nw(self): return round(self.nw_per_box * self.box_count, 2)


@dataclass
class ContainerInfo:
    seq: int
    label: str
    container_no: str
    seal_no: str = ''
    container_size: str = '40HQ'
    vgm_kg: Optional[float] = None
    so_no: str = ''
    etd: str = ''
    port_loading: str = ''


@dataclass
class PackingHeader:
    customer_code: str = ''
    shipment_date: str = ''
    preparer: str = ''
    checker: str = ''


@dataclass
class PackingListData:
    header: PackingHeader
    containers: Dict[int, ContainerInfo]
    container_by_label: Dict[str, ContainerInfo]
    rows: List[PackingRow]
    product_codes: List[str]

    def rows_for_containers(self, seqs: List[int]) -> List[PackingRow]:
        return [r for r in self.rows if r.container_seq in seqs]


def _safe_str(val, default='') -> str:
    if pd.isna(val): return default
    s = str(val).strip()
    return default if s in ('nan','NaN','') else s

def _safe_float(val, default=0.0) -> float:
    try:
        f = float(val)
        return default if pd.isna(f) else f
    except: return default

def _parse_box_size(s: str) -> Tuple[float,float,float]:
    if not s: return 0,0,0
    parts = re.split(r'[*xX×]', s.strip())
    if len(parts) >= 3:
        try: return float(parts[0]), float(parts[1]), float(parts[2])
        except: pass
    return 0,0,0

def _extract_seq(label: str) -> int:
    m = re.search(r'(\d+)', str(label))
    return int(m.group(1)) if m else 0


def parse_packing_list_v2(xlsx_path: str) -> PackingListData:
    xl = pd.ExcelFile(xlsx_path)
    sheet = '最终模板输出' if '最终模板输出' in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None)

    # 货柜号（第0行 col20..col29）
    row0 = df.iloc[0]
    containers: Dict[int, ContainerInfo] = {}
    container_by_label: Dict[str, ContainerInfo] = {}
    for ci in range(20, 30):
        if ci >= len(row0): break
        v = _safe_str(row0[ci])
        if v and re.match(r'[A-Z]{4}\d+', v):
            seq = ci - 19
            c = ContainerInfo(seq=seq, label=f'柜{seq}', container_no=v)
            containers[seq] = c
            container_by_label[f'柜{seq}'] = c

    # 明细行（从index=3开始）
    rows: List[PackingRow] = []
    seen_products: List[str] = []

    for ri in range(3, len(df)):
        row = df.iloc[ri]
        v0 = _safe_str(row[0])
        if not v0 or '编制' in v0: continue
        try: seq_no = int(float(v0))
        except: continue

        prod   = _safe_str(row[1])
        model  = _safe_str(row[2])
        mat    = _safe_str(row[3])
        name_cn = _safe_str(row[6])
        name_en = _safe_str(row[7]) or _safe_str(row[4])
        box_size = _safe_str(row[11])
        nw_pcs   = _safe_float(row[12])
        qty_box  = _safe_float(row[13])
        is_cust  = _safe_str(row[14]) != '否'
        ctnr_lbl = _safe_str(row[15])
        box_cnt  = _safe_float(row[16])
        plan_qty = _safe_float(row[18])

        # 跳过组件汇总行、陆运行、空行
        if not mat or mat in ('组件',) or '陆运' in ctnr_lbl or not ctnr_lbl:
            continue
        seq = _extract_seq(ctnr_lbl)
        if seq == 0: continue

        l, w, h = _parse_box_size(box_size)
        total_qty = plan_qty if plan_qty > 0 else box_cnt * qty_box

        rows.append(PackingRow(
            seq_no=seq_no, product_code=prod, factory_model=model,
            material_code=mat, name_cn=name_cn, name_en=name_en,
            container_seq=seq, container_label=ctnr_lbl,
            box_count=box_cnt, qty_per_box=qty_box, total_qty=total_qty,
            box_size_str=box_size, box_l_mm=l, box_w_mm=w, box_h_mm=h,
            gw_per_box=0.0, nw_per_box=nw_pcs, is_customs=is_cust,
        ))
        if prod and prod not in seen_products:
            seen_products.append(prod)
        if seq not in containers:
            c = ContainerInfo(seq=seq, label=f'柜{seq}', container_no='')
            containers[seq] = c
            container_by_label[f'柜{seq}'] = c

    return PackingListData(
        header=PackingHeader(),
        containers=containers,
        container_by_label=container_by_label,
        rows=rows,
        product_codes=seen_products,
    )


def parse_packing_list_auto(xlsx_path: str) -> PackingListData:
    """自动识别新旧格式"""
    xl = pd.ExcelFile(xlsx_path)
    if '最终模板输出' in xl.sheet_names:
        return parse_packing_list_v2(xlsx_path)
    # 旧格式兼容
    from .parser import parse_packing_list
    old = parse_packing_list(xlsx_path)
    return PackingListData(
        header=old.header,
        containers={k: ContainerInfo(seq=k, label=f'柜{k}',
            container_no=v.container_no, seal_no=v.seal_no,
            container_size=v.container_size, vgm_kg=v.vgm_kg,
            so_no=v.so_no, etd=v.etd, port_loading=v.port_loading)
            for k, v in old.containers.items()},
        container_by_label={f'柜{k}': ContainerInfo(seq=k, label=f'柜{k}',
            container_no=v.container_no) for k, v in old.containers.items()},
        rows=[PackingRow(
            seq_no=i, product_code=r.product_code, factory_model='',
            material_code=r.material_code, name_cn=r.name_cn, name_en=r.name_en,
            container_seq=r.container_seq, container_label=f'柜{r.container_seq}',
            box_count=r.box_count, qty_per_box=r.qty_per_box, total_qty=r.total_qty,
            box_size_str='', box_l_mm=r.box_l_mm, box_w_mm=r.box_w_mm, box_h_mm=r.box_h_mm,
            gw_per_box=r.gw_per_box, nw_per_box=r.nw_per_box,
        ) for i, r in enumerate(old.rows)],
        product_codes=list({r.product_code for r in old.rows}),
    )
