"""
parser.py — 解析仓管员装箱清单，与主数据合并，输出每套单据的结构化数据
"""
import pandas as pd
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from .config import MasterData, ComponentPrice


# ── 装箱清单解析结果 ──────────────────────────────────────────────────────────

@dataclass
class PackingRow:
    """装箱清单明细行（解析后）"""
    seq_no: int                    # 行序号
    product_code: str              # 成品编码
    material_code: str             # 物料编码
    name_cn: str                   # 中文品名
    name_en: str                   # 英文品名
    container_seq: int             # 装柜序号
    box_count: float               # 装柜箱数
    qty_per_box: float             # 每箱数量
    total_qty: float               # 总数量
    box_l_mm: float                # 箱长mm
    box_w_mm: float                # 箱宽mm
    box_h_mm: float                # 箱高mm
    gw_per_box: float              # 毛重kg/箱
    nw_per_box: float              # 净重kg/箱
    origin: str                    # 原产地
    is_dangerous: bool             # 危险品
    is_wood_packing: bool          # 木质包装
    remark: str

    @property
    def cbm_per_box(self) -> float:
        """单箱CBM（m³）"""
        if self.box_l_mm > 0 and self.box_w_mm > 0 and self.box_h_mm > 0:
            return self.box_l_mm / 1000 * self.box_w_mm / 1000 * self.box_h_mm / 1000
        return 0.0

    @property
    def total_cbm(self) -> float:
        return round(self.cbm_per_box * self.box_count, 4)

    @property
    def total_gw(self) -> float:
        return round(self.gw_per_box * self.box_count, 2)

    @property
    def total_nw(self) -> float:
        return round(self.nw_per_box * self.box_count, 2)


@dataclass
class PackingHeader:
    """装箱清单表头信息"""
    customer_code: str
    shipment_date: str
    preparer: str
    checker: str
    remark: str


@dataclass
class ContainerInfo:
    """装箱清单货柜信息"""
    seq: int
    container_no: str
    seal_no: str
    container_size: str
    vgm_kg: Optional[float]
    so_no: str
    etd: str
    port_loading: str


@dataclass
class PackingListData:
    """装箱清单完整解析结果"""
    header: PackingHeader
    containers: Dict[int, ContainerInfo]  # seq -> ContainerInfo
    rows: List[PackingRow]

    def rows_for_container(self, seq: int) -> List[PackingRow]:
        return [r for r in self.rows if r.container_seq == seq]

    def rows_for_containers(self, seqs: List[int]) -> List[PackingRow]:
        return [r for r in self.rows if r.container_seq in seqs]


# ── 单据所需的聚合数据 ────────────────────────────────────────────────────────

@dataclass
class ComponentLineItem:
    """套二发票/PL的一行（部件明细）"""
    product_code: str
    material_code: str
    name_en: str
    name_cn: str
    container_seq: int
    box_count: float
    total_qty: float
    total_cbm: float
    total_gw: float
    total_nw: float
    unit_price: float
    total_amount: float
    hs_code_th: str
    origin: str
    customs_required: bool
    remark: str
    qty_per_box: float = 0.0


@dataclass
class FinishedGoodLineItem:
    """套一发票/PL的一行（整机）"""
    product_code: str
    name_en: str
    name_cn: str
    customs_suits: int
    unit_price: float
    total_amount: float
    hs_code_cn: str
    customs_elements: str
    brand_note: str
    customs_name_cn: str


@dataclass
class DocumentBundle:
    """一套完整单据所需的所有数据"""
    # 元数据
    batch_no: str
    invoice_no: str
    invoice_date: str
    contract_date: str
    customer_code: str
    seller_code: str

    # 套一数据
    set1_lines: List[FinishedGoodLineItem]
    set1_total_amount: float

    # 套二数据
    set2_lines: List[ComponentLineItem]
    set2_total_amount: float

    # 货柜汇总
    containers: List[ContainerInfo]          # 本套单据包含的货柜
    total_pkgs: int
    total_nw: float
    total_gw: float
    total_cbm: float

    # 往来方
    seller_name_en: str
    seller_name_cn: str
    seller_address_en: str
    seller_address_cn: str
    seller_tax_id: str
    seller_port_loading: str
    buyer_name_en: str
    buyer_address_en: str
    buyer_payment_terms: str
    buyer_incoterms: str
    buyer_currency: str
    consignee_name: str
    consignee_address: str
    notify_name: str
    notify_address: str

    # SI
    so_nos: List[str]
    etd: str
    port_discharge: str


# ── 解析器 ────────────────────────────────────────────────────────────────────

def _safe_str(val, default='') -> str:
    if pd.isna(val): return default
    s = str(val).strip()
    return default if s in ('nan', 'NaN', '') else s

def _safe_float(val, default=0.0) -> float:
    try:
        f = float(val)
        return default if pd.isna(f) else f
    except: return default

def _safe_int(val, default=0) -> int:
    try: return int(float(val))
    except: return default


def parse_packing_list(xlsx_path: str) -> PackingListData:
    """解析标准化装箱清单（仓管员版）"""
    df = pd.read_excel(xlsx_path, sheet_name='装箱清单', header=None)

    # ── 表头区（行3-4）────────────────────────────────────────────
    header = PackingHeader(
        customer_code=_safe_str(df.iloc[2, 3]),
        shipment_date=_safe_str(df.iloc[2, 9]),
        preparer=_safe_str(df.iloc[2, 14]),
        checker=_safe_str(df.iloc[3, 9]),
        remark=_safe_str(df.iloc[3, 14]),
    )

    # ── 货柜信息区（行8-15，第0-indexed行7-14）──────────────────────
    containers: Dict[int, ContainerInfo] = {}
    for row_idx in range(7, 15):
        if row_idx >= len(df): break
        row = df.iloc[row_idx]
        seq = _safe_int(row[0])
        container_no = _safe_str(row[1])
        if not container_no: continue
        containers[seq] = ContainerInfo(
            seq=seq,
            container_no=container_no,
            seal_no=_safe_str(row[4]),
            container_size=_safe_str(row[6], '40HQ'),
            vgm_kg=float(row[9]) if _safe_str(row[9]) else None,
            so_no=_safe_str(row[10]),
            etd=_safe_str(row[13]),
            port_loading=_safe_str(row[15]),
        )

    # ── 明细区（从行20开始，0-indexed=19）────────────────────────────
    rows: List[PackingRow] = []
    for row_idx in range(19, len(df)):
        row = df.iloc[row_idx]
        prod = _safe_str(row[1])
        mat  = _safe_str(row[2])
        if not prod or not mat: continue
        # 跳过合计行
        if _safe_str(row[0]) in ('合计', 'TOTAL', '合　计'): continue

        box_count   = _safe_float(row[6])
        qty_per_box = _safe_float(row[7])
        if box_count == 0 and qty_per_box == 0: continue

        # 自动计算总数量（优先用表格的值，否则用公式结果）
        total_qty = _safe_float(row[8])
        if total_qty == 0 and box_count > 0 and qty_per_box > 0:
            total_qty = box_count * qty_per_box

        gw = _safe_float(row[12])
        nw = _safe_float(row[13])
        if nw == 0 and gw > 0:
            nw = round(gw * 0.95, 2)

        pr = PackingRow(
            seq_no=row_idx - 18,
            product_code=prod,
            material_code=mat,
            name_cn=_safe_str(row[3]),
            name_en=_safe_str(row[4]),
            container_seq=_safe_int(row[5]),
            box_count=box_count,
            qty_per_box=qty_per_box,
            total_qty=total_qty,
            box_l_mm=_safe_float(row[9]),
            box_w_mm=_safe_float(row[10]),
            box_h_mm=_safe_float(row[11]),
            gw_per_box=gw,
            nw_per_box=nw,
            origin=_safe_str(row[14], 'China'),
            is_dangerous=_safe_str(row[15]) == '是',
            is_wood_packing=_safe_str(row[16]) == '是',
            remark=_safe_str(row[17]),
        )
        rows.append(pr)

    return PackingListData(header=header, containers=containers, rows=rows)


def build_document_bundle(
    master: MasterData,
    batch_no: str,
    packing: PackingListData,
) -> DocumentBundle:
    """
    把主数据 + 出货批次记录 + 装箱清单 合并，生成一套单据所需的完整数据
    """
    sm = master.shipment_mains.get(batch_no)
    if not sm:
        raise ValueError(f"批次号 {batch_no} 在主数据出货批次记录中未找到")

    # ── 往来方 ────────────────────────────────────────────────────
    seller = master.get_seller(sm.seller_code) or master.get_seller('QF-CN')
    buyer  = master.get_buyer(sm.customer_code)
    forwarder = master.get_forwarder_for_buyer(sm.customer_code)

    # ── 货柜列表（主数据子表优先，装箱清单补充）────────────────────
    container_seqs = sm.container_seqs
    containers_in_bundle: List[ContainerInfo] = []
    for seq in container_seqs:
        # 主数据子表有更准确的柜号/铅封（单证员填的）
        sc = master.shipment_containers.get(seq)
        if sc:
            containers_in_bundle.append(ContainerInfo(
                seq=seq,
                container_no=sc.container_no,
                seal_no=sc.seal_no,
                container_size=sc.container_size,
                vgm_kg=sc.vgm_kg,
                so_no=sc.so_no,
                etd=sc.etd,
                port_loading=sc.port_loading or (seller.port_loading if seller else 'NANSHA'),
            ))
        elif seq in packing.containers:
            containers_in_bundle.append(packing.containers[seq])

    # ── 套一数据（整机归并，支持同一成品编码多品名报关）──────────
    set1_lines: List[FinishedGoodLineItem] = []
    set1_total = 0.0
    for i, prod_code in enumerate(sm.product_codes):
        suits = sm.customs_suits[i] if i < len(sm.customs_suits) else 0
        # 取该成品编码的所有行（多品名报关时有多行）
        fg_list = getattr(master, '_fg_list', {}).get(prod_code)
        if not fg_list:
            fg = master.get_finished_good(prod_code)
            fg_list = [fg] if fg else []
        for fg in fg_list:
            if not fg or fg.unit_price_customs == 0: continue
            amount = round(suits * fg.unit_price_customs, 2)
            set1_total += amount
            set1_lines.append(FinishedGoodLineItem(
                product_code=prod_code,
                name_en=fg.name_en,
                name_cn=fg.name_cn,
                customs_suits=suits,
                unit_price=fg.unit_price_customs,
                total_amount=amount,
                hs_code_cn=fg.hs_code_cn,
                customs_elements=fg.customs_elements,
                brand_note=fg.brand_note,
                customs_name_cn=fg.customs_name_cn,
            ))

    # ── 套二数据（部件明细，从装箱清单汇总）──────────────────────
    packing_rows = packing.rows_for_containers(container_seqs)
    set2_lines: List[ComponentLineItem] = []
    set2_total = 0.0
    total_pkgs = 0
    total_nw = 0.0
    total_gw = 0.0
    total_cbm = 0.0

    for pr in packing_rows:
        # 用成品编码+物料编码查单价
        cp = master.get_component(pr.product_code, pr.material_code)
        unit_price = cp.unit_price if cp else 0.0
        hs_th = cp.hs_code_th if cp else ''
        customs_req = cp.customs_required if cp else True
        # 如果装箱清单有GW，优先用装箱清单的
        gw_box = pr.gw_per_box if pr.gw_per_box > 0 else (cp.gw_per_box or 0)
        nw_box = pr.nw_per_box if pr.nw_per_box > 0 else (cp.nw_per_box or round(gw_box * 0.95, 2))
        # 如果GW=0但NW>0，从NW反推（NW÷0.95）
        if gw_box == 0 and nw_box > 0:
            gw_box = round(nw_box / 0.95, 2)

        t_gw  = round(gw_box * pr.box_count, 2)
        t_nw  = round(nw_box * pr.box_count, 2)
        t_cbm = pr.total_cbm
        t_amt = round(pr.total_qty * unit_price, 2)

        total_pkgs += int(pr.box_count)
        total_gw  += t_gw
        total_nw  += t_nw
        total_cbm += t_cbm
        set2_total += t_amt

        # 英文品名优先用主数据，其次装箱清单
        name_en = (cp.name_en if cp and cp.name_en else '') or pr.name_en
        name_cn = (cp.name_cn if cp and cp.name_cn else '') or pr.name_cn

        set2_lines.append(ComponentLineItem(
            product_code=pr.product_code,
            material_code=pr.material_code,
            name_en=name_en,
            name_cn=name_cn,
            container_seq=pr.container_seq,
            box_count=pr.box_count,
            total_qty=pr.total_qty,
            total_cbm=round(t_cbm, 3),
            total_gw=t_gw,
            total_nw=t_nw,
            unit_price=unit_price,
            total_amount=t_amt,
            hs_code_th=hs_th,
            origin=pr.origin,
            customs_required=customs_req,
            remark=pr.remark,
            qty_per_box=pr.qty_per_box,
        ))

    # ── 合同日期计算 ──────────────────────────────────────────────
    from datetime import datetime, timedelta
    try:
        inv_date = datetime.strptime(sm.shipment_date[:10], '%Y-%m-%d')
        fg_first = master.get_finished_good(sm.product_codes[0]) if sm.product_codes else None
        offset = fg_first.contract_date_offset if fg_first else -30
        contract_date = (inv_date + timedelta(days=offset)).strftime('%Y-%m-%d')
    except:
        contract_date = sm.shipment_date

    so_nos = [c.so_no for c in containers_in_bundle if c.so_no]
    port_discharge = buyer.address_en[:20] if buyer else 'Bangkok'
    etd_val = containers_in_bundle[0].etd if containers_in_bundle else sm.shipment_date

    return DocumentBundle(
        batch_no=batch_no,
        invoice_no=sm.invoice_no,
        invoice_date=sm.shipment_date,
        contract_date=contract_date,
        customer_code=sm.customer_code,
        seller_code=sm.seller_code,

        set1_lines=set1_lines,
        set1_total_amount=round(set1_total, 2),

        set2_lines=set2_lines,
        set2_total_amount=round(set2_total, 2),

        containers=containers_in_bundle,
        total_pkgs=total_pkgs,
        total_nw=round(total_nw, 2),
        total_gw=round(total_gw, 2),
        total_cbm=round(total_cbm, 3),

        seller_name_en=seller.name_en if seller else '',
        seller_name_cn=seller.name_cn if seller else '',
        seller_address_en=seller.address_en if seller else '',
        seller_address_cn=seller.address_en if seller else '',
        seller_tax_id=seller.tax_id if seller else '',
        seller_port_loading=seller.port_loading if seller else 'NANSHA',

        buyer_name_en=buyer.name_en if buyer else '',
        buyer_address_en=buyer.address_en if buyer else '',
        buyer_payment_terms=buyer.payment_terms if buyer else 'T/T 45 days after B/L Date',
        buyer_incoterms=buyer.incoterms if buyer else 'FOB',
        buyer_currency=buyer.currency if buyer else 'USD',

        consignee_name=forwarder.name_en if forwarder else '',
        consignee_address=forwarder.address_en if forwarder else '',
        notify_name=forwarder.notify_name if forwarder else '',
        notify_address=forwarder.notify_address if forwarder else '',

        so_nos=so_nos,
        etd=etd_val,
        port_discharge=(forwarder.port_discharge if forwarder and hasattr(forwarder,'port_discharge') and forwarder.port_discharge
                       else 'Bangkok'),
    )
