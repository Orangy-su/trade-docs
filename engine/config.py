"""
config.py — 从主数据管理手册读取所有配置，不硬编码任何业务数字
"""
import pandas as pd
from dataclasses import dataclass, field
from typing import Dict, List, Optional


# ── 数据类 ────────────────────────────────────────────────────────────────────

@dataclass
class FinishedGood:
    """① 成品档案"""
    code: str                      # 成品编码（主键）
    name_cn: str                   # 成品中文名
    name_en: str                   # 成品英文名（发票/报关抬头）
    category: str                  # 产品大类
    model: str                     # 规格型号（报关申报）
    hs_code_cn: str                # HS Code（中国出口8位）
    customs_name_cn: str           # 报关申报名（中文）
    customs_elements: str          # 申报要素（|分隔格式）
    brand_note: str                # 品牌类型标注
    unit_price_customs: float      # 报关整机单价 USD（FOB）
    hs_code_th: str                # HS Code（泰国进口）
    name_th: str                   # 泰国申报名称
    contract_date_offset: int      # 合同日期偏移天数（负=早于发票）
    conversion_factor: Optional[float]  # 整机换算系数（None=手填套数）
    seller_code: str               # 关联卖方代码


@dataclass
class ComponentPrice:
    """② 物料价格表"""
    product_code: str              # 成品编码（主键1）
    material_code: str             # 物料编码（主键2）
    name_en: str                   # 英文品名（发票用）
    name_cn: str                   # 中文品名
    unit_price: float              # 单价 USD
    currency: str                  # 成交币种
    incoterms: str                 # 贸易术语
    nw_per_box: Optional[float]    # NW kg/箱
    gw_per_box: Optional[float]    # GW kg/箱
    cbm_per_box: Optional[float]   # 单箱CBM（自动从装箱清单算）
    hs_code_th: str                # 泰国HS Code（SI用）
    customs_required: bool         # 是否报关
    remark: str


@dataclass
class Party:
    """③ 往来方档案"""
    code: str
    party_type: str                # SELLER / BUYER / FORWARDER
    name_en: str
    name_cn: str
    address_en: str
    consignee_name: str            # FORWARDER专用
    consignee_address: str         # FORWARDER专用
    contact: str
    payment_terms: str             # BUYER专用
    incoterms: str                 # BUYER专用
    currency: str                  # BUYER专用
    tax_id: str                    # SELLER专用
    port_loading: str              # SELLER专用


@dataclass
class ShipmentMain:
    """⑤ 出货批次记录 - 主表（每套单据）"""
    batch_no: str                  # 批次号
    customer_code: str             # 客户代码
    shipment_date: str             # 出货日期
    invoice_no: str                # Invoice No.
    product_codes: List[str]       # 成品编码列表
    customs_suits: List[int]       # 报关整机套数列表（与成品对应）
    container_seq_str: str         # 包含装柜序号（逗号分隔，如"2,3"）
    seller_code: str               # 关联卖方代码
    remark: str

    @property
    def container_seqs(self) -> List[int]:
        return [int(x.strip()) for x in self.container_seq_str.split(',') if x.strip()]


@dataclass
class ShipmentContainer:
    """⑤ 出货批次记录 - 子表（每个货柜）"""
    seq: int                       # 装柜序号（与装箱清单对应）
    batch_no: str                  # 关联批次号
    container_no: str              # 货柜号
    seal_no: str                   # 铅封号
    container_size: str            # 柜型
    vgm_kg: Optional[float]        # VGM
    so_no: str                     # SO号
    etd: str                       # ETD
    port_loading: str              # 装货港
    remark: str


@dataclass
class DocumentRules:
    """⑥ 单据规则配置"""
    contract_date_offset_days: int = 30
    set1_currency: str = "USD"
    set2_currency: str = "USD"
    nw_ratio: float = 0.95
    mixed_set1_invoice: str = "按成品分行列示"
    si_mode: str = "整机归并"
    invoice_no_format: str = "CF-{yy}-{customer}{mmdd}{seq}A"


# ── 主数据加载器 ──────────────────────────────────────────────────────────────

class MasterData:
    """从主数据管理手册 Excel 加载所有配置"""

    def __init__(self, master_path: str):
        self.path = master_path
        self.finished_goods: Dict[str, FinishedGood] = {}
        self.component_prices: Dict[tuple, ComponentPrice] = {}  # (product_code, material_code)
        self.parties: Dict[str, Party] = {}
        self.shipment_mains: Dict[str, ShipmentMain] = {}        # batch_no -> ShipmentMain
        self.shipment_containers: Dict[int, ShipmentContainer] = {}  # seq -> Container
        self.rules: DocumentRules = DocumentRules()
        self._load()

    def _load(self):
        xl = pd.ExcelFile(self.path)
        sheet_names = xl.sheet_names

        for sh in sheet_names:
            if '成品档案' in sh:
                self._load_finished_goods(xl, sh)
            elif '物料价格' in sh:
                self._load_component_prices(xl, sh)
            elif '往来方' in sh:
                self._load_parties(xl, sh)
            elif '出货批次' in sh:
                self._load_shipments(xl, sh)
            elif '规则配置' in sh:
                self._load_rules(xl, sh)

    def _safe_str(self, val, default='') -> str:
        if pd.isna(val): return default
        s = str(val).strip()
        return default if s in ('nan', 'NaN', '（待确认）', '（待填）', '（你的成品编码）') else s

    def _safe_float(self, val) -> Optional[float]:
        try:
            f = float(val)
            return None if pd.isna(f) else f
        except: return None

    def _safe_int(self, val) -> Optional[int]:
        f = self._safe_float(val)
        return int(f) if f is not None else None

    def _load_finished_goods(self, xl, sheet):
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        # 找数据起始行（找到有实际成品编码的行）
        for idx, row in df.iterrows():
            v = self._safe_str(row[0])
            if not v or '成品编码' in v or '★' in v or v.startswith('▌'): continue
            try:
                fg = FinishedGood(
                    code=v,
                    name_cn=self._safe_str(row[1]),
                    name_en=self._safe_str(row[2]),
                    category=self._safe_str(row[3]),
                    model=self._safe_str(row[4]),
                    hs_code_cn=self._safe_str(row[5]),
                    customs_name_cn=self._safe_str(row[6]),
                    customs_elements=self._safe_str(row[7]),
                    brand_note=self._safe_str(row[8]),
                    unit_price_customs=self._safe_float(row[9]) or 0.0,
                    hs_code_th=self._safe_str(row[10]),
                    name_th=self._safe_str(row[11]),
                    contract_date_offset=self._safe_int(row[12]) or -30,
                    conversion_factor=self._safe_float(row[13]),
                    seller_code=self._safe_str(row[14]),
                )
                self.finished_goods[fg.code] = fg
            except Exception as e:
                pass  # 跳过格式不符的行

    def _load_component_prices(self, xl, sheet):
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        for idx, row in df.iterrows():
            prod = self._safe_str(row[0])
            mat  = self._safe_str(row[1])
            if not prod or not mat: continue
            if '成品编码' in prod or '主键' in prod or prod.startswith('▌'): continue
            try:
                nw = self._safe_float(row[7])
                gw = self._safe_float(row[8])
                # 如果NW是公式结果，从GW推算
                if nw is None and gw is not None:
                    nw = round(gw * 0.95, 2)
                cp = ComponentPrice(
                    product_code=prod,
                    material_code=mat,
                    name_en=self._safe_str(row[2]),
                    name_cn=self._safe_str(row[3]),
                    unit_price=self._safe_float(row[4]) or 0.0,
                    currency=self._safe_str(row[5], 'USD'),
                    incoterms=self._safe_str(row[6], 'FOB'),
                    nw_per_box=nw,
                    gw_per_box=gw,
                    cbm_per_box=self._safe_float(row[9]),
                    hs_code_th=self._safe_str(row[10]),
                    customs_required=str(row[11]).strip() != '否',
                    remark=self._safe_str(row[12]),
                )
                self.component_prices[(prod, mat)] = cp
            except: pass

    def _load_parties(self, xl, sheet):
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        for idx, row in df.iterrows():
            code = self._safe_str(row[0])
            ptype = self._safe_str(row[1])
            if not code or not ptype: continue
            if ptype not in ('SELLER', 'BUYER', 'FORWARDER'): continue
            try:
                p = Party(
                    code=code, party_type=ptype,
                    name_en=self._safe_str(row[2]),
                    name_cn=self._safe_str(row[3]),
                    address_en=self._safe_str(row[4]),
                    consignee_name=self._safe_str(row[5]),
                    consignee_address=self._safe_str(row[6]),
                    contact=self._safe_str(row[7]),
                    payment_terms=self._safe_str(row[8]),
                    incoterms=self._safe_str(row[9], 'FOB'),
                    currency=self._safe_str(row[10], 'USD'),
                    tax_id=self._safe_str(row[11]),
                    port_loading=self._safe_str(row[12]),
                )
                self.parties[code] = p
            except: pass

    def _load_shipments(self, xl, sheet):
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        # 分两段：主表（批次号开头的行）和子表（装柜序号数字开头的行）
        in_sub = False
        for idx, row in df.iterrows():
            v0 = self._safe_str(row[0])
            if not v0: continue
            if '子表' in v0 or '货柜明细' in v0: in_sub = True; continue
            if '主表' in v0 or '批次号' in v0 or v0.startswith('▌'): in_sub = False; continue

            if not in_sub:
                # 主表行：v0=批次号
                try:
                    prod_str = self._safe_str(row[4])
                    suits_str = self._safe_str(row[5])
                    seq_str = self._safe_str(row[6])
                    if not seq_str: continue

                    prods = [p.strip() for p in prod_str.split('|') if p.strip()]
                    suits = []
                    for s in suits_str.split('|'):
                        try: suits.append(int(float(s.strip())))
                        except: suits.append(0)

                    sm = ShipmentMain(
                        batch_no=v0,
                        customer_code=self._safe_str(row[1]),
                        shipment_date=self._safe_str(row[2]),
                        invoice_no=self._safe_str(row[3]),
                        product_codes=prods,
                        customs_suits=suits,
                        container_seq_str=seq_str,
                        seller_code=self._safe_str(row[8]),
                        remark=self._safe_str(row[9]),
                    )
                    self.shipment_mains[v0] = sm
                except: pass
            else:
                # 子表行：v0=装柜序号（数字）
                try:
                    seq = self._safe_int(row[0])
                    if seq is None: continue
                    sc = ShipmentContainer(
                        seq=seq,
                        batch_no=self._safe_str(row[1]),
                        container_no=self._safe_str(row[2]),
                        seal_no=self._safe_str(row[3]),
                        container_size=self._safe_str(row[4], '40HQ'),
                        vgm_kg=self._safe_float(row[5]),
                        so_no=self._safe_str(row[6]),
                        etd=self._safe_str(row[7]),
                        port_loading=self._safe_str(row[8]),
                        remark=self._safe_str(row[9]),
                    )
                    self.shipment_containers[seq] = sc
                except: pass

    def _load_rules(self, xl, sheet):
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        rules_map = {}
        for _, row in df.iterrows():
            key = self._safe_str(row[0])
            val = self._safe_str(row[1])
            if key and val and not key.startswith('──'):
                rules_map[key] = val
        try:
            self.rules.contract_date_offset_days = int(rules_map.get('合同日期偏移（天）', 30))
            self.rules.set1_currency = rules_map.get('套一发票币种', 'USD')
            self.rules.set2_currency = rules_map.get('套二发票币种', 'USD')
            nw_r = rules_map.get('NW计算系数', '0.95')
            self.rules.nw_ratio = float(nw_r) if nw_r else 0.95
        except: pass

    def get_seller(self, code: str) -> Optional[Party]:
        return self.parties.get(code)

    def get_buyer(self, code: str) -> Optional[Party]:
        return self.parties.get(code)

    def get_forwarder_for_buyer(self, buyer_code: str) -> Optional[Party]:
        """根据买方找货代（往来方档案里buyer的contact里没有货代代码，先返回第一个FORWARDER）"""
        for p in self.parties.values():
            if p.party_type == 'FORWARDER':
                return p
        return None

    def get_component(self, product_code: str, material_code: str) -> Optional[ComponentPrice]:
        return self.component_prices.get((product_code, material_code))

    def get_finished_good(self, code: str) -> Optional[FinishedGood]:
        return self.finished_goods.get(code)
