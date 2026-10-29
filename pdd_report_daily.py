# pdd_report_daily.py - 拼多多出单日报统计
# 功能：读取拼多多订单Excel，按状态过滤，提取字段并生成规范数据表

import os
import sys
import csv
from typing import Optional
from tkinter import Tk, filedialog
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from datetime import datetime


# ============================
# 配置
# ============================
def get_desktop_path() -> str:
    """
    获取当前系统桌面路径。
    Windows 下通过注册表读取真实桌面路径，
    兼容桌面被迁移到非默认位置（如 D:\\桌面）的情况。
    """
    if sys.platform == "win32":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            )
            desktop, _ = winreg.QueryValueEx(key, "Desktop")
            winreg.CloseKey(key)
            return desktop
        except Exception:
            pass
    # macOS / Linux 或注册表读取失败时的兜底
    return os.path.join(os.path.expanduser("~"), "Desktop")

DESKTOP = get_desktop_path()

HEADERS = [
    "序号", "日期", "地址", "规格", "颜色", "尺寸", "总米数",
    "封口", "封口个数", "滑轮", "滑轮个数", "膨胀螺丝/套",
    "免钉胶/套", "安装码", "安装码数量", "连接器", "连接器数量",
    "配件", "配件数量", "备注", "订单号", "快递单号/订单编号"
]

OUTPUT_KEYS = [
    "序号", "日期", "地址", "规格", "颜色", "尺寸", "总米数",
    "封口", "封口个数", "滑轮", "滑轮个数", "膨胀螺丝/套",
    "免钉胶/套", "安装码", "安装码数量", "连接器", "连接器数量",
    "配件", "配件数量", "备注", "订单号", "快递单号/订单编号"
]

# 有效订单状态（只保留这些状态的行）
VALID_STATUSES = {"已发货，待收货", "已发货，退款成功", "已收货", "已收货，退款成功"}

# 无效备注关键词：备注里出现这些词的订单并未实际发单，
# 整条视为无效数据，既不进明细表，也不进配件统计表。
# 注意：不要加入"不要"（备注里常见"不要发错/不要切错"，是正常订单的叮嘱）。
INVALID_REMARK_KEYWORDS = ("废除", "作废", "废单", "暂不发货", "取消")

# 快递商家识别映射
EXPRESS_MAP = {
    "YT": "圆通",
    "JT": "极兔",
    "DPK": "德邦",
}

# 中文数字映射
CHINESE_NUMBERS = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
    "十一": 11, "十二": 12, "十三": 13, "十四": 14, "十五": 15,
    "十六": 16, "十七": 17, "十八": 18, "十九": 19, "二十": 20,
    "三十": 30, "四十": 40, "五十": 50, "六十": 60, "七十": 70,
    "八十": 80, "九十": 90, "百": 100
}


# 配件统计表的品类顺序（与明细表字段一一对应）
STAT_CATEGORIES = ["封口", "滑轮", "膨胀螺丝/套", "免钉胶/套", "安装码", "连接器", "配件"]


# ============================
# 工具函数
# ============================

def parse_quantity(qty_str: str) -> str:
    """解析数量，支持阿拉伯数字和中文数字"""
    if not qty_str:
        return "/"
    qty_str = qty_str.strip()
    # 尝试阿拉伯数字
    if qty_str.isdigit():
        return qty_str
    # 尝试中文数字
    if qty_str in CHINESE_NUMBERS:
        return str(CHINESE_NUMBERS[qty_str])
    # 无法识别，返回原值或 "/"
    return qty_str if qty_str else "/"


def detect_express_company(tracking_no: str) -> str:
    """根据快递单号前缀识别快递商家"""
    if not tracking_no:
        return ""
    tracking_no = tracking_no.strip().upper()
    for prefix, name in EXPRESS_MAP.items():
        if tracking_no.startswith(prefix):
            return name
    return ""


def format_date(date_str: str) -> str:
    """将发货时间处理为 '5月6号' 格式"""
    if not date_str:
        return ""
    try:
        # 优先匹配常见格式
        for fmt in ("%Y/%m/%d %H:%M:%S", "%Y年%m月%d日 %H:%M:%S",
                    "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M",
                    "%Y-%m-%d", "%Y年%m月%d日"):
            try:
                dt = datetime.strptime(str(date_str).strip(), fmt)
                return f"{dt.month}月{dt.day}号"
            except ValueError:
                continue
        # 无法解析时返回原值
        return str(date_str)
    except Exception:
        return str(date_str)


ACCESSORY_PATTERNS = {
    "款式/颜色": [
        "F-4", "黑色", "白色", "香槟色",
        "Z21", "Z22", "002", "004", "005", "006", "219", "骏派TS420"
    ],
    "滑轮": [
        "7字纳米大滑块", "7字纳米小滑块", "6号钢片两轮", "6号钢片四轮",
        "8号钢片两轮", "8号钢片四轮", "10号纳米两轮", "10号纳米两轮白色",
        "10钢片两轮白色", "15号轴承轮", "轴承轮", "合金轮"
    ],
    "封口": [
        "F-4", "Z21", "Z22", "002", "006", "219", "大方轨封口", "826封口"
    ],
    "安装码": [
        "28单侧码", "28单顶码", "20单顶码", "20单侧码",
        "20双顶码", "13单侧码", "13单顶码"
    ],
    "配件": [
        "F-4转弯器"
    ]
}

def load_accessory_patterns() -> dict:
    """返回硬编码固定的仓库配件关键词映射。"""
    return ACCESSORY_PATTERNS


# ============================
# 核心处理函数
# ============================

def bei_zhu_chai_jie(remark: str) -> dict:
    """
    商家备注拆解函数

    原内容大致格式为：
    "219【白】 3米*2根，1.3米*1根，1.6米*1根 1.8米*1根
     219封口 10个 13 单侧码20个 10 钢片两轮88个 膨胀螺丝 45套"

    拆解规则（对照案例表）：
      规格        → 开头型号前缀，如 "219"（不含【】）
      颜色        → 【颜色】标签，如 "【白】"（无颜色时为 /）
      尺寸        → 所有 "X米*Y根" 片段，用换行符拼接
      总米数      → 自动累加 X*Y
      封口        → "X封口 Y个" → 规格="X封口"，个数=Y
      安装码/数量 → "X 单侧码Y个" → 安装码="X 单侧码"，数量=Y
      滑轮/个数   → "X 钢片两轮Y个" → 滑轮="X 钢片两轮"，个数=Y
      膨胀螺丝/套 → "膨胀螺丝 Y套" → 数量=Y
      连接器/数量 → "连接器 Y个" → 连接器="连接器"，数量=Y
      配件/数量   → 未出现，预留
      无内容字段  → 填充 "/"

    返回字典，对应新表各列
    """
    if not remark or not remark.strip():
        return {
            "规格": "/", "颜色": "/", "尺寸": "/", "总米数": "/",
            "封口": "/", "个数": "/", "滑轮": "/", "个数": "/",
            "膨胀螺丝/套": "/", "安装码": "/", "数量": "/",
            "连接器": "/", "数量": "/", "配件": "/", "数量": "/",
        }

    import re
    # 关键：把换行符等空白统一成空格，避免 \S 跨行匹配导致数值错位
    text = re.sub(r"\s+", " ", remark.strip())

    # --- 规格 & 颜色：拆分为两个独立字段 ---
    # 支持的型号前缀：219、Z21、Z22、002、004、005、006、大方轨、826 等
    color_match = re.search(r"【(.+?)】", text)
    color_value = color_match.group(1) if color_match else ""
    if color_match:
        prefix = text[:color_match.start()].strip()
        # 进一步检查前缀是否为空，如果是则尝试匹配常见型号
        if not prefix:
            for model in ["大方轨1001", "大方轨", "826", "219", "Z21", "Z22", "002", "004", "005", "006", "F-4", "骏派TS420"]:
                if text.startswith(model + "【"):
                    prefix = model
                    break
        spec = prefix  # 规格 = 前缀型号（不含颜色标签）
        color = f"【{color_value}】"  # 颜色 = 完整【颜色】
    else:
        # 没有【】时，尝试匹配常见型号开头作为规格
        spec = "/"
        color = "/"
        for model in ["大方轨1001", "大方轨", "826", "219", "Z21", "Z22", "002", "004", "005", "006", "F-4", "骏派TS420"]:
            if text.startswith(model):
                spec = model
                break

    # --- 尺寸 & 总米数：支持 "X米*Y根" 和 "(a+b)米*Y根" 两种格式 ---
    # 支持半角() 和全角（）括号；分组1：括号表达式（如 "1.585+1.2"）；分组2：纯数字（如 "1.2"）；分组3：根数
    size_pattern = re.compile(r'(?:[()（）]([\d.+]+)[()（）]|([\d.]+))\s*米\s*\*\s*(\d+)\s*根')
    size_matches = size_pattern.findall(text)
    if size_matches:
        sizes = []
        total_m = 0.0
        for m_expr, m_simple, n_str in size_matches:
            n = int(n_str)
            if m_expr:   # (a+b+...)米 格式
                meters = sum(float(x) for x in m_expr.split('+'))
                # 括号内各项也统一保留3位小数展示
                parts = [f"{float(x):.3f}" for x in m_expr.split('+')]
                sizes.append(f"({'+'.join(parts)})米*{n}根")
            else:         # 普通 X米 格式
                meters = float(m_simple)
                sizes.append(f"{meters:.3f}米*{n}根")
            total_m += meters * n
        size_str = "，".join(sizes)
        total_m_rounded = round(total_m, 3)
        # 统一强制保留3位小数
        total_m_str = f"{total_m_rounded:.3f}"
    else:
        size_str, total_m_str = "/", "/"

    def _find_named_count(names, text):
        for name in sorted(set(names), key=len, reverse=True):
            if not name:
                continue
            if name in text:
                # 支持阿拉伯数字和中文数字
                m = re.search(re.escape(name) + r"\s*([一二三四五六七八九十\d]+)\s*个", text)
                if m:
                    qty = parse_quantity(m.group(1))
                    return name, qty
        return None, "/"

    patterns = load_accessory_patterns()

    # --- 封口 ---
    seal_spec, seal_qty = _find_named_count(patterns.get("封口", []), text)
    if seal_spec is None:
        # 支持带颜色标签的格式：219封口【黑】10个
        seal_match = re.search(r"([^\s]+封口)(?:【([^】]*)】)?\s*([一二三四五六七八九十\d]+)\s*个", text)
        if seal_match:
            seal_name = seal_match.group(1)
            seal_color = seal_match.group(2) or ""
            seal_spec = f"{seal_name}{('【' + seal_color + '】') if seal_color else ''}"
            seal_qty = parse_quantity(seal_match.group(3))
        else:
            seal_spec, seal_qty = "/", "/"

    # --- 安装码（支持多款，格内换行） ---
    # 统一用一条 findall 正则抓取全部安装码，顺序与备注一致
    # group(1): 前缀数字/字母, group(2): 码类型关键字, group(3): 颜色(可空), group(4): 数量
    _dan_pattern = re.compile(
        r"([^\s，,]+?)\s*(单侧码|双侧码|单顶码|双顶码)(?:【([^】]*)】)?\s*(\d+)\s*个"
    )
    dan_all = _dan_pattern.findall(text)
    if dan_all:
        dan_specs, dan_qtys = [], []
        for prefix, keyword, dan_color, qty in dan_all:
            color_tag = f"【{dan_color}】" if dan_color else ""
            dan_specs.append(f"{prefix.strip()} {keyword}{color_tag}")
            dan_qtys.append(parse_quantity(qty))
        dan_spec = "\n".join(dan_specs)
        dan_qty  = "\n".join(dan_qtys)
    else:
        dan_spec, dan_qty = "/", "/"

    # --- 滑轮（支持多款，格内换行） ---
    # 用一条 findall 正则抓取全部滑轮条目
    # group(1): 型号名称(含可选前缀数字), group(2): 颜色(可空), group(3): 数量
    # 前缀数字可选（如"轴承轮" / "10 钢片两轮" 均可匹配）
    _wheel_pattern = re.compile(
        r"(\d*\s*[\w\u4e00-\u9fff]*?"
        r"(?:纳米|钢片|轴承轮|合金轮|滑块|走珠)"
        r"[\w\u4e00-\u9fff]*?)"
        r"(?:【([^】]*)】)?"
        r"\s*(\d+)\s*个"
    )
    wheel_all = _wheel_pattern.findall(text)
    if wheel_all:
        lun_specs, lun_qtys = [], []
        for name, wheel_color, qty in wheel_all:
            color_tag = f"【{wheel_color}】" if wheel_color else ""
            lun_specs.append(f"{name.strip()}{color_tag}")
            lun_qtys.append(parse_quantity(qty))
        lun_spec = "\n".join(lun_specs)
        lun_qty  = "\n".join(lun_qtys)
    else:
        # 兜底：从 patterns 字典里的自定义名称里找
        wheel_names = patterns.get("滑轮", []) + ["轴承轮", "合金轮", "8号纳米"]
        lun_spec, lun_qty = _find_named_count(wheel_names, text)
        if lun_spec is None:
            lun_spec, lun_qty = "/", "/"

    # --- 膨胀螺丝 ---
    peng_match = None
    for pattern in [
        r"膨胀螺丝\s*([一二三四五六七八九十\d]+)\s*套",
        r"([一二三四五六七八九十\d]+)\s*膨胀螺丝\s*套",
        r"膨胀螺丝\s*([一二三四五六七八九十\d]+)",
        r"([一二三四五六七八九十\d]+)\s*膨胀螺丝",
    ]:
        peng_match = re.search(pattern, text)
        if peng_match:
            break
    peng_qty = parse_quantity(peng_match.group(1)) if peng_match else "/"

    # --- 免钉胶 ---
    mdj_match = re.search(r"免钉胶\s*([一二三四五六七八九十\d]+)\s*套", text)
    mdj_qty = parse_quantity(mdj_match.group(1)) if mdj_match else "/"

    # --- 连接器 ---
    # 支持多种型号前缀：F630、F525、F轨、1001、219、WR303、826 等
    lian_patterns = [
        r"(F630连接器|F525连接器|F轨连接器|1001连接器|219连接器|WR303连接器|826连接器|大方轨1001连接器|大方轨连接器|连接器)(?:【([^】]*)】)?\s*([一二三四五六七八九十\d]+)\s*个"
    ]
    lian_match = None
    for pattern in lian_patterns:
        lian_match = re.search(pattern, text)
        if lian_match:
            break
    if lian_match:
        lian_name = lian_match.group(1)
        lian_color = lian_match.group(2) or ""
        lian_spec = f"{lian_name}{('【' + lian_color + '】') if lian_color else ''}"
        lian_qty = parse_quantity(lian_match.group(3))
    else:
        lian_spec, lian_qty = "/", "/"

    # --- 配件 ---
    pei_spec, pei_qty = _find_named_count(patterns.get("配件", []), text)
    if pei_spec is None:
        pei_spec, pei_qty = "/", "/"

    return {
        "规格":        spec,
        "颜色":        color,
        "尺寸":       size_str,
        "总米数":     total_m_str,
        "封口":       seal_spec,
        "封口个数":   seal_qty,
        "滑轮":       lun_spec,
        "滑轮个数":   lun_qty,
        "膨胀螺丝/套": peng_qty,
        "免钉胶/套":   mdj_qty,
        "安装码":     dan_spec,
        "安装码数量": dan_qty,
        "连接器":     lian_spec,
        "连接器数量": lian_qty,
        "配件":       pei_spec,
        "配件数量":   pei_qty,
    }


def process_row(row_data: dict) -> dict:
    """
    处理单行数据，返回符合新表格式的字典

    row_data 包含原始行的各字段字典
    """
    # 提取所需字段
    order_id = str(row_data.get("订单号", "")).strip()
    order_status = str(row_data.get("订单状态", "")).strip()
    total_price = row_data.get("商品总价", "")
    remark = str(row_data.get("商家备注", "")).strip()
    tracking_no = str(row_data.get("快递单号", "")).strip()
    ship_time = str(row_data.get("发货时间", "")).strip()

    # 1. 快递单号/订单编号 = 快递公司名称 + 快递单号 + "订单号：" + 订单号
    express_company = detect_express_company(tracking_no)
    new_tracking = f"{express_company}{tracking_no}订单号：{order_id}" if tracking_no else f"订单号：{order_id}"

    # 2. 日期
    new_date = format_date(ship_time)

    # 3. 商家备注拆解
    parsed = bei_zhu_chai_jie(remark)

    def _f(v):
        """空值统一填 /"""
        return v if v not in ("", None) else "/"

    # 4. 组合新行数据（所有空值填 /）
    new_row = {
        "序号":              "",   # 序号由 main() 调用方填入
        "日期":              _f(new_date),
        "地址":              "/",
        "规格":              _f(parsed.get("规格", "")),
        "颜色":              _f(parsed.get("颜色", "")),
        "尺寸":              _f(parsed.get("尺寸", "")),
        "总米数":            _f(parsed.get("总米数", "")),
        "封口":              _f(parsed.get("封口", "")),
        "封口个数":          _f(parsed.get("封口个数", "")),
        "滑轮":              _f(parsed.get("滑轮", "")),
        "滑轮个数":          _f(parsed.get("滑轮个数", "")),
        "膨胀螺丝/套":       _f(parsed.get("膨胀螺丝/套", "")),
        "免钉胶/套":         _f(parsed.get("免钉胶/套", "")),
        "安装码":            _f(parsed.get("安装码", "")),
        "安装码数量":        _f(parsed.get("安装码数量", "")),
        "连接器":            _f(parsed.get("连接器", "")),
        "连接器数量":        _f(parsed.get("连接器数量", "")),
        "配件":              _f(parsed.get("配件", "")),
        "配件数量":          _f(parsed.get("配件数量", "")),
        "备注":              remark if remark else "/",
        "订单号":            _f(order_id),
        "快递单号/订单编号":  _f(new_tracking),
    }
    return new_row


def create_output_file() -> Workbook:
    """创建输出文件，写入表头"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(HEADERS)
    return wb


# ============================
# 配件数量统计（新增功能，不影响原明细输出）
# ============================

def _to_int(qty) -> Optional[int]:
    """把数量安全转成整数，转不了（如 '/'）返回 None"""
    try:
        return int(str(qty).strip())
    except (TypeError, ValueError):
        return None


def collect_stats(stats: dict, processed: dict):
    """
    把一行明细中的配件数量累加进统计字典。

    stats 结构：{品类: {细分种类: 使用总数}}
    数据直接取自明细表的解析结果，保证统计与明细永远一致。
    """
    def add(cat: str, name, qty):
        n = _to_int(qty)
        if n is None or n <= 0:
            return
        if not name:
            return
        # 规范空白字符（多个空格/换行合并为一个空格）
        name = " ".join(str(name).split())
        if not name or name == "/":
            return
        bucket = stats.setdefault(cat, {})
        bucket[name] = bucket.get(name, 0) + n

    # 封口（单条目）
    add("封口", processed.get("封口"), processed.get("封口个数"))

    # 滑轮（多条目，格内换行，规格与数量逐行对应）
    lun_specs = str(processed.get("滑轮", "/")).split("\n")
    lun_qtys = str(processed.get("滑轮个数", "/")).split("\n")
    if len(lun_qtys) < len(lun_specs):
        lun_qtys += ["/"] * (len(lun_specs) - len(lun_qtys))
    for s, q in zip(lun_specs, lun_qtys):
        add("滑轮", s, q)

    # 膨胀螺丝 / 免钉胶（无细分型号，只有总量，细分种类填品类名）
    add("膨胀螺丝/套", "膨胀螺丝", processed.get("膨胀螺丝/套"))
    add("免钉胶/套", "免钉胶", processed.get("免钉胶/套"))

    # 安装码（多条目，格内换行）
    dan_specs = str(processed.get("安装码", "/")).split("\n")
    dan_qtys = str(processed.get("安装码数量", "/")).split("\n")
    if len(dan_qtys) < len(dan_specs):
        dan_qtys += ["/"] * (len(dan_specs) - len(dan_qtys))
    for s, q in zip(dan_specs, dan_qtys):
        add("安装码", s, q)

    # 连接器 / 配件（单条目）
    add("连接器", processed.get("连接器"), processed.get("连接器数量"))
    add("配件", processed.get("配件"), processed.get("配件数量"))


def create_stats_workbook(stats: dict) -> Workbook:
    """
    根据统计字典生成配件统计表工作簿。

    版式（对照案例图）：
      第1行：A列留空 | B1 细分种类 | C1 使用总数
      每个品类一块：A列品类名纵向合并居中，B列细分种类，C列使用总数
      品类块之间空一行；同一品类内按使用总数从高到低排序
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "配件统计"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")

    # 表头
    ws.cell(row=1, column=1).border = border
    for col, title in ((2, "细分种类"), (3, "使用总数")):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = Font(bold=True)
        cell.alignment = center
        cell.border = border

    row = 2
    for cat in STAT_CATEGORIES:
        items = stats.get(cat)
        if not items:
            continue
        # 同品类内按使用总数从高到低
        sorted_items = sorted(items.items(), key=lambda kv: kv[1], reverse=True)
        start = row
        for name, total in sorted_items:
            b = ws.cell(row=row, column=2, value=name)
            b.border = border
            c = ws.cell(row=row, column=3, value=total)
            c.border = border
            c.alignment = right
            row += 1
        end = row - 1
        # 先给品类列每个格子加边框，再纵向合并品类名
        for r in range(start, end + 1):
            ws.cell(row=r, column=1).border = border
        ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
        cat_cell = ws.cell(row=start, column=1, value=cat)
        cat_cell.alignment = center
        row += 1  # 品类块之间空一行

    # 列宽
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 12
    return wb


def scan_lzc_files() -> str:
    """让用户选择输入文件"""
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 设置文件对话框
    file_path = filedialog.askopenfilename(
        title="选择拼多多订单数据文件",
        initialdir=DESKTOP,
        filetypes=[
            ("Excel文件", "*.xlsx *.xls"),
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*")
        ]
    )
    
    if file_path:
        print(f"   选择文件：{file_path}")
        return file_path
    else:
        print("⚠️  未选择文件。")
        return ""


def clean_header(header: str) -> str:
    """清理表头中的BOM、制表符和前后空白字符"""
    if not header:
        return ""
    # 移除BOM字符 (\ufeff)、制表符，并去除首尾空白
    return header.replace('\ufeff', '').replace('\t', '').strip()


def read_source_file(file_path: str) -> list:
    """
    读取源Excel或CSV文件，逐行提取数据
    返回有效行的字典列表
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.csv':
        # 处理CSV文件
        with open(file_path, 'r', encoding='utf-8-sig') as f:  # utf-8-sig 自动处理BOM
            reader = csv.DictReader(f)
            # 清理表头中的制表符和空白字符
            reader.fieldnames = [clean_header(h) for h in reader.fieldnames]
            all_rows = list(reader)
    else:
        # 处理Excel文件
        wb = load_workbook(file_path)
        ws = wb.active
        
        # 获取表头行
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        headers = [clean_header(str(h)) if h is not None else "" for h in header_row]
        
        all_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_dict = dict(zip(headers, row))
            all_rows.append(row_dict)
    
    # 过滤有效订单状态，并丢弃备注为空的行
    valid_rows = []
    invalid_cnt = 0
    for row_dict in all_rows:
        order_status = str(row_dict.get("订单状态", "")).strip()
        remark = str(row_dict.get("商家备注", "")).strip()
        if order_status not in VALID_STATUSES or not remark:
            continue
        # 备注含"废除"等无效标记 → 该单并未实际发单，整条丢弃
        if any(kw in remark for kw in INVALID_REMARK_KEYWORDS):
            invalid_cnt += 1
            continue
        valid_rows.append(row_dict)

    if invalid_cnt:
        print(f"   已过滤无效备注（{'/'.join(INVALID_REMARK_KEYWORDS)}）：{invalid_cnt} 条")

    return valid_rows


def save_output(wb: Workbook, file_path: str):
    """保存输出文件"""
    wb.save(file_path)
    print(f"✅ 已保存至：{file_path}")


# ============================
# 主函数
# ============================

def main():
    print("=" * 40)
    print("  拼多多出单日报统计")
    print("=" * 40)

    # Step 1: 让用户选择输入文件
    print("\n📂 Step 1: 选择输入文件...")
    source_path = scan_lzc_files()
    if not source_path:
        print("程序退出。")
        return
    
    # 根据输入文件名生成输出文件名
    input_filename = os.path.splitext(os.path.basename(source_path))[0]
    output_filename = f"{input_filename}（已处理）.xlsx"
    OUTPUT_FILE = os.path.join(DESKTOP, output_filename)
    
    print(f"   输入文件：{source_path}")
    print(f"   输出文件：{OUTPUT_FILE}")

    # Step 2: 创建新的输出Excel文件
    print("\n📄 Step 2: 创建输出文件...")
    out_wb = create_output_file()
    out_ws = out_wb.active
    print(f"   表头数量：{len(HEADERS)} 列")

    # Step 3: 读取并过滤源数据
    print("\n📖 Step 3: 读取源文件，过滤有效订单...")
    try:
        source_rows = read_source_file(source_path)
        print(f"   有效订单行数：{len(source_rows)} 行")
    except Exception as e:
        print(f"❌ 读取文件时出错：{e}")
        return

    if not source_rows:
        print("⚠️  没有找到符合条件的订单，程序退出。")
        return

    # Step 4: 处理每行数据并写入新表
    print("\n⚙️  Step 4: 处理数据并写入新表...")
    seq = 1
    stats = {}  # 配件统计：{品类: {细分种类: 使用总数}}
    for row_data in source_rows:
        processed = process_row(row_data)
        processed["序号"] = seq
        row_values = [str(processed.get(key, "/")) for key in OUTPUT_KEYS]
        seq += 1
        out_ws.append(row_values)
        collect_stats(stats, processed)
        print(f"   [{seq - 1}] 订单号: {row_data.get('订单号', '-')} | "
              f"状态: {row_data.get('订单状态', '-')} | "
              f"备注: {str(row_data.get('商家备注', '-'))[:20]}...")

    # Step 5: 保存明细表
    print("\n💾 Step 5: 保存文件...")
    save_output(out_wb, OUTPUT_FILE)

    # Step 6: 生成配件统计表（新增功能，独立文件，不影响原明细表）
    print("\n📊 Step 6: 生成配件统计表...")
    if stats:
        stats_file = os.path.join(DESKTOP, f"{input_filename}配件统计表.xlsx")
        stats_wb = create_stats_workbook(stats)
        stats_wb.save(stats_file)
        print(f"✅ 配件统计表已保存至：{stats_file}")
    else:
        print("⚠️  没有统计到任何配件数量，跳过统计表生成。")

    print("\n🎉 处理完成！")


if __name__ == "__main__":
    main()
