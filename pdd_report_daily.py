# pdd_report_daily.py - 拼多多出单日报统计
# 功能：读取拼多多订单Excel，按状态过滤，提取字段并生成规范数据表

import os
import sys
import csv
from tkinter import Tk, filedialog
from openpyxl import Workbook, load_workbook
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
    "序号", "日期", "地址", "规格/颜色", "尺寸", "总米数",
    "封口", "封口个数", "滑轮", "滑轮个数", "膨胀螺丝/套",
    "免钉胶/套", "安装码", "安装码数量", "连接器", "连接器数量",
    "配件", "配件数量", "备注", "订单号", "快递单号/订单编号"
]

OUTPUT_KEYS = [
    "序号", "日期", "地址", "规格/颜色", "尺寸", "总米数",
    "封口", "封口个数", "滑轮", "滑轮个数", "膨胀螺丝/套",
    "免钉胶/套", "安装码", "安装码数量", "连接器", "连接器数量",
    "配件", "配件数量", "备注", "订单号", "快递单号/订单编号"
]

# 有效订单状态（只保留这些状态的行）
VALID_STATUSES = {"已发货，待收货", "已发货，退款成功", "已收货", "已收货，退款成功"}

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
        "F-4", "Z21", "Z22", "002", "006", "219", "大方轨封口"
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
      规格/颜色   → 开头型号【颜色】，如 "219【白】"
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
            "规格/颜色": "/", "尺寸": "/", "总米数": "/",
            "封口": "/", "个数": "/", "滑轮": "/", "个数": "/",
            "膨胀螺丝/套": "/", "安装码": "/", "数量": "/",
            "连接器": "/", "数量": "/", "配件": "/", "数量": "/",
        }

    import re
    # 关键：把换行符等空白统一成空格，避免 \S 跨行匹配导致数值错位
    text = re.sub(r"\s+", " ", remark.strip())

    # --- 规格/颜色：取【】中的颜色，拼上型号前缀 ---
    color_match = re.search(r"【(.+?)】", text)
    color = color_match.group(1) if color_match else ""
    if color_match:
        prefix = text[:color_match.start()].strip()
        spec_color = f"{prefix}【{color}】" if prefix else f"【{color}】"
    else:
        spec_color = "/"

    # --- 尺寸 & 总米数：提取所有 "X米*Y根"，保留原分隔符格式 ---
    size_pattern = re.compile(r"([\d.]+)米\s*\*\s*(\d+)\s*根")
    size_matches = size_pattern.findall(text)
    if size_matches:
        sizes = [f"{m}米*{n}根" for m, n in size_matches]
        size_str = "\n".join(sizes)
        total_m = sum(float(m) * int(n) for m, n in size_matches)
        total_m_rounded = round(total_m, 2)
        total_m_str = str(int(total_m_rounded)) if total_m_rounded == int(total_m_rounded) else str(total_m_rounded)
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
            color = seal_match.group(2) or ""
            seal_spec = f"{seal_name}{('【' + color + '】') if color else ''}"
            seal_qty = parse_quantity(seal_match.group(3))
        else:
            seal_spec, seal_qty = "/", "/"

    # --- 安装码 ---
    dan_spec, dan_qty = _find_named_count(patterns.get("安装码", []), text)
    if dan_spec is None:
        dan_match = None
        for pattern in [
            # 支持带颜色标签的格式：单侧码【黑】20个
            r"([^\s]+?)\s*(单侧码)(?:【([^】]*)】)?\s*(\d+)\s*个",
            r"([^\s]+?)\s*(双侧码)(?:【([^】]*)】)?\s*(\d+)\s*个",
            r"([^\s]+?)\s*(单顶码)(?:【([^】]*)】)?\s*(\d+)\s*个",
            r"([^\s]+?)\s*(双顶码)(?:【([^】]*)】)?\s*(\d+)\s*个",
        ]:
            dan_match = re.search(pattern, text)
            if dan_match:
                break
        if dan_match:
            # group(1): 前缀, group(2): 关键字类型, group(3): 颜色, group(4): 数量
            prefix = dan_match.group(1).strip()
            keyword = dan_match.group(2)
            color = dan_match.group(3) or ""  # 颜色可能为空
            dan_qty = parse_quantity(dan_match.group(4))
            dan_spec = f"{prefix} {keyword}{('【' + color + '】') if color else ''}"
        else:
            dan_spec, dan_qty = "/", "/"

    # --- 滑轮 ---
    wheel_names = patterns.get("滑轮", []) + ["轴承轮", "合金轮"]
    lun_spec, lun_qty = _find_named_count(wheel_names, text)
    if lun_spec is None:
        # 支持带颜色标签的格式：10 钢片两轮【白】88个
        wheel_match = None
        for pattern in [
            # 支持数字+空格+关键词 或 数字汉字混合（如"7字纳米滑块"），末尾可选颜色标签
            r"(\d+\s*[\w\u4e00-\u9fff]*?纳米[\w\u4e00-\u9fff]*?)(?:【([^】]*)】)?\s*(\d+)\s*个",
            r"(\d+\s*[\w\u4e00-\u9fff]*?钢片[\w\u4e00-\u9fff]*?)(?:【([^】]*)】)?\s*(\d+)\s*个",
            r"(\d+\s*[\w\u4e00-\u9fff]*?轴承轮[\w\u4e00-\u9fff]*?)(?:【([^】]*)】)?\s*(\d+)\s*个",
            r"(\d+\s*[\w\u4e00-\u9fff]*?合金轮[\w\u4e00-\u9fff]*?)(?:【([^】]*)】)?\s*(\d+)\s*个",
            r"(\d+\s*[\w\u4e00-\u9fff]*?滑块[\w\u4e00-\u9fff]*?)(?:【([^】]*)】)?\s*(\d+)\s*个",
            r"([\w\u4e00-\u9fff]*?走珠[\w\u4e00-\u9fff]*?)(?:【([^】]*)】)?\s*(\d+)\s*个",
        ]:
            wheel_match = re.search(pattern, text)
            if wheel_match:
                break
        if wheel_match:
            lun_name = wheel_match.group(1).strip()
            color = wheel_match.group(2) or ""
            lun_spec = f"{lun_name}{('【' + color + '】') if color else ''}"
            lun_qty = parse_quantity(wheel_match.group(3))
        else:
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
    # 支持带颜色标签的格式：连接器【黑】10个
    lian_match = re.search(r"连接器(?:【([^】]*)】)?\s*([一二三四五六七八九十\d]+)\s*个", text)
    if lian_match:
        color = lian_match.group(1) or ""
        lian_spec = f"连接器{('【' + color + '】') if color else ''}"
        lian_qty = parse_quantity(lian_match.group(2))
    else:
        lian_spec, lian_qty = "/", "/"

    # --- 配件 ---
    pei_spec, pei_qty = _find_named_count(patterns.get("配件", []), text)
    if pei_spec is None:
        pei_spec, pei_qty = "/", "/"

    return {
        "规格/颜色": spec_color,
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
        "规格/颜色":         _f(parsed.get("规格/颜色", "")),
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
    for row_dict in all_rows:
        order_status = str(row_dict.get("订单状态", "")).strip()
        remark = str(row_dict.get("商家备注", "")).strip()
        if order_status in VALID_STATUSES and remark:
            valid_rows.append(row_dict)
    
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
    for row_data in source_rows:
        processed = process_row(row_data)
        processed["序号"] = seq
        row_values = [str(processed.get(key, "/")) for key in OUTPUT_KEYS]
        seq += 1
        out_ws.append(row_values)
        print(f"   [{seq - 1}] 订单号: {row_data.get('订单号', '-')} | "
              f"状态: {row_data.get('订单状态', '-')} | "
              f"备注: {str(row_data.get('商家备注', '-'))[:20]}...")

    # Step 5: 保存
    print("\n💾 Step 5: 保存文件...")
    save_output(out_wb, OUTPUT_FILE)
    print("\n🎉 处理完成！")


if __name__ == "__main__":
    main()
