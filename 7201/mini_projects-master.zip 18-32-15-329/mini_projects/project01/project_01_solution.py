# =============================================================================
# project_01_solution.py（简化版 · 行级基础校验 · 详细注释版）
#
# 这份代码只做一件事：扫描 8 个 CSV 事件日志，把“每一行自己就存在的问题”找出来，
# 然后把这些行的原始 event_id 写进 problematic_event_ids.txt（每行一个，无表头/空行/重复）。
#
# 覆盖的检查分为两大块：
# 【A. 行级检查】——“每一行自己就存在的问题”，逐行独立判断，不需要跨行对比：
#   1. 必填字段缺失      —— 每行都必须有的 5 个字段，缺了就报
#   2. 格式不规范        —— ID、时间、金额、姓名、电话等写法不符合规范就报
#   3. 取值不合法        —— 数量、单价、运费、折扣、金额等数值范围不对就报
#   4. 各事件类型的必填项 —— 不同事件类型各自要求填的字段，缺了就报
#   5. 创建行的价格公式  —— 下单行的金额不等于“单价×数量-折扣+运费”就报
#
# 【B. 跨行检查】——必须“读遍全部 8 个文件、把数据拼成一张大表”后才能判断：
#   6. 重复上报检测      —— 同一笔交易被记录了多次，只留最小 ID，其余重复行上报
#   7. 订单级业务冲突    —— 同一订单内：平台/币种/客户信息要一致、事件时序不能乱、
#                            签收要有对应发货、累计退款不能超过累计收款
#
# 【一个贯穿全文的简化思路】
#   B 部分要做“跨行比较”，理论上得先把各种不规范写法（中文、大小写、多余符号）
#   折算成标准值再比，代码会很啰嗦。但因为——任何写法不规范的行，早在 A 部分就被
#   上报过了——所以 B 部分我们只拿“写法干净、规范”的行来做比较即可，直接省掉那套
#   折算代码，逻辑清爽很多，也不会漏报（脏行已经在结果里了）。
#
# 设计原则：一个检查就写一个小函数，函数只返回“哪些行有问题”的布尔标记或 ID 集合；
#           最后把所有结果合并。这样每一步都能单独读懂，互不纠缠。
# =============================================================================


# =============================================================================
# 【pandas 基础概念速查】——看代码前先花一分钟扫一遍，后面就不卡了
# -----------------------------------------------------------------------------
# pandas 是 Python 处理“表格数据”的库，本文件里它有两个核心数据结构：
#
#   1) DataFrame（数据框）：就是一整张表，有行有列，像 Excel 里的一张工作表。
#      代码里变量 events / batch / df 都是 DataFrame。
#      写法 events["event_id"] 表示“从 events 这张表里，把 event_id 这一整列取出来”。
#
#   2) Series（序列）：单独一列就是一个 Series。
#      你可以把它理解成“一竖排数据 + 一个行号索引(index)”。
#      events["event_id"] 取出来的就是 Series。
#
# 常用操作（本文件反复出现）：
#   • .str.xxx         ：对“字符串类型的 Series”逐格做文本处理。
#                        例如 .str.strip() 是把每格文字的首尾空格去掉。
#                        （.str 就是把 Python 字符串的那些方法，批量作用到整列上）
#   • == / !=          ：整列和某个值逐个比较，结果是“一列 True/False”。
#   • .isin(集合)      ：判断每格的值“是否属于这个集合”，返回一列 True/False。
#   • .str.match(正则) ：判断每格文字“是否匹配某个正则规则”，返回一列 True/False。
#   • .notna() / .isna()：判断每格“有没有值 / 是不是缺失(NaN)”，返回一列 True/False。
#   • ~xxx             ：取反，把 True 变 False、False 变 True（“非”）。
#   • a | b            ：两个布尔列按位“或”，只要其一为 True 结果就 True。
#   • a & b            ：两个布尔列按位“且”，两者都 True 结果才 True。
#   • loc[布尔列, 列名] ：按“布尔列”筛选行，再取指定列。
#                        例如 events.loc[bad, "event_id"] = 只保留 bad 为 True 的行的 event_id。
#   • 布尔列本质        ：我们常把“一列 True/False”称作“掩码(mask)”，True=命中/有问题。
#
# 关于 NaN：NaN 是 pandas 表示“空/缺失/算不出来”的特殊值，任何比较碰到它通常是 False。
# =============================================================================


# 用 Path 以面向对象方式拼接文件路径（模板的所有路径都基于它）
from pathlib import Path
# 用 ZipFile 在缺少已解压 CSV 时，从 Datasets.zip 里按需解出 8 个 CSV
from zipfile import ZipFile
# 用 re（正则表达式）判断各列是不是“规范写法”
import re
# pandas：读 CSV、构造表格、按列做批量判断；import x as pd 是给 pandas 起个短别名 pd
import pandas as pd


# -----------------------------------------------------------------------------
# 第 0 部分：文件路径（沿用模板，评测平台依赖这些固定路径与输出文件名，别改）
# -----------------------------------------------------------------------------
# BASE_DIR = 当前脚本所在的目录。
# Path(__file__) 把“本文件路径字符串”包成 Path 对象；
# .resolve() 解析成绝对路径（补全 .. 等）；.parent 取上一级，也就是文件所在目录。
BASE_DIR = Path(__file__).resolve().parent
# 用 `/` 运算符把目录和文件名拼成完整路径（Path 重载了除号当“路径拼接”）。
ARCHIVE_PATH = BASE_DIR / "Datasets.zip"          # 数据压缩包路径
DATA_DIR = BASE_DIR / "Datasets"                   # 解压后 8 个 CSV 存放的目录名
# 列表推导式：i 取 1..8，一次性拼出 events_part1.csv ~ events_part8.csv 的完整路径清单。
# range(1, 9) 表示从 1 到 8（左闭右开，不含 9）。
CSV_FILES = [DATA_DIR / f"events_part{i}.csv" for i in range(1, 9)]
# 最终结果输出文件路径
OUTPUT_PATH = BASE_DIR / "problematic_event_ids.txt"


# -----------------------------------------------------------------------------
# 第 1 部分：“规范值”集合 与 “规范格式”正则
#   规范值 = 用集合(set)列出允许的取值，判断时问“在不在集合里”；
#   规范格式 = 用正则表达式描述允许的写法，判断时问“匹不匹配”。
# -----------------------------------------------------------------------------
# 合法平台（必须大写、无空格）。{} 直接写成花括号的是“集合”，成员之间用逗号分隔。
SOURCE_CANON = {"MEITUAN", "JD", "TAOBAO", "PDD", "XIANYU"}
CURRENCY_CANON = {"CNY", "USD", "EUR"}                     # 合法币种
PAY_CANON = {"VISA", "MASTERCARD", "PAYPAL", "ALIPAY", "WECHAT_PAY"}  # 合法支付/退款方式
# 合法事件类型（大小写/空格/中文/同义词都算不规范）
EVENT_TYPE_CANON = {"ORDER_CREATED", "PAYMENT_SUCCESS", "PAYMENT_FAILED",
                    "SHIPPED", "DELIVERED", "REFUND", "CANCELLED"}

# -----------------------------------------------------------------------------
# 【正则表达式逐字符拆解】—— r"..." 里 r 前缀表示“原始字符串”，反斜杠 \ 不用额外转义。
#   常见符号含义：
#     ^   匹配“开头”          $   匹配“结尾”
#     \d  任意一个数字(0-9)    \d+ 一个或多个数字      \d{10} 正好 10 个数字
#     ()  分组               |    或（二选一）         .    一个小数点（这里做了转义）
#   所以整个正则读作：“从头到尾必须是……这一种写法”，否则就不匹配。
# -----------------------------------------------------------------------------
RE_EVENT_ID = re.compile(r"^EV\d+$")     # 事件号：开头 EV + 至少一位数字 + 结尾，如 EV200101
RE_ORDER_ID = re.compile(r"^ORD\d+$")    # 订单号：ORD + 数字，如 ORD2001
RE_CUST_ID = re.compile(r"^C\d+$")       # 客户号：C + 数字，如 C2001
RE_PROD_ID = re.compile(r"^SKU\d+$")     # 商品号：SKU + 数字，如 SKU2001
# 交易号：开头是 TX 或 RF（(TX|RF) 的 | 表示二选一）+ 数字，如 TX200101 / RF200101
RE_TX_ID = re.compile(r"^(TX|RF)\d+$")
RE_SHP_ID = re.compile(r"^SHP\d+$")      # 包裹号：SHP + 数字，如 SHP2001
# 电话：以 1 开头(\d{10} 再跟正好 10 个数字)，共 11 位纯数字；不能带 +86/空格/短横
RE_TEL = re.compile(r"^1\d{10}$")
# 时间：4位年-2位月-2位日 T 2位时:2位分:2位秒 +08:00。
#   \d{4} 四位数字，\- 这里写成 - 直接是连字符，T 是字面字母 T，\+08:00 是字面 "+08:00"。
RE_TIME = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\+08:00$")
# 金额：至少一位数字，后面可选地跟“小数点+至少一位数字”。
#   (\.\d+)? 里的 ? 表示“这一段可有可无”，所以 60 和 60.10 都匹配，但 -2、CNY 60 不匹配。
RE_NUM = re.compile(r"^\d+(\.\d+)?$")
# 数量：一位或多位数字、且仅此而已（不含小数点、单位、符号），如 2、10；不匹配 1.0、-2、"1件"。
RE_INT = re.compile(r"^\d+$")


# -----------------------------------------------------------------------------
# 第 2 部分：两个常用小工具（判空、把文本转成数字）
# -----------------------------------------------------------------------------
def is_blank(col):
    """判断一列里每个格子是不是“空的”。空串、纯空格都算空。返回一列 True/False。"""
    # col.str.strip()：对整列逐个把“首尾空白字符”去掉；
    # 再 == ""：看去掉后是不是空串。是空串 → True（代表这格是空的）。
    # 于是 "  "（全是空格）strip 后变 ""，也会被判为 True。
    return col.str.strip() == ""


def to_num(col):
    """把金额/数量这种文本列安全地转成数字。
    能转的转成小数；不能转的（像 "CNY 60"、"-2"、空串）转成 NaN（缺失）。
    注意：转不了的脏值本来已在“格式检查”里被上报，后面做取值判断时遇到 NaN 跳过即可。
    """
    # pd.to_numeric：把 Series 逐个转成数字；
    # errors="coerce"：遇到转不动的值不报错，而是置为 NaN。
    return pd.to_numeric(col, errors="coerce")


# -----------------------------------------------------------------------------
# 第 3 部分：两个“填了但不对”的通用小工具 + 各个独立的小检查。
#   小检查函数都接收整张事件表 events（DataFrame），
#   返回一个和行数一样长的 True/False 列（Series 掩码），True=该行有此类问题。
# -----------------------------------------------------------------------------
def bad_format(events, col, pattern):
    """【工具】找出“某列填了值、但写法不符合规范”的行。
    用在 ID、金额、数量、电话这类“有固定写法格式”的列上。
    返回一列 True/False，True 表示“这行这列填了，却写错了”。
    分三步理解：
      1) filled —— 这一格到底填没填（非空）。没填的行不归本工具管（由“必填检查”负责）。
      2) matches —— 填的内容符不符合给定写法（正则 pattern）。
      3) 返回 —— “填了” 且 “不符合” 才算这行有问题。
    """
    # ~is_blank(...)：is_blank 给的是“为空”，前面加 ~（取反）就变成“非空/填了值”
    filled = ~is_blank(events[col])
    # .str.match(pattern)：逐格套用正则，写法符合返回 True
    matches = events[col].str.match(pattern)
    # filled & ~matches：填了值 且 写法不对（~matches） → 这行有问题
    return filled & ~matches


def bad_choice(events, col, allowed):
    """【工具】找出“某列填了值、但取值不在允许清单里”的行。
    用在平台/币种/事件类型/支付方式这类“只能从固定几个标准值里选”的列上。
    返回一列 True/False。
    分三步理解：
      1) filled —— 填没填；
      2) in_allowed —— 填的值是否恰好在允许清单 allowed 里（用 .isin 逐个判断）；
      3) 返回 —— “填了” 且 “不在清单里” 才算问题。
    """
    # 填了值（非空）
    filled = ~is_blank(events[col])
    # .isin(allowed)：逐格判断值是否属于集合 allowed，属于返回 True
    in_allowed = events[col].isin(allowed)
    # 填了 但 不在清单里（~in_allowed） → 问题
    return filled & ~in_allowed


def check_missing_common(events):
    """检查 3.1：每行都必须填的 5 个字段，有没有缺失。
    这 5 个是：event_id、order_id、event_type、event_time、source_system。
    只要其中任意一个为空，这行就算有问题。
    """
    # 把“每行都必须填”的列名写成一个清单，后面循环逐列检查
    always_required = ["event_id", "order_id", "event_type", "event_time", "source_system"]
    # pd.Series(False, index=...)：先造一列“全 False”，长度=行数，行号与 events 对齐
    bad = pd.Series(False, index=events.index)
    # 逐个必填列检查：哪一行这列是空的(is_blank)，就把那一行标成 True
    # bad = bad | is_blank(...)：把“这一列的空行”并入总问题行（| 是逐格“或”）
    for col in always_required:
        bad = bad | is_blank(events[col])
    return bad


def check_id_format(events):
    """检查 3.2-a：各种 ID 的写法规不规范（只在“填了值”时才判写法）。
    规则：某个 ID 只要有值，但不是“大写前缀 + 纯数字”这种标准写法，就算问题。
    （空不空不归这里管，由“必填检查”负责；这里只管“写了但写错了”。）
    """
    # 字典：列名 -> 该列应遵循的写法正则。把六种 ID 集中列出，一目了然
    id_rules = {
        "event_id": RE_EVENT_ID,
        "order_id": RE_ORDER_ID,
        "customer_id": RE_CUST_ID,
        "product_id": RE_PROD_ID,
        "transaction_id": RE_TX_ID,
        "shipment_id": RE_SHP_ID,
    }
    # 先造一列“全 False”用于累积
    bad = pd.Series(False, index=events.index)
    # 逐种 ID 调用工具 bad_format（“填了但写法不对”），把结果并入
    for col, pattern in id_rules.items():
        bad = bad | bad_format(events, col, pattern)
    return bad


def check_category_format(events):
    """检查 3.2-b：分类字段（平台/事件类型/币种/支付方式）的取值合不合法。
    规则：这些字段只能从“标准值清单”里选，写法必须完全一致。
    空值这里不判（交给必填检查）；填了但不在清单里，就算不规范。
    """
    # 字典：列名 -> 该列允许的取值集合
    choice_rules = {
        "source_system": SOURCE_CANON,
        "event_type": EVENT_TYPE_CANON,
        "currency": CURRENCY_CANON,
        "payment_method": PAY_CANON,
    }
    bad = pd.Series(False, index=events.index)
    # 逐个字段调用工具 bad_choice（“填了但取值不在清单里”），并入结果
    for col, allowed in choice_rules.items():
        bad = bad | bad_choice(events, col, allowed)
    return bad


def check_time_format(events):
    """检查 3.2-c：事件时间 event_time 是否规范。
    一条合格时间要同时满足：填了 + 格式对 + 年份是 2025。
    任何一条不满足，就算问题（下面分成三种情况分别标出）。
    """
    # 先把时间这一列取出来，省得下面反复写 events["event_time"]
    time_col = events["event_time"]
    # 情况一：压根没填
    empty = is_blank(time_col)
    # format_ok：格式对不对（匹配“带 +08:00 的 ISO 写法”正则）；~format_ok 即“格式错”
    format_ok = time_col.str.match(RE_TIME)
    wrong_format = ~format_ok
    # .str.slice(0, 4) 取每格前 4 个字符（就是年份），再 == "2025" 逐格比较
    year_ok = time_col.str.slice(0, 4) == "2025"
    # 情况三：格式虽然对、但年份不是 2025（readme 要求所有时间落在 2025）
    wrong_year = format_ok & ~year_ok
    # 三种情况任一成立就算问题
    return empty | wrong_format | wrong_year


def check_number_format(events):
    """检查 3.2-d：数量/金额这些“本该是数字”的列，写法对不对（非空时才判）。
    像 "CNY 60.10"、"-74.45"、"1件"、"1.0" 这类都算写法不规范。
    """
    # 数量列：写法必须是“纯整数”（RE_INT）
    bad = bad_format(events, "quantity", RE_INT)
    # 金额类四列：写法必须是“非负小数”（RE_NUM）
    for col in ["unit_price", "discount_amount", "shipping_fee", "amount"]:
        bad = bad | bad_format(events, col, RE_NUM)
    return bad


def check_name_format(events):
    """检查 3.2-e：客户姓名 name 的写法（只在填了的情况下判）。
    合格姓名：不该有首尾空格，也不该中间出现连续两个空格。
    """
    # 取出姓名这一列
    name = events["name"]
    # 只有“填了姓名”的行才判（没填的交给必填检查）
    filled = ~is_blank(name)
    # name.str.strip() 是把首尾空格去掉后的版本；若去掉后和原来不相等(!=)，说明原本有多余首尾空格
    edge_space = name.str.strip() != name
    # .str.contains("  ", regex=False)：逐格查“有没有两个连续空格”；regex=False 表示按普通文本查
    double_space = name.str.contains("  ", regex=False)
    # 两种多余空格任一出现就算“写法有问题”
    bad_space = edge_space | double_space
    # 填了 且 写法有问题 → 才算这行的问题
    return filled & bad_space


def check_tel_format(events):
    """检查 3.2-f：电话 tel 的写法（非空时才判）。
    合格电话是“1 开头的 11 位纯数字”；带 +86、空格、短横的都算不规范。
    """
    # 直接套用工具：填了值 但 不匹配规范电话格式 → 问题
    return bad_format(events, "tel", RE_TEL)


def check_values(events):
    """检查 3.3：数值范围是否合理（readme 3.2 的价格/取值约束）。
    前提：这里只看“能干净转成数字”的值；写法乱的脏值（转成 NaN）已在格式检查里报过，
    所以每步都先用 .notna() 把“有有效数字的行”挑出来，只对这些行做范围判断。
    规则：
      - 数量必须是正整数；单价必须 > 0；运费、折扣必须 >= 0；
      - 折扣不能超过小计（数量×单价）；
      - 下单/支付/退款这些事件的金额必须 > 0。
    """
    # 把相关列转成数字列（转不了的脏值变 NaN）
    quantity = to_num(events["quantity"])
    unit_price = to_num(events["unit_price"])
    discount = to_num(events["discount_amount"])
    shipping = to_num(events["shipping_fee"])
    amount = to_num(events["amount"])

    # 数量：有值(notna) 且（<=0 或 不是整数）→ 非法
    #   quantity.round() 把数量四舍五入成整数；若原值 != 取整值，说明它带小数、不是整数
    bad_quantity = quantity.notna() & ((quantity <= 0) | (quantity != quantity.round()))
    # 单价：有值且 <= 0 → 非法（单价必须大于 0）
    bad_unit_price = unit_price.notna() & (unit_price <= 0)
    # 运费：有值且 < 0 → 非法（运费不能为负）
    bad_shipping = shipping.notna() & (shipping < 0)
    # 折扣：有值且 < 0 → 非法（折扣不能为负）
    bad_discount_negative = discount.notna() & (discount < 0)
    # 折扣：有值 且 折扣 > 数量×单价 → 折扣超过了小计，非法（+0.005 容忍浮点误差）
    bad_discount_too_big = (
        discount.notna() & unit_price.notna() & quantity.notna()
        & (discount > quantity * unit_price + 0.005)
    )
    # 哪些事件要求“金额必须为正”（发货/签收/取消的 amount 允许为空，不受此约束）
    amount_must_positive = events["event_type"].isin(
        ["ORDER_CREATED", "PAYMENT_SUCCESS", "PAYMENT_FAILED", "REFUND"]
    )
    # 属于这些事件、金额有值、且 <= 0 → 非法
    bad_amount = amount_must_positive & amount.notna() & (amount <= 0)

    # 把上面所有取值问题用 | 合并成一个总布尔列
    return (
        bad_quantity | bad_unit_price | bad_shipping | bad_discount_negative
        | bad_discount_too_big | bad_amount
    )


def check_required_by_event_type(events):
    """检查 3.4：不同事件类型各自要求填写的字段（只判事件类型写得规范的那些行）。
    如果 event_type 本身就写错了，它已在格式检查里被报了，这里就不重复处理。
    """
    # 字典(dict)：键是“事件类型”，值是“该类型必须填的字段清单(list)”。
    required_map = {
        "ORDER_CREATED": ["customer_id", "name", "tel", "product_id", "quantity",
                          "unit_price", "discount_amount", "shipping_fee", "amount", "currency"],
        "PAYMENT_SUCCESS": ["amount", "currency", "payment_method", "transaction_id"],
        "PAYMENT_FAILED": ["amount", "currency", "payment_method", "transaction_id"],
        "REFUND": ["amount", "currency", "payment_method", "transaction_id"],
        "SHIPPED": ["shipment_id"],
        "DELIVERED": ["shipment_id"],
        "CANCELLED": [],  # 取消事件没有额外的必填字段
    }
    # pd.Series(False, index=events.index)：造一列“全 False”的布尔掩码，
    #   index 与 events 完全对齐（行号一一对应），长度和行数相同。
    bad = pd.Series(False, index=events.index)
    # .items()：把字典按 (键, 值) 成对遍历出来
    for event_type, cols in required_map.items():
        # 这一行事件类型是否等于当前类型（逐格 == 比较，得到布尔列）
        is_this_type = events["event_type"] == event_type
        # 遍历该类型的每个必填字段
        for col in cols:
            # “是该类型” 且 “该字段为空” → 问题；bad = bad | (...) 把结果并入
            bad = bad | (is_this_type & is_blank(events[col]))
    return bad


def check_price_formula(events):
    """检查 3.5：下单行 ORDER_CREATED 的金额是否等于价格公式结果。
    公式：amount = 数量×单价 - 折扣 + 运费。
    只有当事件类型是 ORDER_CREATED 且五个数都能转成数字时，才核对公式。
    """
    # 相关列转数字
    quantity = to_num(events["quantity"])
    unit_price = to_num(events["unit_price"])
    discount = to_num(events["discount_amount"])
    shipping = to_num(events["shipping_fee"])
    amount = to_num(events["amount"])
    # 挑出下单行（布尔列）
    is_created = events["event_type"] == "ORDER_CREATED"
    # 五个数都成功解析(notna) 才有资格核对公式
    all_parseable = (
        quantity.notna() & unit_price.notna() & discount.notna()
        & shipping.notna() & amount.notna()
    )
    # 算出“实际金额 - 公式金额”的差，.abs() 取绝对值；<=0.005 视为公式成立（容忍浮点误差）
    formula_ok = (amount - (quantity * unit_price - discount + shipping)).abs() <= 0.005
    # 是下单行 & 数字都能解析 & 公式不成立(~formula_ok) → 问题
    return is_created & all_parseable & ~formula_ok


# -----------------------------------------------------------------------------
# 第 4 部分：把所有小检查合并成“这行到底有没有问题”
# -----------------------------------------------------------------------------
def find_problem_rows(events):
    """把第 3 部分的每个检查结果叠在一起，返回所有有问题行的原始 event_id 集合(set)。"""
    # 一列和行数对齐、初始全 False 的布尔掩码，用来累积各类问题
    bad = pd.Series(False, index=events.index)
    # 逐个把每类检查结果“或”进来：只要任一检查命中，bad 对应行就变 True
    bad = bad | check_missing_common(events)        # 必填字段缺失
    bad = bad | check_id_format(events)            # ID 写法
    bad = bad | check_category_format(events)      # 分类取值写法
    bad = bad | check_time_format(events)          # 时间格式与年份
    bad = bad | check_number_format(events)        # 金额/数量写法
    bad = bad | check_name_format(events)          # 姓名多余空格
    bad = bad | check_tel_format(events)           # 电话写法
    bad = bad | check_values(events)               # 数值范围
    bad = bad | check_required_by_event_type(events)  # 各事件类型必填项
    bad = bad | check_price_formula(events)        # 下单价格公式
    # events.loc[bad, "event_id"]：按 bad 掩码只保留“有问题的行”，取它们的 event_id 列；
    # set(...)：转成集合，自动去重（同一 ID 若在多块/多行出现，只留一份）。
    return set(events.loc[bad, "event_id"])


# -----------------------------------------------------------------------------
# 第 4-B 部分：跨行检查（需要传进来的是“全部 8 个文件拼成的一张大表”）
# -----------------------------------------------------------------------------
def find_duplicate_ids(events):
    """检查 6：重复上报检测。返回“多余重复行”的原始 event_id 集合。

    什么叫重复：同一笔交易（transaction_id 相同），除了 event_id 以外
    所有业务字段都一模一样的多行，就是被重复上报了。
    怎么留谁：按 readme 3.3，保留 event_id 最小那一条，其余“多出来的”上报。

    简化说明：这里直接拿“原始值”逐列比较（不做中英文/大小写折算）。
    因为写法不规范的行早在行级检查里就被报了，这里只负责把“完全一模一样的重复”挑出来。
    """
    # 只有填了交易号的行才可能是“交易重复”（is_blank 判空，加 ~ 取反得到“非空”）
    has_tx = ~is_blank(events["transaction_id"])
    # .copy()：拷一份子集，避免后续排序时对原表报警（副本改动不影响外部）
    sub = events[has_tx].copy()
    # 少于 2 行不可能有重复，直接返回空集
    if len(sub) < 2:
        return set()
    # “参与比较的列” = 除 event_id 以外的所有列（业务内容完全一致才算重复）
    # events.columns 是所有列名，列表推导式把 event_id 排除掉
    compare_cols = [c for c in events.columns if c != "event_id"]
    # 先按 event_id 升序排序：这样每组里“第一条”就是 ID 最小、应保留的那条
    # kind="mergesort" 是稳定排序（相同值保持原有相对顺序）
    sub = sub.sort_values("event_id", kind="mergesort")
    # .duplicated(subset=..., keep="first")：在 compare_cols 这些列上找重复，
    #   每组第一条返回 False（保留），后续重复行返回 True（即“多余上报行”）
    dup = sub.duplicated(subset=compare_cols, keep="first")
    # 取出被标记为重复(多余)行的原始 event_id，返回集合
    return set(sub.loc[dup, "event_id"])


def _build_order_work(events):
    """【准备步骤】把全量事件表整理成一张“工作表”，只留跨行比较用得上的列。

    把订单比作一个包裹：同一个 order_id 会多条“扫描记录”（创建/支付/发货/签收…）。
    我们只把“订单号、事件类型都写成规范格式”的行收进来（写法不规范的脏行已在行级检查里报了）。
    返回的 work 是一个新 DataFrame，列名都缩短为后面好用的名字。
    """
    # 两个布尔列：订单号是否规范、事件类型是否规范
    ok_order = events["order_id"].str.match(RE_ORDER_ID)
    ok_type = events["event_type"].isin(EVENT_TYPE_CANON)
    # 时间列：只有格式规范的才能安全转成时间戳；非规范的变 NaT（时间版的缺失）
    ts = pd.to_datetime(
        events["event_time"].where(events["event_time"].str.match(RE_TIME)), errors="coerce")
    # 组装工作表（都从原表取，保持同一行顺序）
    work = pd.DataFrame({
        "raw_eid": events["event_id"].values,      # 原始 event_id，最后上报用
        "order": events["order_id"].values,        # 分组用的订单号
        "etype": events["event_type"].values,      # 事件类型
        "source": events["source_system"].values,  # 平台
        "currency": events["currency"].values,     # 币种
        "cust": events["customer_id"].values,      # 客户号
        "name": events["name"].values,             # 姓名
        "tel": events["tel"].values,               # 电话
        "ship": events["shipment_id"].values,      # 包裹号
        "ts": ts.values,                           # 时间戳
        "amount": to_num(events["amount"]).values,  # 金额数值（脏值→NaN）
    })
    # 只保留“订单号与事件类型都规范”的行；重排行号方便后面按位置对齐
    work = work[ok_order.values & ok_type.values].reset_index(drop=True)
    return work


def _check_info_consistency(work):
    """【子检查 7.1】同一个订单里，平台/币种/客户信息（只要填了）必须和“创建行”一致。

    【这是什么意思——举个具体例子】
    订单 ORD100 在表里有 4 行事件（它们 order 都是 ORD100）：
        创建行  平台=MEITUAN
        支付行  平台=MEITUAN   ← 和创建行一样，正常
        发货行  平台=JD         ← 订单在美团创建，发货却变成京东，矛盾！→ 上报发货行
    “创建行”写下的是订单最初的标准答案，后面任何一行如果这列填了、又和答案不同，就是冲突。

    【为什么代码不按订单循环？】
    全量有几百万个订单，一个订单一个订单地循环会慢到几十分钟、超评测限时。
    所以这里用“批量”做法（对整张表一次算完），步骤永远是这三步：
      第 1 步 答案 ：算出每个订单的“标准答案”（来自创建行）。
      order
        ORD100    MEITUAN
        ORD200    JD
        Name: source, dtype: object          ← ORD300 没有创建行，不在字典里
      第 2 步 贴回旁边 ：把该订单的标准答案，复制标注到属于它的每一行旁边。
      行0: work["order"][0]="ORD100" → 查到 "MEITUAN"
        行1: work["order"][1]="ORD100" → 查到 "MEITUAN"
        行2: work["order"][2]="ORD100" → 查到 "MEITUAN"
        行3: work["order"][3]="ORD200" → 查到 "JD"
        行4: work["order"][4]="ORD200" → 查到 "JD"
        行5: work["order"][5]="ORD300" → 查不到 → NaN
      第 3 步 逐行比较 ：本行的值 与 旁边的标准答案 不一样（且本行非空）→ 冲突。
      行号	order	source（本行自己填的）	answer_per_row（该订单的标准答案）
        0	ORD100	MEITUAN	MEITUAN
        1	ORD100	MEITUAN	MEITUAN
        2	ORD100	JD	MEITUAN ← 不一样！
        3	ORD200	JD	JD
        4	ORD200	JD	JD
        5	ORD300	TB	NaN（无创建行，不作判定）
    """
    problems = set()
    # 先挑出所有订单的“创建行”（它们记录了订单最初的正确信息）
    created = work[work["etype"] == "ORDER_CREATED"]
    # 对这五列分别做上面的“三步”（用一个循环，避免重复写五遍）
    for col in ["source", "currency", "cust", "name", "tel"]:
        # 第 1 步：“每个订单的标准答案”= 该订单创建行在这一列上的值。
        #   groupby("order")[col].first()：按订单号分组，每组取第一个创建行的值。
        answer = created.groupby("order")[col].first()
        # 第 2 步：把标准答案按订单号“贴”到每一行旁边，得到一列“每行对应的答案”。
        #   work["order"].map(answer)：拿本行的订单号，去 answer 里查它对应的标准答案。
        answer_per_row = work["order"].map(answer)
        # 第 3 步：该订单确实有答案(notna) 且 本行这列填了(!="") 且 与答案不同 → 冲突。
        bad = answer_per_row.notna() & (work[col] != "") & (work[col] != answer_per_row)
        # 把这轮找到的冲突行加入结果集
        problems |= set(work.loc[bad, "raw_eid"])
    return problems


def _check_sequence(work):
    """【子检查 7.2】同一个订单里，事件发生的时间先后不能乱。

    【正常顺序】先“创建”，再“支付/发货/签收”，中途可能“退款”或“取消”。
    两种典型的乱序：
      (a) 太早：在“创建”之前就有支付/发货等。
          例：ORD100 创建时间是 3-01 10:00，但有一笔支付写在 3-01 09:00（比创建还早）→ 不可能，上报。
      (b) 太晚：订单已经“取消”了，之后却又发生支付/发货/签收。
          例：3-02 取消了，3-03 又发货 → 矛盾，上报。（退款可以发生在取消后，不算冲突）

    同样批量处理（不按订单循环），三步：
      第 1 步：算出每个订单“最早的创建时间”与“最早的取消时间”。
      第 2 步：把它们贴回每一行旁边。
      第 3 步：本行时间比创建还早 / 比取消还晚，且属于受约束的类型 → 冲突。
    """

    #     行号	raw_eid	order	etype	ts（时间戳）
    # 0	EV101	ORD100	ORDER_CREATED	3-01 10:00
    # 1	EV102	ORD100	PAYMENT_SUCCESS	3-01 09:00 ⚠️
    # 2	EV103	ORD100	SHIPPED	3-02 08:00
    # 3	EV104	ORD100	CANCELLED	3-03 12:00
    # 4	EV105	ORD100	DELIVERED	3-04 09:00 ⚠️
    # 5	EV201	ORD200	PAYMENT_SUCCESS	3-05 10:00
    # 6	EV202	ORD200	ORDER_CREATED	3-05 12:00 ⚠️
    problems = set()
    created = work[work["etype"] == "ORDER_CREATED"]
    # 第 1 步：每个订单最早的“创建时间”与最早的“取消时间”（.min() 在时间戳上取最早）
    #     order
    # ORD100   3-01 10:00
    # ORD200   3-05 12:00
    # Name: ts, dtype: datetime64
    first_create = created.groupby("order")["ts"].min()
    # EV104 → 3-03 12:00
    first_cancel = work[work["etype"] == "CANCELLED"].groupby("order")["ts"].min()
    # 第 2 步：把这两个参考时间按订单号贴到每一行旁边
    #     行号	order	create_t	cancel_t
    # 0	ORD100	3-01 10:00	3-03 12:00
    # 1	ORD100	3-01 10:00	3-03 12:00
    # 2	ORD100	3-01 10:00	3-03 12:00
    # 3	ORD100	3-01 10:00	3-03 12:00
    # 4	ORD100	3-01 10:00	3-03 12:00
    # 5	ORD200	3-05 12:00	NaT（查不到）
    # 6	ORD200	3-05 12:00	NaT
    create_t = work["order"].map(first_create)
    cancel_t = work["order"].map(first_cancel)

    # 第 3 步 (a) 太早：这些类型都不应早于创建时间。
    #   逐行条件：属于这些类型 且 本行时间存在 且 创建时间存在 且 本行时间 < 创建时间
    #     行	etype	本行 ts	create_t	ts < create_t?	结论
    # 1	PAYMENT_SUCCESS	3-01 09:00	3-01 10:00	✅ 是	命中 → EV102
    # 2	SHIPPED	3-02 08:00	3-01 10:00	否	正常
    # 3	CANCELLED	3-03 12:00	3-01 10:00	否	正常
    # 4	DELIVERED	3-04 09:00	3-01 10:00	否	正常
    # 5	PAYMENT_SUCCESS	3-05 10:00	3-05 12:00	✅ 是	命中 → EV201
    must_after_create = work["etype"].isin(
        ["PAYMENT_SUCCESS", "PAYMENT_FAILED", "REFUND", "SHIPPED", "DELIVERED", "CANCELLED"])
    too_early = (must_after_create & work["ts"].notna()
                 & create_t.notna() & (work["ts"] < create_t))
    problems |= set(work.loc[too_early, "raw_eid"])

    # 第 3 步 (b) 太晚：取消之后不应再发生这些类型（注意：不包含 REFUND，退款可在取消后）
    # 行	etype	本行 ts	cancel_t	ts > cancel_t?	结论
    # 1	PAYMENT_SUCCESS	3-01 09:00	3-03 12:00	否	不命中
    # 2	SHIPPED	3-02 08:00	3-03 12:00	否	不命中
    # 4	DELIVERED	3-04 09:00	3-03 12:00	✅ 是	命中 → EV105
    # 5	PAYMENT_SUCCESS	3-05 10:00	NaT	—	cancel_t.notna() 为 False，跳过
    must_before_cancel = work["etype"].isin(
        ["PAYMENT_SUCCESS", "PAYMENT_FAILED", "SHIPPED", "DELIVERED"])
    too_late = (must_before_cancel & work["ts"].notna()
                & cancel_t.notna() & (work["ts"] > cancel_t))
    problems |= set(work.loc[too_late, "raw_eid"])
    return problems


def _check_delivery_needs_shipment(work):
    """【子检查 7.3】“签收”必须先有对应的“发货”，且包裹号要对得上。

    思路：先汇总每个订单“发过哪些包裹号”，再看每条签收记录的包裹号在不在里面。
    """
    #行号	raw_eid	order	etype	ship（包裹号）
    # 0	EV301	ORD100	ORDER_CREATED	""
    # 1	EV302	ORD100	SHIPPED	SHP1001
    # 2	EV303	ORD100	SHIPPED	SHP1002
    # 3	EV304	ORD100	DELIVERED	SHP1001 ✅ 有对应发货
    # 4	EV305	ORD100	DELIVERED	SHP9999 ❌ 没发过这个包裹
    # 5	EV306	ORD200	SHIPPED	SHP2001
    # 6	EV307	ORD200	DELIVERED	"" ❌ 签收却没填包裹号
    # 7	EV308	ORD300	DELIVERED	SHP3001 ❌ 整个订单从没发货
    # 每个订单已发货行用到的包裹号集合：groupby 后用 apply(set) 把同组包裹号归成一个集合
    shipped_ids = work[work["etype"] == "SHIPPED"].groupby("order")["ship"].apply(set)
    #     order
    # ORD100    {SHP1001, SHP1002}      ← 行1、行2 的包裹号归成集合
    # ORD200    {SHP2001}
    # Name: ship, dtype: object
    # 把本订单的“已发货包裹号集合”广播到每一行（没发货记录的订单会得到 NaN）
    allowed = work["order"].map(shipped_ids)
    #  行号	order	allowed（本订单发过的包裹集合）
    # 0	ORD100	{SHP1001, SHP1002}
    # 1	ORD100	{SHP1001, SHP1002}
    # 2	ORD100	{SHP1001, SHP1002}
    # 3	ORD100	{SHP1001, SHP1002}
    # 4	ORD100	{SHP1001, SHP1002}
    # 5	ORD200	{SHP2001}
    # 6	ORD200	{SHP2001}
    # 7	ORD300	NaN（字典里查不到）
    # 逐行判断签收行的包裹号是否在该集合里：
    #   allowed 不是集合(没发货) → False；ship 不在集合里 → False；否则 True
    matched = [isinstance(a, set) and s in a for a, s in zip(allowed, work["ship"])]
    #  行号	a（allowed）	s（ship）	isinstance(a, set)?	s in a?	matched
    # 0	{SHP1001, SHP1002}	""	是	"" 不在集合里	False
    # 1	{SHP1001, SHP1002}	SHP1001	是	是	True
    # 2	{SHP1001, SHP1002}	SHP1002	是	是	True
    # 3	{SHP1001, SHP1002}	SHP1001	是	是	True ✅
    # 4	{SHP1001, SHP1002}	SHP9999	是	否	False ❌
    # 5	{SHP2001}	SHP2001	是	是	True
    # 6	{SHP2001}	""	是	否	False ❌
    # 7	NaN	SHP3001	否（没发货记录）	—	False ❌
    # 把上面这个普通列表变回与 work 行号对齐的布尔列
    matched = pd.Series(matched, index=work.index)
    # 是签收行 且 包裹号没匹配上(~matched) → 上报
    # 行号	etype	~matched	两者都成立？
    # 0	ORDER_CREATED	True	否（不是签收行，放行）
    # 1、2、5	SHIPPED	False	否
    # 3	DELIVERED	False	否（匹配成功，正常签收）
    # 4	DELIVERED	True	✅ EV305
    # 6	DELIVERED	True	✅ EV307
    # 7	DELIVERED	True	✅ EV308
    is_delivered = work["etype"] == "DELIVERED"
    return set(work.loc[is_delivered & ~matched, "raw_eid"])


def _check_refund_over_payment(work):
    """【子检查 7.4】按时间累加，任意时刻累计退款不能超过累计成功收款。

    思路：把每笔收款当“进账(+)”、退款当“出账(-)”，按时间顺序累加；
    一旦某次退款后余额变负，说明退得比收得多，这笔退款就是超额。
    """
#行号	raw_eid	order	etype	ts	amount
# 0	EV401	ORD100	PAYMENT_SUCCESS	3-01 10:00	100
# 1	EV402	ORD100	REFUND	3-02 09:00	60
# 2	EV403	ORD100	REFUND	3-03 14:00	50
# 3	EV404	ORD200	PAYMENT_SUCCESS	3-01 08:00	200
# 4	EV405	ORD200	REFUND	3-02 08:00	200
# 5	EV406	ORD200	REFUND	3-03 08:00	10
# 6	EV407	ORD300	REFUND	3-05 10:00	30 ← 该订单一分钱都没收过就直接退
# 7	EV408	ORD100	SHIPPED	3-04 10:00	"" ← 发货行，和钱无关
# 8	EV409	ORD200	PAYMENT_FAILED	3-01 09:00	200 ← 支付失败，不算进账
    # 只看“成功收款/退款”、且时间与金额都能解析的行（收款和退款才影响余额）
    money = work[work["etype"].isin(["PAYMENT_SUCCESS", "REFUND"])
                 & work["ts"].notna() & work["amount"].notna()].copy()
    # .where(条件, 其他值)：条件为真保留本行金额，为假则取另一个值。
    # “收款记正、退款记负”，方便后面累加
    # .where(条件, 其他值)：条件为真保留原金额，为假取括号里的另一个值。即收款为正、退款为负
    # raw_eid	order	etype	amount	delta
    # EV401	ORD100	PAYMENT_SUCCESS	100	+100
    # EV402	ORD100	REFUND	60	−60
    # EV403	ORD100	REFUND	50	−50
    # EV404	ORD200	PAYMENT_SUCCESS	200	+200
    # EV405	ORD200	REFUND	200	−200
    # EV406	ORD200	REFUND	10	−10
    money["delta"] = money["amount"].where(money["etype"] == "PAYMENT_SUCCESS", -money["amount"])
    # 按 (订单, 时间) 排序，保证“同一订单按时间先后”累加
    money = money.sort_values(["order", "ts"], kind="mergesort")
    # 每个订单内部累加 delta；余额为负即退款超过了收款
    # cumsum 是组内累加
    #     时间顺序	交易	delta	balance（滚动余额）
    # 3-01 10:00	EV401 收款 100	+100	100
    # 3-02 09:00	EV402 退款 60	−60	40 ✅ 还收着 40，退得合理
    # 3-03 14:00	EV403 退款 50	−50	−10 ❌ 只剩 40 却退了 50！
    #  时间顺序	交易	delta	balance
    # 3-01 08:00	EV404 收款 200	+200	200
    # 3-02 08:00	EV405 退款 200	−200	0 ✅ 全额退，恰好退完，不超
    # 3-03 08:00	EV406 退款 10	−10	−10 ❌ 已退完又退 10
    # 时间顺序	交易	delta	balance
    # 3-05 10:00	EV407 退款 30	−30	−30 ❌ 从没收过钱就退款
    money["balance"] = money.groupby("order")["delta"].cumsum()
    # “退款行 且 此时余额<0” → 超额退款，上报（-0.005 容忍浮点误差）
    #  raw_eid	etype	balance	是退款行?	balance < −0.005?	上报?
    # EV402	REFUND	40	是	否	否
    # EV403	REFUND	−10	是	是
    over = (money["etype"] == "REFUND") & (money["balance"] < -0.005)
    return set(money.loc[over, "raw_eid"])


def find_order_conflict_ids(events):
    """【总入口】订单级业务冲突：先把工作表准备好，再依次跑四个子检查，把结果并起来。

    这个函数只负责“调度”，具体逻辑都分到四个 _check_xxx 小函数里，每个只看一件短。
    """
    # 第 0 步：准备一张只含规范行的工作表
    work = _build_order_work(events)
    # 把四个子检查的结果并集到一个集合里
    problems = set()
    problems |= _check_info_consistency(work)        # 7.1 信息一致性
    problems |= _check_sequence(work)                # 7.2 时序
    problems |= _check_delivery_needs_shipment(work) # 7.3 签收需对应发货
    problems |= _check_refund_over_payment(work)     # 7.4 累计退款不超额
    return problems


# -----------------------------------------------------------------------------
# 第 5 部分：读取数据的生成器（把每个 CSV 切成小块产出，避免一次性全塞内存）
# -----------------------------------------------------------------------------
def iter_event_batches(csv_files, chunksize=None):
    """把 8 个 CSV 变成“批次流”：每次产出一个 DataFrame 小块。
    chunksize=None 时每个文件整块产出；给了数值时按该行数切块。
    用了 yield 的函数叫“生成器”：要一个给一个，内存里始终只留一小块。
    """
    # 逐个遍历文件路径
    for csv_path in csv_files:
        # dtype="string"：所有列按字符串读，保留 ID/电话/脏值原样；
        # na_filter=False：空单元格保持为空串""而不是 NaN，方便统一按空串判空；
        # encoding="utf-8-sig"：自动去掉文件开头的 BOM（一种隐藏字节标记）。
        options = dict(dtype="string", na_filter=False, encoding="utf-8-sig")
        if chunksize is None:
            # 整读：pd.read_csv 直接返回整个文件的 DataFrame，用 yield 交给调用方
            yield pd.read_csv(csv_path, **options)
        else:
            # 分块：给了 chunksize 后 read_csv 返回一个“逐块读取器”，用 with 保证读完自动关闭
            with pd.read_csv(csv_path, chunksize=chunksize, **options) as reader:
                # yield from reader：把读取器里的每一块依次透传给调用方
                yield from reader


# =============================================================================
# 主流程：读数据 → 检查 → 写结果
#   下面这段用 if __name__ == "__main__": 包起来，意思是“只有直接运行本文件时才执行”，
#   被别的文件 import 时不会自动跑，方便单独调用上面那些函数做测试。
# =============================================================================
if __name__ == "__main__":
    # %% 1. 读数据：若 8 个 CSV 没全就位，就从 Datasets.zip 里按需解压
    # all(...) 判断 8 个文件是否都在；只要有一个缺，not all 为 True，进入解压分支。
    if not all(csv_path.is_file() for csv_path in CSV_FILES):
        # 缺 CSV 又没有压缩包 → 直接报错，提示把文件放哪
        if not ARCHIVE_PATH.is_file():
            raise FileNotFoundError(f"Place Datasets.zip or all eight extracted CSV files in: {BASE_DIR}")
        print("Extracting Datasets.zip...")
        # 以只读方式打开压缩包，with 块结束会自动关闭
        with ZipFile(ARCHIVE_PATH) as archive:
            # 逐个把目标 CSV 解压到 BASE_DIR（压缩包内路径为 Datasets/文件名）
            for csv_path in CSV_FILES:
                archive.extract(f"Datasets/{csv_path.name}", path=BASE_DIR)
    # 解压后再核对一次，还有缺的就报错
    missing_files = [str(p) for p in CSV_FILES if not p.is_file()]
    if missing_files:
        raise FileNotFoundError(f"Missing CSV files: {missing_files}")
    print("All 8 CSV files are ready.")

    # %% 2. 分析：先把 8 个 CSV 整读拼成一张全量大表，再跑行级 + 跨行两类检查。
    # 说明：重复检测与订单级冲突必须“看遍全部 8 个文件”才能判断，因此需要完整数据；
    #       评测平台提供 16GiB 内存，足以容纳这张全量表。
    # pd.concat(...)：把 8 个 DataFrame 纵向拼接成一张大表；ignore_index=True 重排连续行号。
    events = pd.concat(iter_event_batches(CSV_FILES, chunksize=None), ignore_index=True)
    print(f"Total rows loaded into memory: {len(events):,}")

    # 第一类：行级检查（逐行独立的格式/必填/取值/价格公式），返回问题 ID 集合
    problematic_event_ids = find_problem_rows(events)
    print(f"After row-level checks: {len(problematic_event_ids):,} IDs")

    # 第二类：重复上报检测（跨行/跨文件），结果并集进总集合（|= 是“并集赋值”）
    problematic_event_ids |= find_duplicate_ids(events)
    print(f"After duplicate checks: {len(problematic_event_ids):,} IDs")

    # 第三类：订单级业务冲突（跨行/跨文件），同样并集进总集合
    problematic_event_ids |= find_order_conflict_ids(events)
    print(f"After order-level checks: {len(problematic_event_ids):,} IDs")

    # %% 3. 输出：每行一个原始 event_id，去重并排序，无表头/空行
    # sorted(...)：把集合排序成列表，保证输出稳定；集合本身已去重。
    output_ids = sorted(problematic_event_ids)
    # OUTPUT_PATH.open("w",...)：以覆盖写打开文件；encoding="utf-8"；
    # newline="\n" 强制 Unix 换行风格，跨平台字节一致，评测按行切分不会残留 \r。
    with OUTPUT_PATH.open("w", encoding="utf-8", newline="\n") as output_file:
        for event_id in output_ids:
            # 每个 ID 写一行（手动补换行符 \n）
            output_file.write(f"{event_id}\n")
    # f"...{len(...):,}" 里的 :, 表示“千分位逗号”格式化大数字，便于阅读
    print(f"Saved {len(output_ids):,} problematic event IDs to: {OUTPUT_PATH}")
