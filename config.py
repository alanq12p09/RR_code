"""
集中配置文件：数据计算参数、邮件参数、调度参数统一管理。
修改配置无需改动业务代码。
"""
import os
from datetime import datetime

# =========================
# 数据计算参数
# =========================
INCLUDE_CURRENT_MONTH_IN_RECOMMENDATION = False
MIN_RECENT_MONTHS_WITH_RR = 3          # 最近4个月中至少几个月有RR（1~4）
AP_SUBGEO_MIN_RECOMMENDED_RR = 1       # AP分配后绝对值下限（件）
AP_SUBGEO_MIN_SHARE = 0.01            # AP分配后占比下限（3%）

# =========================
# 文件路径
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "Data")

PN_INFO_PATH = r"C:\Users\xuecz1\OneDrive - Lenovo\Option Share Files - DeptDocument\Power BI Group\PN Information.xlsx"
KICKOFF_PATH = r"C:\Users\xuecz1\OneDrive - Lenovo\Option Share Files - DeptDocument\Power BI Group\DataSource\Kickoff Package\Accessories Commercial&Consumer KickOff Package.xlsx"
_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

# OUTPUT_PATH = os.path.join(os.path.dirname(DATA_DIR), f"RR_Low_RR_Recommendation_{_TIMESTAMP}.xlsx")
OUTPUT_PATH = os.path.join(DATA_DIR, f"RR_Low_RR_Recommendation_{_TIMESTAMP}.xlsx") 

# =========================
# 邮件参数
# =========================
# 发送方式："outlook" = 调用本机Outlook客户端（默认）
#          "smtp"    = 通过SMTP协议发送（服务器部署时切换到这个）
MAIL_METHOD = "outlook"

# --- Outlook 方式只需要配置以下内容 ---

# 按 GEO 分发：每个 GEO 只收到自己的推荐数据（保密）
# enabled: True = 发送, False = 跳过（暂不需要发的GEO设为False）
# to: 收件人列表
# cc: 抄送列表
GEO_RECIPIENTS = {
    "AP": {
        "enabled": True,
        "to": [
            "xuecz1@lenovo.com",
            "wangying62@lenovo.com",
            "lina6@lenovo.com",
            "zhangchun6@lenovo.com",
            "bveeraraghav@lenovo.com",
            "wuguihua@lenovo.com",
        ],
        "cc": [],
    },
    "BRAZ": {
        "enabled": True,
        "to": [
            "xuecz1@lenovo.com",
            "dsilva23@lenovo.com",
            "greis1@lenovo.com",
            "wuguihua@lenovo.com",
        ],
        "cc": [],
    },
    "EMEA": {
        "enabled": True,
        "to": [
            "xuecz1@lenovo.com",
            "atusekova@lenovo.com",
            "wuguihua@lenovo.com",
        ],
        "cc": [],
    },
    "NA": {
        "enabled": True,
        "to": [
            "xuecz1@lenovo.com",
            "tangfeng1@lenovo.com",
            "ugautam1@lenovo.com",
            "wuguihua@lenovo.com",
        ],
        "cc": [],
    },
    "PRC": {
        "enabled": False,
        "to": ["prc_planner@company.com"],
        "cc": [],
    },
    "LAS": {
        "enabled": True,
        "to": [
            "xuecz1@lenovo.com",
            "cguzman@lenovo.com",
            "gmhernandez@lenovo.com",
            "wuguihua@lenovo.com",
        ],
        "cc": [],
    },
}

# GEO_RECIPIENTS = {
#     "AP": {
#         "enabled": True,
#         "to": ["xuecz1@lenovo.com,liuhw23@lenovo.com"],
#         "cc": [],
#     },
#     "BRAZ": {
#         "enabled": False,
#         "to": ["braz_planner@company.com"],
#         "cc": [],
#     },
#     "EMEA": {
#         "enabled": False,
#         "to": ["xuecz1@lenovo.com"],
#         "cc": [],
#     },
#     "LAS": {
#         "enabled": False,
#         "to": ["las_planner@company.com"],
#         "cc": [],
#     },
#     "NA": {
#         "enabled": True,
#         "to": ["xuecz1@lenovo.com,liuhw23@lenovo.com"],
#         "cc": [],
#     },
#     "PRC": {
#         "enabled": False,
#         "to": ["prc_planner@company.com"],
#         "cc": [],
#     },
# }


# 汇总报告收件人（可选，收到完整报告；留空则不发汇总）
#MAIL_TO_SUMMARY = ["xuecz1@lenovo.com ; wuguihua@lenovo.com ; zhouxx10@lenovo.com"]
MAIL_TO_SUMMARY = ["xuecz1@lenovo.com;zhouxx10@lenovo.com;wuguihua@lenovo.com;huangjin@lenovo.com;xuhy1@lenovo.com;zhangxj1@lenovo.com"]
MAIL_CC_SUMMARY = []

MAIL_SUBJECT_TEMPLATE = "[Auto] L220 Low RR Recommendation - {geo} - {data_max_month}"
MAIL_BODY_TEMPLATE = """Hi Team,

Please find the attached Low RR Recommendation for {geo}.

Summary:
- Report Date: {date}
- Data Max Month: {data_max_month}
- Recommendation Rows ({geo}): {recommendation_count}

This is an automated email. Please do not reply directly.

Best Regards,
Chaozhong Xue
"""

# --- 以下仅 MAIL_METHOD = "smtp" 时需要配置 ---
SMTP_HOST = "smtp.office365.com"        # SMTP服务器地址
SMTP_PORT = 587                         # 端口：587(TLS) / 465(SSL) / 25(无加密)
SMTP_USE_TLS = True                     # 是否使用TLS
SMTP_USER = "your_account@company.com"  # 发件人账号（留空则跳过认证）
SMTP_PASSWORD = ""                      # 密码（建议从环境变量读取）
MAIL_FROM = "your_account@company.com"  # 发件人地址

# 如果密码配置了环境变量，优先使用环境变量
if not SMTP_PASSWORD:
    SMTP_PASSWORD = os.environ.get("RR_SMTP_PASSWORD", "")

# =========================
# GEO 上传模板配置
# =========================
# 每个 GEO 的模板参数独立配置
# 新增其他 GEO 模板时：
#   1. 在 rr_recommendation.py 中新建 build_XX_upload_sheets() 函数
#   2. 在 GEO_UPLOAD_BUILDERS 字典里注册
#   3. 在这里加该 GEO 的配置项
UPLOAD_CONFIGS = {
    "AP": {
        "source": "0001000835",     # Source 固定值
        "overrides": {},            # 后门：{ "PN": { "SubGeo": 推荐值 } }
    },
    # 以后其他 GEO 需要模板时在这里加，例如：
    # "EMEA": {
    #     "source": "0002000xxx",
    #     "overrides": {},
    # },
}

# =========================
# 调度参数
# =========================
# cron表达式含义：每周一早上8点执行
SCHEDULE_CRON_DAY_OF_WEEK = "mon"
SCHEDULE_CRON_HOUR = 8
SCHEDULE_CRON_MINUTE = 0
