"""
调度入口：定时执行数据计算 + 邮件发送。
支持两种模式：
  1. 手动执行：python scheduler.py --run-now
  2. 定时运行：python scheduler.py（按 config.py 中的 cron 配置循环执行，需要 apscheduler）

也可以不用 apscheduler，直接用系统调度工具：
  - Windows 任务计划程序：python scheduler.py --run-now
  - Linux crontab: 0 8 * * 1 cd /path/to/RR_Forecast && python scheduler.py --run-now
"""
import sys
import logging
import argparse
from datetime import datetime

import config
from rr_recommendation import run_recommendation
from mail_sender import send_recommendation_email

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(
            f"rr_scheduler_{datetime.now():%Y%m}.log",
            encoding="utf-8"
        ),
    ]
)
logger = logging.getLogger(__name__)


def job():
    """单次执行：计算 -> 导出 -> 发邮件"""
    logger.info("=" * 60)
    logger.info("开始执行 Low RR Recommendation 任务")
    logger.info("=" * 60)

    try:
        # Step 1: 数据计算 & 导出Excel
        result = run_recommendation(
            pn_info_path=config.PN_INFO_PATH,
            output_path=config.OUTPUT_PATH,
            kickoff_path=config.KICKOFF_PATH,
            include_current_month=config.INCLUDE_CURRENT_MONTH_IN_RECOMMENDATION,
            min_recent_months=config.MIN_RECENT_MONTHS_WITH_RR,
            ap_subgeo_min_rr=config.AP_SUBGEO_MIN_RECOMMENDED_RR,
            ap_subgeo_min_share=config.AP_SUBGEO_MIN_SHARE,
            upload_configs=config.UPLOAD_CONFIGS,
        )

        logger.info(
            f"计算完成 | Qualified: {result['qualified_count']:,} | "
            f"Rows: {result['recommendation_count']:,}"
        )

        # Step 2: 按 GEO 分别发送邮件
        if getattr(config, "GEO_RECIPIENTS", {}):
            send_recommendation_email(result, config)
            logger.info("邮件发送完成")
        else:
            logger.warning("未配置 GEO_RECIPIENTS，跳过邮件发送")

    except Exception:
        logger.exception("任务执行失败")
        # 可在此处接入告警（企业微信、钉钉、Slack webhook 等）

    logger.info("任务结束")
    logger.info("=" * 60)


def run_with_scheduler():
    """使用 APScheduler 定时执行"""
    try:
        from apscheduler.schedulers.blocking import BlockingScheduler
        from apscheduler.triggers.cron import CronTrigger
    except ImportError:
        logger.error(
            "定时模式需要 apscheduler 库，请先安装: pip install apscheduler\n"
            "或者使用系统 cron 配合 --run-now 参数"
        )
        sys.exit(1)

    scheduler = BlockingScheduler()

    trigger = CronTrigger(
        day_of_week=config.SCHEDULE_CRON_DAY_OF_WEEK,
        hour=config.SCHEDULE_CRON_HOUR,
        minute=config.SCHEDULE_CRON_MINUTE,
    )

    scheduler.add_job(job, trigger=trigger, id="rr_recommendation_job")

    logger.info(
        f"调度器已启动 | 执行时间: 每周{config.SCHEDULE_CRON_DAY_OF_WEEK} "
        f"{config.SCHEDULE_CRON_HOUR:02d}:{config.SCHEDULE_CRON_MINUTE:02d}"
    )
    logger.info("按 Ctrl+C 停止")

    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        logger.info("调度器已停止")


def main():
    parser = argparse.ArgumentParser(description="Low RR Recommendation 调度器")
    parser.add_argument(
        "--run-now",
        action="store_true",
        help="立即执行一次（不启动定时调度）"
    )

    args = parser.parse_args()

    if args.run_now:
        logger.info("模式: 立即执行")
        job()
    else:
        logger.info("模式: 定时调度")
        run_with_scheduler()


if __name__ == "__main__":
    main()
