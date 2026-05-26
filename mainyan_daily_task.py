import logging
import logging.handlers
import schedule
import subprocess
import time
from pathlib import Path

LOG_DIR = Path(__file__).resolve().parent / "log"
LOG_DIR.mkdir(exist_ok=True)

_file_handler = logging.handlers.TimedRotatingFileHandler(
    LOG_DIR / "mainyan_daily_task.log",
    when="midnight",
    backupCount=30,
    encoding="utf-8",
)
_file_handler.suffix = "%Y-%m-%d"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[_file_handler, logging.StreamHandler()],
)
logger = logging.getLogger(__name__)


def run_task(days):
    logger.info(f"开始执行任务: days={days}")
    try:
        result = subprocess.run(
            [r"C:\Users\Wadec\AppData\Local\Programs\Python\Python311\python.exe", Path(__file__).resolve().parent / "product_request_export.py", "--headless", f"--days={days}"],
            timeout=1800,
        )
        if result.returncode != 0:
            logger.error(f"任务异常退出，返回码: {result.returncode}")
        else:
            logger.info(f"任务完成: days={days}")
    except subprocess.TimeoutExpired:
        logger.error(f"任务超时 (30分钟): days={days}")
    except Exception as e:
        logger.error(f"任务执行异常: {e}")


schedule.every().day.at("07:50").do(run_task, days=1)
schedule.every().day.at("12:50").do(run_task, days=2)

logger.info("调度器启动")
logger.info("  07:50 -> days=1")
logger.info("  12:50 -> days=2")

while True:
    schedule.run_pending()
    time.sleep(30)