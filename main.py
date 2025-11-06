import asyncio
import logging
import pytz
import os

from datetime import datetime, timezone, time
from logging.handlers import TimedRotatingFileHandler
from telegram import Bot
from automation_check import main
from dotenv import load_dotenv

load_dotenv()


class KSTFormatter(logging.Formatter):
    def converter(self, timestamp):
        dt = datetime.fromtimestamp(timestamp)
        kst = pytz.timezone('Asia/Seoul')
        return dt.replace(tzinfo=timezone.utc).astimezone(kst)

    def formatTime(self, record, datefmt=None):
        dt = self.converter(record.created)
        if datefmt:
            return dt.strftime(datefmt)
        return dt.strftime('%Y-%m-%d %H:%M:%S')

def setup_logger(name):
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    log_dir = 'logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # 파일 핸들러
    log_file = os.path.join(log_dir, f'{name}.log')
    file_handler = TimedRotatingFileHandler(
        log_file,
        when='midnight',
        interval=1,
        backupCount=30,
        encoding='utf-8',
        atTime=time(hour=0, minute=0, second=0)
    )
    file_handler.suffix = "%Y-%m-%d"
    
    # KST 포매터 적용
    kst_formatter = KSTFormatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    )
    file_handler.setFormatter(kst_formatter)
    
    # 콘솔 핸들러에도 동일한 KST 포매터 적용
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(kst_formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

logger = setup_logger('market_automation_check')

async def send_telegram_alert(error_message):
    bot_token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    
    try:
        bot = Bot(token=bot_token)  # telegram.Bot이 아닌 Bot으로 사용
        await bot.send_message(
            chat_id=chat_id,
            text=f"🚨 에러 발생!\n{error_message}"
        )
    except Exception as e:
        logger.error(f"Telegram 알림 전송 실패: {e}")

async def run_with_retry(max_retries=3):
    for attempt in range(max_retries):
        try:
            return await main(logger=logger, send_alert=send_telegram_alert)
        except Exception as e:
            logger.error(f"Attempt {attempt + 1}/{max_retries} failed: {e}")
            logger.exception("상세 에러:")
            if attempt < max_retries - 1:
                await asyncio.sleep(60)
            else:
                raise

async def scheduler():
    kst = pytz.timezone('Asia/Seoul')
    while True:
        try:
            start_time = datetime.now(timezone.utc).astimezone(kst)
            logger.info(f"Starting execution at {start_time}")
            
            orders = await run_with_retry()
            logger.info(f"Processed orders: {orders}")
            
            logger.info(f"Completed execution at {datetime.now(timezone.utc).astimezone(kst)}")
            await asyncio.sleep(1800)
            
        except Exception as e:
            error_msg = f"Automation Check critical error occurred: {e}"
            logger.error(error_msg)
            logger.exception("상세 에러:")
            await send_telegram_alert(error_msg)
            await asyncio.sleep(60)

if __name__ == "__main__":
    loop = asyncio.get_event_loop()
    try:
        logger.info("서비스 시작")
        loop.run_until_complete(scheduler())
    except KeyboardInterrupt:
        logger.info("Scheduler stopped by user")
    except Exception as e:
        logger.error(f"Scheduler stopped due to error: {e}")
        logger.exception("상세 에러:")
    finally:
        logger.info("서비스 종료")
        loop.close()