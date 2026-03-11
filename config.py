import dotenv
import os
dotenv.load_dotenv()
from Atoken import Get_access_token

import logging



# 日志配置
#log_formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
log_formatter = logging.Formatter('%(message)s')
file_handler = logging.FileHandler('sharepoint_downloader.log', encoding='utf-8')
file_handler.setFormatter(log_formatter)
console_handler = logging.StreamHandler()
console_handler.setFormatter(log_formatter)
logging.basicConfig(level=logging.INFO, handlers=[file_handler])
import time
logging.info(f"{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))}")
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
#ACCESS_TOKEN = Get_access_token()
headers = {
    'Authorization': 'Bearer ' + ACCESS_TOKEN
}