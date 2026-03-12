
import requests
import logging
from urllib.parse import urlparse, unquote
from typing import Tuple, Optional
import dotenv
import os
dotenv.load_dotenv()

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

def Get_access_token():
    """获取访问令牌"""
    url = f"https://login.microsoftonline.com/{os.getenv("AZURE_TENANT_ID")}/oauth2/v2.0/token"
    data = {
        "client_id": os.getenv("AZURE_CLIENT_ID"),
        "client_secret": os.getenv("AZURE_CLIENT_SECRET"),
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    logging.info("#"*20)
    logging.info("获取访问令牌")
    logging.info(url)
    logging.info(data)
    
    response = requests.post(url, data=data)
    response.raise_for_status()
    logging.info(response.json())
    return response.json()["access_token"]

def split_url(shared_url: str) -> Tuple[str, str, str]:
    """
    Parse SharePoint shared link and extract domain, site path, and file path.
    """
    parsed = urlparse(shared_url)
    domain = parsed.netloc
    path_parts = parsed.path.strip('/').split('/')
    try:
        sites_index = path_parts.index('sites')
        site_path = "/".join(path_parts[sites_index:sites_index+2])
        shared_docs_index = path_parts.index('Shared%20Documents')
        file_path = "/".join(path_parts[shared_docs_index+1:])
    except ValueError:
        logging.error("URL format is incorrect, unable to find 'sites' or 'Shared Documents'")
        site_path = ""
        file_path = ""
    return domain, site_path, file_path

def pretty_print(title,msg):
    
    logging.info("#"*20)
    logging.info(f"{title}")
    logging.info(f"{msg}")

def httpGet(url):
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        logging.error(f"Failed to get: {e}")
        return None
def get_site_id(domain: str, site_path: str) -> Optional[str]:
    url = f'https://graph.microsoft.com/v1.0/sites/{domain}:/{site_path}'
    pretty_print("get_site_id", url)
    try:
        response = httpGet(url)
        return response.get('id') if response else None
    except Exception as e:
        logging.error(f"Failed to get siteID: {e}")
        return None

def get_drive_id(site_id: str) -> Optional[str]:
    url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives'
    pretty_print("get_drive_id", url)
    try:
        response = httpGet(url)
        drives = response.get('value', []) if response else []
        if drives:
            return drives[0]['id']
    except Exception as e:
        logging.error(f"Failed to get driveID: {e}")
    return None

def get_file_info(site_id: str, drive_id: str, file_path: str) -> Optional[dict]:
    url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{file_path}'
    pretty_print("get_file_info", url)
    try:
        response = httpGet(url)
        return response.json() if response else None
    except Exception as e:
        logging.error(f"Failed to get file info: {e}")
        return None

def download_file(download_url: str, local_path: str):
    try:
        logging.info(f"Downloading {local_path} from URL: {download_url}")
        with requests.get(download_url, stream=True) as r:
            r.raise_for_status()
            with open(local_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        logging.info(f"Download completed: {local_path}")
    except Exception as e:
        logging.error(f"Download failed: {e}")

def download_from_sharepoint(shared_url: str, save_dir: str = '.'):
    domain, site_path, file_path = split_url(shared_url)
    if not (domain and site_path and file_path):
        logging.error("URL parsing failed, aborting download.")
        return
    dest_file = unquote(file_path).split('/')[-1]
    site_id = get_site_id(domain, site_path)
    if not site_id:
        logging.error("Unable to get siteID, aborting download.")
        return
    drive_id = get_drive_id(site_id)
    if not drive_id:
        logging.error("Unable to get driveID, aborting download.")
        return
    file_info = get_file_info(site_id, drive_id, file_path)
    if not file_info or '@microsoft.graph.downloadUrl' not in file_info:
        logging.error("Unable to get download link, aborting download.")
        return
    download_url = file_info['@microsoft.graph.downloadUrl']
    logging.info(f"@microsoft.graph.downloadUrl {download_url}")
    local_path = f"{save_dir}/downloadUrl_{dest_file}"
    download_file(download_url, local_path)

if __name__ == "__main__":
    # Example usage
    shared_url = 'https://riedelcommunications.sharepoint.com/:u:/r/sites/SimplyLiveInternal/Shared%20Documents/R%26D/VideoEngine/TcTableAnalyzer/11.26.4.5/TcTableAnalyzer11.26.4.5.zip?csf=1&web=1&e=MnZxu8'
    download_from_sharepoint(shared_url)
