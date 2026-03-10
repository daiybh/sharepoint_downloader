import requests
import config
from urllib.parse import urlparse, unquote
from typing import Tuple, Optional

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
        print("URL format is incorrect, unable to find 'sites' or 'Shared Documents'")
        site_path = ""
        file_path = ""
    return domain, site_path, file_path

def get_site_id(domain: str, site_path: str) -> Optional[str]:
    url = f'https://graph.microsoft.com/v1.0/sites/{domain}:/{site_path}'
    try:
        response = requests.get(url, headers=config.headers)
        response.raise_for_status()
        return response.json().get('id')
    except Exception as e:
        print(f"Failed to get siteID: {e}")
        return None

def get_drive_id(site_id: str) -> Optional[str]:
    url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives'
    try:
        response = requests.get(url, headers=config.headers)
        response.raise_for_status()
        drives = response.json().get('value', [])
        if drives:
            return drives[0]['id']
    except Exception as e:
        print(f"Failed to get driveID: {e}")
    return None

def get_file_info(site_id: str, drive_id: str, file_path: str) -> Optional[dict]:
    url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{file_path}'
    try:
        response = requests.get(url, headers=config.headers)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Failed to get file info: {e}")
        return None

def download_file(download_url: str, local_path: str):
    try:
        print(f"\nDownloading {local_path} from URL: {download_url}\n")
        with requests.get(download_url, stream=True) as r:
            r.raise_for_status()
            with open(local_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        print(f"Download completed: {local_path}")
    except Exception as e:
        print(f"Download failed: {e}")

def download_from_sharepoint(shared_url: str, save_dir: str = '.'):
    domain, site_path, file_path = split_url(shared_url)
    if not (domain and site_path and file_path):
        print("URL parsing failed, aborting download.")
        return
    dest_file = unquote(file_path).split('/')[-1]
    site_id = get_site_id(domain, site_path)
    if not site_id:
        print("Unable to get siteID, aborting download.")
        return
    drive_id = get_drive_id(site_id)
    if not drive_id:
        print("Unable to get driveID, aborting download.")
        return
    file_info = get_file_info(site_id, drive_id, file_path)
    if not file_info or '@microsoft.graph.downloadUrl' not in file_info:
        print("Unable to get download link, aborting download.")
        return
    download_url = file_info['@microsoft.graph.downloadUrl']
    local_path = f"{save_dir}/downloadUrl_{dest_file}"
    download_file(download_url, local_path)

if __name__ == "__main__":
    # Example usage
    shared_url = 'https://riedelcommunications.sharepoint.com/:u:/r/sites/SimplyLiveInternal/Shared%20Documents/R%26D/VideoEngine/TcTableAnalyzer/11.26.4.5/TcTableAnalyzer11.26.4.5.zip?csf=1&web=1&e=MnZxu8'
    download_from_sharepoint(shared_url)
