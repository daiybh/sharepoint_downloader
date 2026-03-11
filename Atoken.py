import requests
import os

import logging

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
    return response.json()["access_token"]