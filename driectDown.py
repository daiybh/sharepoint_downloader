
import requests
import logging
import dotenv
import os
dotenv.load_dotenv()

import coloredlogs

coloredlogs.install(level='DEBUG',fmt=' %(message)s')


def setup2_adminconsent():
    url=f'https://login.microsoftonline.com/{os.getenv("AZURE_TENANT_ID")}/adminconsent?client_id={os.getenv("AZURE_CLIENT_ID")}&state=12345&redirect_uri=https://localhost/myapp/permissions'
    logging.info("\nsetup2_adminconsent  \n\t"+url)
    response = requests.get(url)
    response.raise_for_status()
    if response.status_code != 200:
        logging.error(f"\t❌ Failed to setup admin consent: {response.status_code} ")
        logging.error(response.text)
    else:
        logging.info(f"\t✅ Setup admin consent successfully ")


def Get_access_token():
    """Get_access_token"""
    url = f"https://login.microsoftonline.com/{os.getenv("AZURE_TENANT_ID")}/oauth2/v2.0/token"
    data = {
        "client_id": os.getenv("AZURE_CLIENT_ID"),
        "client_secret": os.getenv("AZURE_CLIENT_SECRET"),
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    
    response = requests.post(url, data=data)
    response.raise_for_status()
    #logging.info(response.json())
    if response.status_code != 200:
        logging.error(f"\t❌ Failed to get access token: {response.status_code} ")
        logging.error(response.json())
    else:
        logging.info(f"\t✅ Get_access_token successfully ")
    logging.info("")
    return response.json()["access_token"]

#headers = config.headers
def testGet(url,headers):
    try:
        logging.info("\n BEGIN TEST \n\t"+url)
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        r = response.json()
        if response.status_code != 200:
            logging.info(r)
            logging.error(f"\t❌ Failed testGet: {response.status_code} ")
        else:
            logging.info(f"\t✅ testGet successfully")
            return r
    except Exception as e:
        logging.error(f"\t❌ Failed testGet: \n\t{e}")
        return None

setup2_adminconsent()


url=[

'https://graph.microsoft.com/v1.0/sites/riedelcommunications.sharepoint.com:/sites/SimplyLiveInternal',
'https://graph.microsoft.com/v1.0/sites/riedelcommunications.sharepoint.com,2f37d60a-2b81-4a88-9dee-288b5fc259f2,16909fc8-ad87-4a9f-96c2-22dcad480a93/drives',
'https://graph.microsoft.com/v1.0/sites/riedelcommunications.sharepoint.com,2f37d60a-2b81-4a88-9dee-288b5fc259f2,16909fc8-ad87-4a9f-96c2-22dcad480a93/drives/b!CtY3L4EriEqd7iiLX8JZ8sifkBaHrZ9KlsIi3K1ICpPR4Y7ssFseQpJlF9TBR2Yi/root:/R%26D/VideoEngine/TcTableAnalyzer/11.26.4.5/TcTableAnalyzer11.26.4.5.zip'
]

logging.warning("#"*20+"Use Get_access_token() "+"#"*20)
headers = {
        'Authorization': 'Bearer ' + Get_access_token()
    }

for u in url:
    testGet(u,headers)

logging.info("\n\n")
logging.warning("#"*20+"Use getenv(\"ACCESS_TOKEN\") "+"#"*20)
headers = {
    'Authorization': 'Bearer ' + os.getenv("ACCESS_TOKEN")
}
for u in url:
    testGet(u,headers)