import dotenv
import os
dotenv.load_dotenv()

ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")



headers = {
    'Authorization': 'Bearer ' + ACCESS_TOKEN

}