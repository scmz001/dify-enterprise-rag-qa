import os
import requests
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("DIFY_API_KEY")
base_url = os.getenv("DIFY_BASE_URL")

print(f"API Key 前缀：{api_key[:10] if api_key else '未读取到'}...")
print(f"Base URL：{base_url}")

url = f"{base_url}/chat-messages"
headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json"
}
data = {
    "inputs": {},
    "query": "年假几天",
    "response_mode": "blocking",
    "user": "test_1"
}

try:
    response = requests.post(url, headers=headers, json=data, timeout=60)
    print(f"\n状态码：{response.status_code}")
    print(f"返回内容：{response.text[:800]}")
except Exception as e:
    print(f"请求异常：{e}")