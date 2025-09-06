import requests
import json

url = 'http://localhost:8000/log'
headers = {'Content-Type': 'application/json'}

# 发送三条不同级别的日志
test_logs = [
    {"level": "INFO", "message": "这是一条信息日志", "source": "test_script", "timestamp": "2025-09-05T15:30:00Z"},
    {"level": "WARNING", "message": "这是一条警告日志", "source": "test_script", "timestamp": "2025-09-05T15:31:00Z"},
    {"level": "ERROR", "message": "这是一条错误日志", "source": "test_script", "timestamp": "2025-09-05T15:32:00Z"}
]

for log in test_logs:
    response = requests.post(url, headers=headers, data=json.dumps(log))
    print(f"发送日志: {log['level']} - {response.status_code} - {response.text}")