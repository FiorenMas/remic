import requests
import json
import os
import random
from datetime import datetime, timedelta, timezone
import feedparser

TENANT_ID = os.getenv('Directory_ID')
CLIENT_ID = os.getenv('Application_ID')
CLIENT_SECRET = os.getenv('Client_Secret')
IMAGE_FOLDER = './Images'

# ===== AUTHENTICATION =====
def get_token():
    url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
    data = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': 'https://graph.microsoft.com/.default'
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()['access_token']

# ===== GET USERS =====
def get_users(token):
    headers = {'Authorization': f'Bearer {token}'}
    r = requests.get('https://graph.microsoft.com/v1.0/users', headers=headers)
    r.raise_for_status()
    users = r.json().get('value', [])
    print("✅ Danh sách user:")
    for u in users:
        print(" 👤", u['userPrincipalName'])
    return users

# ===== GET CALENDAR EVENTS =====
def get_calendar(token, user_id, email):
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/calendar/events'
    headers = {'Authorization': f'Bearer {token}'}
    r = requests.get(url, headers=headers)
    print(f"📅 Lịch của {email} – Status: {r.status_code}")
    if r.status_code == 200:
        print(f"📆 Số sự kiện: {len(r.json().get('value', []))}")
    else:
        print(r.text)

# ===== CREATE DAILY EVENT =====
def create_daily_event(token, user_id):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/events"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    now = datetime.now(timezone.utc)
    start_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
    end_time = start_time + timedelta(minutes=30)
    payload = {
        "subject": "📌 Daily Auto Event",
        "body": {
            "contentType": "HTML",
            "content": "Tự động tạo để duy trì hoạt động lịch mỗi ngày."
        },
        "start": {
            "dateTime": start_time.isoformat(),
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": end_time.isoformat(),
            "timeZone": "UTC"
        }
    }
    r = requests.post(url, headers=headers, json=payload)
    print(f"📆 Tạo sự kiện lịch – Status: {r.status_code}")
    if r.status_code not in [200, 201]:
        print(r.text)

# ===== RSS NEWS & COPILOT =====
def get_news_rss():
    feed = feedparser.parse("https://vnexpress.net/rss/tin-moi-nhat.rss")
    return "\n".join(f"- {entry.title}" for entry in feed.entries[:5])

def generate_copilot_mock():
    samples = [
        "🧠 Hôm nay thời tiết tại Hà Nội nắng ráo, nhiệt độ cao nhất 35°C.",
        "📈 VN-Index tăng nhẹ, nhà đầu tư nước ngoài mua ròng mạnh.",
        "💡 Mẹo: Dùng Ctrl+Shift+V để dán văn bản không định dạng.",
        "🧘 Copilot gợi ý: Thử bài thở 4-7-8 để giảm căng thẳng.",
        "📅 Hôm nay có cuộc họp vào lúc 10h, đừng quên chuẩn bị.",
        "🌐 Trợ lý Copilot sẵn sàng giúp bạn viết email hoặc tạo bảng."
    ]
    return random.choice(samples)

# ===== FOLDER MANAGEMENT =====
def ensure_e5_subfolder(token, user_id, subfolder):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/E5/{subfolder}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    r = requests.get(url, headers=headers)
    if r.status_code == 404:
        create_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root/children"
        folder_structure = {
            "name": "E5",
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace"
        }
        requests.post(create_url, headers=headers, json=folder_structure)

        # Then create subfolder
        create_sub_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/E5/{subfolder}:/children"
        subfolder_data = {
            "name": subfolder,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        requests.post(create_url, headers=headers, json=subfolder_data)

# ===== EMAIL SENDER + FILE UPLOAD =====
def send_personalized_mails(token, sender_email, recipient_list, user_id):
    subject = "📌 MS365 – Bản tin & phản hồi Copilot"
    url = f'https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    ensure_e5_subfolder(token, user_id, "Copilot")

    for recipient in recipient_list:
        today = datetime.now(timezone.utc).strftime('%Y-%m-%d')
        news = get_news_rss()
        copilot_msg = generate_copilot_mock()
        body = f"""📢 Bản tin cá nhân hóa ngày {today}:
{news}

🤖 Copilot nói:
{copilot_msg}

✅ Email tự động để duy trì hoạt động tài khoản."""

        # Send email
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": recipient}}]
            }
        }
        print(f"📧 {sender_email} gửi đến: {recipient}")
        r = requests.post(url, headers=headers, json=payload)
        print(f"📨 Trạng thái: {r.status_code}")
        if r.status_code != 202:
            print(r.text)

        # Upload to OneDrive under /E5/Copilot/
        filename = f"copilot_{recipient.replace('@','_')}.txt"
        with open(filename, "w", encoding="utf-8") as f:
            f.write(body)
        upload_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/E5/Copilot/{filename}:/content"
        with open(filename, "rb") as f:
            upload_res = requests.put(upload_url, headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "text/plain"
            }, data=f)
        print(f"☁️ Upload Copilot – Status: {upload_res.status_code}")

# ===== IMAGE UPLOADER =====
def upload_random_images(token, user_id):
    if not os.path.exists(IMAGE_FOLDER):
        print("❌ Thư mục ảnh không tồn tại:", IMAGE_FOLDER)
        return
    files = [f for f in os.listdir(IMAGE_FOLDER) if f.lower().endswith(('.jpg', '.png'))]
    if not files:
        print("❌ Không tìm thấy ảnh.")
        return
    ensure_e5_subfolder(token, user_id, "Images")
    selected = random.sample(files, min(3, len(files)))
    for filename in selected:
        path = os.path.join(IMAGE_FOLDER, filename)
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/E5/Images/{filename}:/content"
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/octet-stream'
        }
        with open(path, 'rb') as f:
            r = requests.put(url, headers=headers, data=f)
        print(f"🖼️ Upload {filename} – Status: {r.status_code}")
        if r.status_code not in [200, 201]:
            print(r.text)

# ===== ONEDRIVE CHECK =====
def check_onedrive_ready(token, user_id):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive"
    headers = { "Authorization": f"Bearer {token}" }
    r = requests.get(url, headers=headers)
    return r.status_code == 200

# ===== MAIN =====
if __name__ == '__main__':
    try:
        token = get_token()
        users = get_users(token)
        for sender in users:
            sender_email = sender['userPrincipalName']
            sender_id = sender['id']
            print(f"\n🔄 Đang xử lý: {sender_email}")
            get_calendar(token, sender_id, sender_email)
            create_daily_event(token, sender_id)

            others = [u['userPrincipalName'] for u in users if u['userPrincipalName'] != sender_email]
            if others:
                recipients = random.sample(others, min(random.randint(2, 5), len(others)))
                send_personalized_mails(token, sender_email, recipients, sender_id)

            if check_onedrive_ready(token, sender_id):
                upload_random_images(token, sender_id)
            else:
                print("⚠️ OneDrive chưa sẵn sàng.")

    except Exception as e:
        print("❌ Lỗi toàn cục:", e)
