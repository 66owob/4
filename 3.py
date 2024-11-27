import re
import sqlite3
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import requests

DATABASE_NAME = "contacts.db"

# 全局變數
tree = None
url_entry = None

def create_database():
    """建立 SQLite 資料庫和資料表。"""
    with sqlite3.connect(DATABASE_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS contacts (
            iid INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            title TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE
        )
        """)
        conn.commit()

def save_to_database(contacts: list[dict]):
    """將聯絡資訊保存到資料庫，避免重複資料。
    
    Args:
        contacts (list[dict]): 聯絡資訊列表。
    """
    with sqlite3.connect(DATABASE_NAME) as conn:
        cursor = conn.cursor()
        for contact in contacts:
            try:
                cursor.execute("""
                INSERT OR IGNORE INTO contacts (name, title, email)
                VALUES (:name, :title, :email)
                """, contact)
            except sqlite3.Error as e:
                print(f"資料庫錯誤: {e}")
        conn.commit()

def fetch_emails_from_ncut(url: str) -> list[str]:
    """從指定的 URL 擷取 @ncut.edu.tw 的電子郵件。
    
    Args:
        url (str): 目標 URL。
    
    Returns:
        list: 擷取到的電子郵件列表。
    """
    try:
        # 發送 HTTP GET 請求
        response = requests.get(url)
        response.raise_for_status()  # 檢查請求是否成功
        content = response.text

        # 使用正則表達式匹配 @ncut.edu.tw 的電子郵件
        email_pattern = r'[a-zA-Z0-9._%+-]+@ncut\.edu\.tw'
        emails = re.findall(email_pattern, content)

        # 去重與排序
        unique_emails = sorted(set(emails))
        return unique_emails
    except requests.exceptions.RequestException as e:
        print(f"Error fetching the URL: {e}")
        return []

def extract_contacts(html: str, emails: list[str]) -> list[dict]:
    """從 HTML 中提取聯絡資訊（姓名、職稱、Email），並將 @ncut.edu.tw 電子郵件添加到資料中。
    
    Args:
        html (str): HTML 內容。
        emails (list): 擷取到的電子郵件列表。
    
    Returns:
        list[dict]: 聯絡資訊列表。
    """
    contacts = []
    email_index = 0  # 用來分配電子郵件的索引

    # 用正則表達式提取每位教師的區塊
    teacher_blocks = re.findall(r'<div class="teacher-list">.*?</div>\s*</div>', html, re.DOTALL)

    for block in teacher_blocks:
        # 提取姓名
        name_match = re.search(r'<div class="member_name">.*?<a[^>]*>(.*?)</a>', block, re.DOTALL)
        name = name_match.group(1).strip() if name_match else "未知"

        # 提取職稱
        title_match = re.search(r'<div class="member_info_title">.*?職稱.*?</div>\s*<div class="member_info_content">(.*?)</div>', block, re.DOTALL)
        title = title_match.group(1).strip() if title_match else "未知"

        # 使用擷取到的 @ncut.edu.tw 電子郵件
        email = emails[email_index] if email_index < len(emails) else "未知"
        email_index += 1

        # 添加到聯絡人列表
        contacts.append({
            "name": name,
            "title": title,
            "email": email,
        })

    return contacts

def display_contacts(tree: ttk.Treeview, contacts: list[dict]):
    """在圖形界面中顯示聯絡資訊。
    
    Args:
        tree (ttk.Treeview): Tkinter Treeview 元件。
        contacts (list[dict]): 聯絡資訊列表。
    """
    for item in tree.get_children():
        tree.delete(item)
    for contact in contacts:
        tree.insert("", "end", values=(contact["name"], contact["title"], contact["email"]))

def fetch_and_display():
    """從 URL 擷取聯絡人資訊並顯示於界面。"""
    url = url_entry.get()
    if not url:
        messagebox.showerror("錯誤", "請輸入有效的 URL！")
        return
    html = fetch_html(url)
    if html:
        # 擷取 @ncut.edu.tw 的電子郵件
        emails = fetch_emails_from_ncut(url)
        contacts = extract_contacts(html, emails)
        save_to_database(contacts)
        display_contacts(tree, contacts)

def fetch_html(url: str) -> str:
    """從指定的 URL 獲取 HTML 內容。
    
    Args:
        url (str): 目標 URL。
    
    Returns:
        str: HTML 內容。
    """
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.text
    except requests.RequestException as e:
        messagebox.showerror("錯誤", f"無法獲取網頁內容: {e}")
        return ""

def main():
    """主程式入口。"""
    create_database()

    global tree, url_entry  # 宣告為全局變數

    # 建立 Tkinter 主窗口
    root = Tk()
    root.title("聯絡資訊爬蟲")
    root.geometry("800x600")

    # URL 輸入框
    frame = Frame(root)
    frame.pack(fill=X, padx=10, pady=10)
    Label(frame, text="目標 URL:").pack(side=LEFT)
    url_entry = Entry(frame, width=50)
    url_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
    Button(frame, text="抓取", command=fetch_and_display).pack(side=LEFT, padx=5)

    # Treeview 表格顯示聯絡資訊
    columns = ("name", "title", "email")
    tree = ttk.Treeview(root, columns=columns, show="headings")
    tree.heading("name", text="姓名")
    tree.heading("title", text="職稱")
    tree.heading("email", text="信箱")
    tree.column("name", width=150, anchor=W)
    tree.column("title", width=150, anchor=W)
    tree.column("email", width=300, anchor=W)
    tree.pack(fill=BOTH, expand=True, padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
