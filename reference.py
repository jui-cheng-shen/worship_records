import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import simpledialog, Toplevel, StringVar
import datetime
import sqlite3

# 使用固定名稱的資料庫，不再分年份
db_filename = 'DALIN_TJC_record.db'

# 連接資料庫
conn = sqlite3.connect(db_filename)
database = conn.cursor()

# 創建或更新資料表結構
database.execute('''
    CREATE TABLE IF NOT EXISTS record (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        meeting_type TEXT,
        time_period TEXT,
        week TEXT,
        topic TEXT,
        hymn1 TEXT,
        hymn2 TEXT,
        leader TEXT,
        translator TEXT,
        brother_sister INTEGER,
        male_seeker INTEGER,
        female_seeker INTEGER,
        total INTEGER
    )
''')

# 檢查是否需要從舊表格遷移資料
database.execute("PRAGMA table_info(record)")
columns = [col[1] for col in database.fetchall()]
conn.commit()

# 主視窗設定
win = tk.Tk()
win.geometry("600x650")
win.title("聚會記錄系統")

# 建立框架
main_frame = tk.Frame(win)
main_frame.pack(fill="both", expand=True)

# 添加Tab控制項
notebook = ttk.Notebook(main_frame)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# 建立輸入標籤頁
input_tab = ttk.Frame(notebook)
notebook.add(input_tab, text="資料輸入")

# 建立查詢標籤頁
query_tab = ttk.Frame(notebook)
notebook.add(query_tab, text="資料查詢")

# 在輸入標籤頁中添加群組
input_group = tk.LabelFrame(input_tab, text="基本資料", padx=10, pady=10)
input_group.pack(fill="x", padx=10, pady=5)

# 聚會類別和時段選項
meeting_types = ["安息日", "晚間聚會", "家庭聚會", "唱詩祈禱會", "團契聚會", "特別聚會", "靈恩會", "佈道會"]
time_periods = ["上午", "下午", "晚上"]

# 創建輸入欄位
labels = [
    "日期(YYYY-MM-DD)", "星期", "聚會類別", "時段", "主題", 
    "讚美詩1", "讚美詩2", "主領", "翻譯", 
    "主內信徒人數", "男性慕道者人數", "女性慕道者人數", "總人數"
]

entries = []

# 創建基本資料欄位
for i, label_text in enumerate(labels):
    frame = tk.Frame(input_group)
    frame.pack(fill="x", pady=2)
    
    lbl = tk.Label(frame, text=label_text, width=15, anchor="w")
    lbl.pack(side="left")
    
    if label_text == "日期(YYYY-MM-DD)":
        entry = tk.Entry(frame)
        entry.pack(side="left", fill="x", expand=True)
    elif label_text == "聚會類別":
        entry = ttk.Combobox(frame, values=meeting_types)
        entry.pack(side="left", fill="x", expand=True)
    elif label_text == "時段":
        entry = ttk.Combobox(frame, values=time_periods)
        entry.pack(side="left", fill="x", expand=True)
    elif label_text == "總人數":
        entry = tk.Entry(frame, state="readonly")
        entry.pack(side="left", fill="x", expand=True)
    else:
        entry = tk.Entry(frame)
        entry.pack(side="left", fill="x", expand=True)
    
    entries.append(entry)

# 自動生成星期
def update_weekday(event=None):
    date_str = entries[0].get().strip()
    if not date_str:
        return
        
    try:
        # 直接解析完整日期格式 YYYY-MM-DD
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
            
        weekday = date_obj.strftime("%A")
        weekday_zh = {
            "Monday": "星期一", 
            "Tuesday": "星期二", 
            "Wednesday": "星期三", 
            "Thursday": "星期四", 
            "Friday": "星期五", 
            "Saturday": "星期六", 
            "Sunday": "星期日"
        }
        
        entries[1].delete(0, tk.END)
        entries[1].insert(0, weekday_zh.get(weekday, weekday))
    except ValueError:
        pass

# 自動計算總人數
def calculate_total(event=None):
    try:
        believers = int(entries[9].get() or 0)
        male_seekers = int(entries[10].get() or 0)
        female_seekers = int(entries[11].get() or 0)
        total = believers + male_seekers + female_seekers
        
        # 更新總人數欄位
        entries[12].configure(state="normal")
        entries[12].delete(0, tk.END)
        entries[12].insert(0, str(total))
        entries[12].configure(state="readonly")
    except ValueError:
        pass

# 綁定事件
entries[0].bind("<FocusOut>", update_weekday)
entries[9].bind("<KeyRelease>", calculate_total)
entries[10].bind("<KeyRelease>", calculate_total)
entries[11].bind("<KeyRelease>", calculate_total)

# 初始化星期欄位
update_weekday()

def add_record():
    # 獲取所有欄位資料
    data = [entry.get() if isinstance(entry, tk.Entry) or isinstance(entry, ttk.Combobox) else entry.get() for entry in entries]
    
    # 檢查必填欄位
    required_fields = [0, 2, 3, 4, 9, 10, 11]  # 翻譯欄位可以為空
    if any(not data[i].strip() for i in required_fields):
        messagebox.showerror("錯誤", "請填寫所有必要欄位")
        return
        
    # 處理日期
    date_str = data[0].strip()

    try:
        # 驗證日期格式
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        data[0] = date_obj.strftime("%Y-%m-%d")
    except ValueError:
        messagebox.showerror("日期格式錯誤", "請使用 YYYY-MM-DD 格式")
        return

    # 檢查人數欄位
    try:
        believers = int(data[9])
        male_seekers = int(data[10])
        female_seekers = int(data[11])
        total = believers + male_seekers + female_seekers
        data[12] = str(total)
    except ValueError:
        messagebox.showerror("數據錯誤", "人數欄位必須為數字")
        return

    # 儲存資料
    store_record(data)
    messagebox.showinfo("成功", "資料已成功儲存")

    # 清空輸入欄位，但日期保留為當前日期
    for i, entry in enumerate(entries):
        if i == 0:
            entry.delete(0, tk.END)
            entry.insert(0, datetime.date.today().strftime("%Y-%m-%d"))
            update_weekday()  # 更新星期
        elif i != 12:  # 跳過總人數欄位
            if isinstance(entry, tk.Entry):
                entry.delete(0, tk.END)
            elif isinstance(entry, ttk.Combobox):
                entry.set('')

def store_record(data):
    # 將資料存入資料庫
    database.execute('''
        INSERT INTO record (date, week, meeting_type, time_period, topic, hymn1, hymn2, leader, translator, 
        brother_sister, male_seeker, female_seeker, total)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', data)
    conn.commit()

# 添加儲存按鈕
btn_frame = tk.Frame(input_tab)
btn_frame.pack(pady=10)

save_btn = tk.Button(btn_frame, text="儲存", command=add_record, width=10)
save_btn.pack(side="left", padx=5)

# 查詢功能實作
query_group = tk.LabelFrame(query_tab, text="查詢條件", padx=10, pady=10)
query_group.pack(fill="x", padx=10, pady=5)

# 查詢條件欄位
tk.Label(query_group, text="起始日期(YYYY-MM-DD):").grid(row=0, column=0, padx=5, pady=5)
start_date = tk.Entry(query_group)
start_date.grid(row=0, column=1, padx=5, pady=5)

tk.Label(query_group, text="結束日期(YYYY-MM-DD):").grid(row=0, column=2, padx=5, pady=5)
end_date = tk.Entry(query_group)
end_date.grid(row=0, column=3, padx=5, pady=5)

tk.Label(query_group, text="聚會類別:").grid(row=1, column=0, padx=5, pady=5)
query_meeting_type = ttk.Combobox(query_group, values=[""] + meeting_types)
query_meeting_type.grid(row=1, column=1, padx=5, pady=5)

tk.Label(query_group, text="時段:").grid(row=1, column=2, padx=5, pady=5)
query_time_period = ttk.Combobox(query_group, values=[""] + time_periods)
query_time_period.grid(row=1, column=3, padx=5, pady=5)

# 查詢結果顯示
result_group = tk.LabelFrame(query_tab, text="查詢結果", padx=10, pady=10)
result_group.pack(fill="both", expand=True, padx=10, pady=5)

# 創建 Treeview 用於顯示結果
columns = ("日期", "星期", "聚會類別", "時段", "主題", "主內信徒", "男性慕道者", "女性慕道者", "總人數")
result_tree = ttk.Treeview(result_group, columns=columns, show="headings", height=10)

# 設定欄位標題
for col in columns:
    result_tree.heading(col, text=col)
    result_tree.column(col, width=80)

result_tree.pack(fill="both", expand=True)

# 統計框架
stats_frame = tk.Frame(query_tab)
stats_frame.pack(fill="x", padx=10, pady=5)

stats_label = tk.Label(stats_frame, text="統計結果:")
stats_label.pack(anchor="w")

def search_records():
    # 清空原有結果
    result_tree.delete(*result_tree.get_children())
    
    # 準備查詢條件
    conditions = []
    params = []
    
    # 日期條件
    if start_date.get().strip():
        try:
            start = datetime.datetime.strptime(start_date.get().strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
            conditions.append("date >= ?")
            params.append(start)
        except ValueError:
            messagebox.showerror("錯誤", "起始日期格式不正確，請使用 YYYY-MM-DD")
            return
    
    if end_date.get().strip():
        try:
            end = datetime.datetime.strptime(end_date.get().strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
            conditions.append("date <= ?")
            params.append(end)
        except ValueError:
            messagebox.showerror("錯誤", "結束日期格式不正確，請使用 YYYY-MM-DD")
            return
    
    # 聚會類別條件
    if query_meeting_type.get():
        conditions.append("meeting_type = ?")
        params.append(query_meeting_type.get())
    
    # 時段條件
    if query_time_period.get():
        conditions.append("time_period = ?")
        params.append(query_time_period.get())
    
    # 建構SQL查詢
    sql = "SELECT date, week, meeting_type, time_period, topic, brother_sister, male_seeker, female_seeker, total FROM record"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY date"
    
    try:
        database.execute(sql, params)
        results = database.fetchall()
        
        if not results:
            messagebox.showinfo("查詢結果", "沒有符合條件的記錄")
            return
        
        # 顯示結果
        for row in results:
            result_tree.insert("", "end", values=row)
        
        # 計算統計數據
        calculate_statistics(results)
        
    except sqlite3.Error as e:
        messagebox.showerror("查詢錯誤", str(e))

def calculate_statistics(results):
    # 按聚會類別統計
    stats = {}
    for row in results:
        meeting_type = row[2]  # 聚會類別在索引2
        if meeting_type not in stats:
            stats[meeting_type] = {
                "count": 0,
                "believers": 0,
                "male_seekers": 0,
                "female_seekers": 0,
                "total": 0
            }
        
        stats[meeting_type]["count"] += 1
        stats[meeting_type]["believers"] += row[5]
        stats[meeting_type]["male_seekers"] += row[6]
        stats[meeting_type]["female_seekers"] += row[7]
        stats[meeting_type]["total"] += row[8]
    
    # 顯示統計結果
    stats_text = "統計結果:\n\n"
    for meeting_type, data in stats.items():
        avg_believers = data["believers"] / data["count"] if data["count"] > 0 else 0
        avg_male = data["male_seekers"] / data["count"] if data["count"] > 0 else 0
        avg_female = data["female_seekers"] / data["count"] if data["count"] > 0 else 0
        avg_total = data["total"] / data["count"] if data["count"] > 0 else 0
        
        stats_text += f"{meeting_type} (共{data['count']}次):\n"
        stats_text += f"  平均主內信徒: {avg_believers:.1f}人\n"
        stats_text += f"  平均男性慕道者: {avg_male:.1f}人\n"
        stats_text += f"  平均女性慕道者: {avg_female:.1f}人\n"
        stats_text += f"  平均總人數: {avg_total:.1f}人\n\n"
    
    # 顯示統計視窗
    stats_window = Toplevel(win)
    stats_window.title("統計結果")
    stats_window.geometry("400x300")
    
    stats_text_widget = tk.Text(stats_window)
    stats_text_widget.pack(fill="both", expand=True, padx=10, pady=10)
    stats_text_widget.insert(tk.END, stats_text)
    stats_text_widget.configure(state="disabled")

# 修改記錄功能
def edit_record():
    selected_item = result_tree.focus()
    if not selected_item:
        messagebox.showinfo("提示", "請先選擇一筆記錄")
        return
    
    # 獲取選中記錄的數據
    values = result_tree.item(selected_item, "values")
    
    # 查詢完整記錄
    date_val = values[0]
    meeting_type_val = values[2]
    time_period_val = values[3]
    
    database.execute('''
        SELECT * FROM record 
        WHERE date = ? AND meeting_type = ? AND time_period = ?
    ''', (date_val, meeting_type_val, time_period_val))
    
    record = database.fetchone()
    if not record:
        messagebox.showerror("錯誤", "找不到該記錄")
        return
    
    # 創建修改視窗
    edit_window = Toplevel(win)
    edit_window.title("修改記錄")
    edit_window.geometry("500x500")
    
    edit_frame = tk.LabelFrame(edit_window, text="修改資料", padx=10, pady=10)
    edit_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    edit_entries = []
    
    # 創建修改欄位
    for i, label_text in enumerate(labels):
        frame = tk.Frame(edit_frame)
        frame.pack(fill="x", pady=2)
        
        lbl = tk.Label(frame, text=label_text, width=15, anchor="w")
        lbl.pack(side="left")
        
        if label_text == "聚會類別":
            entry = ttk.Combobox(frame, values=meeting_types)
            entry.set(record[2])  # meeting_type
        elif label_text == "時段":
            entry = ttk.Combobox(frame, values=time_periods)
            entry.set(record[3])  # time_period
        elif label_text == "總人數":
            entry = tk.Entry(frame, state="readonly")
            entry.insert(0, str(record[13]))  # total
        else:
            entry = tk.Entry(frame)
            # 對應資料庫欄位到UI欄位
            if i == 0:  # 日期
                entry.insert(0, record[1])
            elif i == 1:  # 星期
                entry.insert(0, record[4])
            elif i == 4:  # 主題
                entry.insert(0, record[5])
            elif i == 5:  # 讚美詩1
                entry.insert(0, record[6])
            elif i == 6:  # 讚美詩2
                entry.insert(0, record[7])
            elif i == 7:  # 主領
                entry.insert(0, record[8])
            elif i == 8:  # 翻譯
                entry.insert(0, record[9])
            elif i == 9:  # 主內信徒
                entry.insert(0, str(record[10]))
            elif i == 10:  # 男性慕道者
                entry.insert(0, str(record[11]))
            elif i == 11:  # 女性慕道者
                entry.insert(0, str(record[12]))
        
        entry.pack(side="left", fill="x", expand=True)
        edit_entries.append(entry)
    
    # 綁定修改視窗中的事件處理
    def edit_calculate_total(event=None):
        try:
            believers = int(edit_entries[9].get() or 0)
            male_seekers = int(edit_entries[10].get() or 0)
            female_seekers = int(edit_entries[11].get() or 0)
            total = believers + male_seekers + female_seekers
            
            edit_entries[12].configure(state="normal")
            edit_entries[12].delete(0, tk.END)
            edit_entries[12].insert(0, str(total))
            edit_entries[12].configure(state="readonly")
        except ValueError:
            pass
    
    edit_entries[9].bind("<KeyRelease>", edit_calculate_total)
    edit_entries[10].bind("<KeyRelease>", edit_calculate_total)
    edit_entries[11].bind("<KeyRelease>", edit_calculate_total)
    
    # 保存修改
    def save_changes():
        # 獲取修改後的數據
        updated_data = [entry.get() if isinstance(entry, tk.Entry) or isinstance(entry, ttk.Combobox) else entry.get() for entry in edit_entries]
        
        # 驗證日期格式
        try:
            date_obj = datetime.datetime.strptime(updated_data[0], "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("錯誤", "日期格式不正確，請使用 YYYY-MM-DD")
            return
            
        try:
            # 更新數據庫
            database.execute('''
                UPDATE record 
                SET date=?, week=?, meeting_type=?, time_period=?, topic=?, 
                    hymn1=?, hymn2=?, leader=?, translator=?,
                    brother_sister=?, male_seeker=?, female_seeker=?, total=?
                WHERE id=?
            ''', (
                updated_data[0], updated_data[1], updated_data[2], updated_data[3], 
                updated_data[4], updated_data[5], updated_data[6], updated_data[7], 
                updated_data[8], updated_data[9], updated_data[10], updated_data[11],
                updated_data[12], record[0]
            ))
            conn.commit()
            
            messagebox.showinfo("成功", "資料已更新")
            edit_window.destroy()
            
            # 重新查詢以更新顯示
            search_records()
            
        except Exception as e:
            messagebox.showerror("錯誤", f"更新失敗: {str(e)}")
    
    # 添加保存按鈕
    save_btn = tk.Button(edit_window, text="保存修改", command=save_changes)
    save_btn.pack(pady=10)

# 添加查詢和修改按鈕
search_btn = tk.Button(query_group, text="查詢", command=search_records)
search_btn.grid(row=2, column=1, pady=10)

edit_btn = tk.Button(query_group, text="修改選中記錄", command=edit_record)
edit_btn.grid(row=2, column=2, pady=10)

win.mainloop()