import tkinter as tk
from tkinter import ttk , Toplevel
from tkinter import messagebox
import datetime
import sqlite3


db_filename = "DALIN_TJC_record.db"
conn = sqlite3.connect(db_filename)
database = conn.cursor()

database.execute('''CREATE TABLE IF NOT EXISTS count (
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
        brother INTEGER,
        sister INTEGER,
        male_seeker INTEGER,
        female_seeker INTEGER,
        total_believer INTEGER,
        total_seeker INTEGER,
        total INTEGER
    )
''')

database.execute("PRAGMA table_info(count)")
columns = [col[1] for col in database.fetchall()]
conn.commit()

win = tk.Tk()
win.title("DALIN TJC RECORD")
win.geometry("1200x800")

main_frame = ttk.Frame(win)
main_frame.pack(fill=tk.BOTH, expand=True)
notebook = ttk.Notebook(main_frame)
notebook.pack(fill="both", expand=True)

input_tab = ttk.Notebook(main_frame)
notebook.add(input_tab, text="聚會記錄輸入")

query_tab = ttk.Frame(main_frame)
notebook.add(query_tab, text="聚會記錄查詢")

revise_tab = ttk.Frame(main_frame)
notebook.add(revise_tab, text="聚會記錄修改")

input_group = tk.LabelFrame(input_tab, text="輸入聚會記錄", padx=10, pady=10)
input_group.pack(fill="x", padx=10, pady=10)

reunion_type = ["家庭聚會","唱詩祈禱會","晚間聚會","安息日聚會","團契聚會","特別聚會","佈道會","靈恩會"]
reunion_time = ["上午","下午","晚上"]

labels = [
    "日期(西元年/月/日)", "星期" , "聚會類別" ,"時間段" , "主題" , "讚美詩1" , "讚美詩2"
    ,"主領", "翻譯", "弟兄人數", "姊妹人數" ,"慕道者(男)" , "慕道者(女)", "主內總人數" ,"慕道者總人數" , "總人數"
]

entries = []

for i , label_text in enumerate(labels):
    frame = tk.Frame(input_group)
    frame.pack(fill="x", pady = 5)

    lbl = tk.Label(frame, text=label_text,width=15, anchor="w")
    lbl.pack(side="left")

    if label_text == "日期(西元年/月/日)":
        entry = tk.Entry(frame)
        entry.pack(side="left",fill ="x", expand=True)
    elif label_text == "聚會類別" :
        entry = ttk.Combobox(frame , values=reunion_type)
        entry.pack(side="left", fill="x", expand=True)
    elif label_text == "時間段":
        entry = ttk.Combobox(frame, values=reunion_time)
        entry.pack(side="left", fill="x", expand=True)
    elif label_text == "主內總人數" or label_text == "慕道者總人數" or label_text == "總人數":
        entry = tk.Entry(frame)
        entry.pack(side="left", fill="x", expand=True)
    else:
        entry = tk.Entry(frame)
        entry.pack(side="left", fill="x", expand=True)

    entries.append(entry)

def generate_week(event = None):
    date_str = entries[0].get().strip()

    weekday_zh ={
            "Monday": "星期一",
            "Tuesday": "星期二",
            "Wednesday": "星期三",
            "Thursday": "星期四",
            "Friday": "星期五",
            "Saturday": "星期六",
            "Sunday": "星期日"
        }

    try :
        data_object = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        weekday = data_object.strftime("%A")

        entries[1].delete(0, tk.END)  # 先清空原有內容
        entries[1].insert(0, weekday_zh[weekday])
    except ValueError:
        try:
            data_object = datetime.datetime.strptime(date_str, "%Y/%m/%d")
            weekday = data_object.strftime("%A")
            entries[1].delete(0, tk.END)
            entries[1].insert(0, weekday_zh[weekday])
        except ValueError:
            pass

def calculate_total(event=None):
    try:
        brother_count = int(entries[9].get() or 0)
        sister_count = int(entries[10].get() or 0)
        total_believer = brother_count + sister_count

        male_seeker = int(entries[11].get() or 0)
        female_seeker = int(entries[12].get() or 0)
        total_seeker = male_seeker + female_seeker

        entries[13].delete(0, tk.END)
        entries[13].insert(0, total_believer)

        entries[14].delete(0, tk.END)
        entries[14].insert(0, total_seeker)

        entries[15].delete(0, tk.END)
        entries[15].insert(0, total_believer + total_seeker)
    except ValueError:
        pass

# 綁定事件,使日期輸入框失去焦點時自動生成星期
entries[0].bind("<FocusOut>", generate_week)
entries[9].bind("<KeyRelease>", calculate_total)
entries[10].bind("<KeyRelease>", calculate_total)
entries[11].bind("<KeyRelease>", calculate_total)
entries[12].bind("<KeyRelease>", calculate_total)

generate_week()

def save_record():
    data = [entry.get() if isinstance(entry, tk.Entry) else entry.get() for entry in entries]

    required_fields = [0,2,3,9,10,11,12,13,14,15]
    if any(not data[i] for i in required_fields):
        messagebox.showerror("錯誤", "請填寫所有必填欄位")
        return

    date_str = data[0].strip()

#檢查日期
    try:
        try:
            data_object = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            data_object = datetime.datetime.strptime(date_str, "%Y/%m/%d")
        data[0] = data_object.strftime("%Y-%m-%d")
    except ValueError:
        messagebox.showerror("錯誤", "日期格式錯誤，請使用西元年/月/日格式")
        return

#檢查人數
    try:
        for i in [9, 10, 11, 12, 13, 14, 15]:
            data[i] = int(data[i]) if data[i].isdigit() else 0
    except ValueError:
        messagebox.showerror("error","人數請用數字填寫")
        return

    store_data(data)

    messagebox.showinfo("成功", "聚會記錄已保存")

    for entry in entries:
        if isinstance(entry, tk.Entry):
            entry.delete(0, tk.END)
        elif isinstance(entry, ttk.Combobox):
            entry.set('')  # 對 Combobox 正確清空方法

def store_data(data):
    database.execute('''
        INSERT INTO count (date, meeting_type, time_period, week, topic, hymn1, hymn2, leader, translator,
        brother, sister, male_seeker, female_seeker, total_believer, total_seeker, total)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', data)
    conn.commit()


btn_frame = tk.Frame(input_tab)
btn_frame.pack(fill="x",pady=10)

save_batton = tk.Button(btn_frame, text="儲存", command=save_record)
save_batton.pack(side="left", padx=5)

query_group= tk.LabelFrame(query_tab, text="查詢聚會記錄", padx=10, pady=10)
query_group.pack(fill="x", padx=10, pady=10)

#查詢
tk.Label(query_group, text="起始日期(西元年/月/日):").grid(row=0, column=0,padx=5, pady=5, sticky="w")
start_date = tk.Entry(query_group)
start_date.grid(row=0, column=1, padx=5, pady=5)

tk.Label(query_group, text="結束日期(西元年/月/日):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
end_date = tk.Entry(query_group)
end_date.grid(row=0, column=3, padx=5, pady=5)

weekday_zh ={
            "Monday": "星期一",
            "Tuesday": "星期二",
            "Wednesday": "星期三",
            "Thursday": "星期四",
            "Friday": "星期五",
            "Saturday": "星期六",
            "Sunday": "星期日"
        }

tk.Label(query_group, text="聚會類別:").grid(row=1, column=0, padx=5, pady=5, sticky="nw")

# 創建框架容納複選框
type_frame = tk.Frame(query_group)
type_frame.grid(row=1, column=1, padx=5, pady=5, sticky="w")

query_type_vars = {}

# 為每個聚會類別創建一個複選框
for i, type_name in enumerate(reunion_type):
    var = tk.BooleanVar()
    query_type_vars[type_name] = var

    row_pos = i // 4
    col_pos = i % 4

    chk = tk.Checkbutton(type_frame, text=type_name, variable=var)
    chk.grid(row=row_pos, column=col_pos, sticky="w", padx=3)

tk.Label(query_group, text="時間段:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
query_reunion_time = ttk.Combobox(query_group, values=reunion_time)
query_reunion_time.grid(row=1, column=3, padx=5, pady=5)

result_group = tk.LabelFrame(query_tab, text="查詢結果", padx=10, pady=10)
result_group.pack(fill="both", padx=10, pady=10, expand=True)

columns = ["日期", "星期", "聚會類別", "時間段",  "主領",
            "弟兄人數", "姊妹人數", "慕道者(男)", "慕道者(女)", "主內總人數", "慕道者總人數", "總人數"]
result_tree= ttk.Treeview(result_group, columns=columns, show="headings", height=10)

for col in columns:
    result_tree.heading(col, text=col)
    result_tree.column(col, width=80)

result_tree.pack(fill="both", expand=True)

stats_frame = tk.Frame(query_tab)
stats_frame.pack(fill="both", expand=True, padx=10, pady=5)

stats_header = tk.Label(stats_frame, text="統計結果", font=("Arial", 10, "bold"))
stats_header.pack(anchor="w", pady=(5,0))

# 添加一個可滾動的文本區域用於顯示統計結果
stats_text_frame = tk.Frame(stats_frame)
stats_text_frame.pack(fill="both", expand=True)

stats_scrollbar = tk.Scrollbar(stats_text_frame)
stats_scrollbar.pack(side="right", fill="y")

stats_text_widget = tk.Text(stats_text_frame, height=8, yscrollcommand=stats_scrollbar.set, 
                            relief="sunken", borderwidth=1)
stats_text_widget.pack(side="left", fill="both", expand=True)
stats_scrollbar.config(command=stats_text_widget.yview)

def search_records():
    result_tree.delete(*result_tree.get_children())

    condition = []
    params = []

    # 處理日期範圍
    if start_date.get().strip():
        try:
            try:
                start = datetime.datetime.strptime(start_date.get().strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:

                start = datetime.datetime.strptime(start_date.get().strip(), "%Y/%m/%d").strftime("%Y-%m-%d")

            condition.append("date >= ?")
            params.append(start)
        except ValueError:
            messagebox.showerror("錯誤", "起始日期格式錯誤")
            return

    if end_date.get().strip():
        try:
            try:
                end = datetime.datetime.strptime(end_date.get().strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                end = datetime.datetime.strptime(end_date.get().strip(), "%Y/%m/%d").strftime("%Y-%m-%d")

            condition.append("date <= ?")
            params.append(end)
        except ValueError:
            messagebox.showerror("錯誤", "結束日期格式錯誤")
            return

    # 處理聚會類別複選框
    selected_types = [type_name for type_name, var in query_type_vars.items() if var.get()]
    if selected_types:
        placeholders = ", ".join(["?" for _ in selected_types])
        condition.append(f"meeting_type IN ({placeholders})")
        params.extend(selected_types)

    # 處理時間段
    if query_reunion_time.get():
        condition.append("time_period = ?")
        params.append(query_reunion_time.get())

    # 建構完整查詢
    sql = "SELECT * FROM count"
    if condition:
        sql += " WHERE " + " AND ".join(condition)
    sql += " ORDER BY date"

    try:
        database.execute(sql, params)
        records = database.fetchall()

        if not records:
            messagebox.showinfo("查詢結果", "沒有符合條件的記錄")
            stats_text_widget.delete(1.0, tk.END)  # 清空統計區域
            return

        for record in records:
            values = (
                record[1],
                record[4],
                record[2],
                record[3],
                record[8],
                record[10],
                record[11],
                record[12],
                record[13],
                record[14],
                record[15],
                record[16]
            )
            result_tree.insert("", "end", values=values)

        # 計算統計數據
        stats_text = calculate_stats(records)

        # 更新統計文本區域
        stats_text_widget.delete(1.0, tk.END)
        stats_text_widget.insert(tk.END, stats_text)

    except Exception as e:
        messagebox.showerror("查詢錯誤", str(e))

def calculate_stats(results):
    stats = {}
    for row in results:
        meeting_type = row[2]
        if meeting_type not in stats:
            stats[meeting_type] = {
                "count": 0,
                "brothers": 0,
                "sisters": 0,
                "total_believers": 0,
                "total_seekers": 0,
                "total": 0
            }
        stats[meeting_type]["count"] += 1
        stats[meeting_type]["brothers"] += row[10]
        stats[meeting_type]["sisters"] += row[11]
        stats[meeting_type]["total_believers"] += row[14]
        stats[meeting_type]["total_seekers"] += row[15]
        stats[meeting_type]["total"] += row[16]

    stats_text = "統計結果:\n\n"
    for meeting_type, data in stats.items():
        avg_brothers = data["brothers"] / data["count"] if data["count"] > 0 else 0
        avg_sisters = data["sisters"] / data["count"] if data["count"] > 0 else 0
        avg_total_believers = data["total_believers"] / data["count"] if data["count"] > 0 else 0
        avg_total_seekers = data["total_seekers"] / data["count"] if data["count"] > 0 else 0
        avg_total = data["total"] / data["count"] if data["count"] > 0 else 0

        stats_text += f"{meeting_type} (共{data['count']}次):\n"
        stats_text += f"平均弟兄人數: {avg_brothers:.2f}\n"
        stats_text += f"平均姊妹人數: {avg_sisters:.2f}\n"
        stats_text += f"平均主內總人數: {avg_total_believers:.2f}\n"
        stats_text += f"平均慕道者總人數: {avg_total_seekers:.2f}\n"
        stats_text += f"平均總人數: {avg_total:.2f}\n\n"
    return stats_text

def show_stats_window(stats_text):
    stats_window = Toplevel(win)
    stats_window.title("統計結果")
    stats_window.geometry("400x300")

    stats_text_widget = tk.Text(stats_window)
    stats_text_widget.pack(fill="both", expand=True, padx=10, pady=10)
    stats_text_widget.insert(tk.END, stats_text)
    stats_text_widget.configure(state="disabled")

search_button = tk.Button(query_group, text="查詢", command=search_records)
search_button.grid(row=2, column=0, columnspan=4, pady=10)


#修改

revise_group = tk.LabelFrame(revise_tab, text="修改聚會記錄", padx=10, pady=10)
revise_group.pack(fill="x", padx=10, pady=10)

tk.Label(revise_group, text="聚會紀錄日期(西元年/月/日):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
revise_date_entry = tk.Entry(revise_group)
revise_date_entry.grid(row=0, column=1, padx=5, pady=5)

revise_entries = []

for i , label_text in enumerate(labels):
    frame = tk.Frame(revise_group)
    frame.grid(row=(i // 2) + 1, column=(i % 2) * 2, columnspan=2, sticky="w", padx=5, pady=3 )

    lbl = tk.Label(frame, text=label_text, width=15, anchor="w")
    lbl.pack(side="left")

    if label_text == "聚會類別":
        entry = ttk.Combobox(frame , values=reunion_type)
    elif label_text == "時間段":
        entry = ttk.Combobox(frame, values=reunion_time)
    else:
        entry = tk.Entry(frame)

    entry.pack(side="left", fill="x", expand=True)
    revise_entries.append(entry)

def load_record():
    date_str = revise_date_entry.get().strip()
    try:
        try:
            parsed_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            parsed_date = datetime.datetime.strptime(date_str, "%Y/%m/%d").strftime("%Y-%m-%d")
    except ValueError:
        messagebox.showerror("錯誤", "日期格式錯誤，請使用 YYYY-MM-DD 或 YYYY/MM/DD")
        return

    database.execute("SELECT * FROM count WHERE date = ?", (parsed_date,))
    record = database.fetchone()

    if not record:
        messagebox.showerror("錯誤", "沒有找到該日期的聚會記錄")
        return

    for i in range(16):
        if i < len(revise_entries):
            revise_entries[i].delete(0, tk.END)
            revise_entries[i].insert(0, record[i + 1])

load_button = tk.Button(revise_group, text="載入記錄", command=load_record)
load_button.grid(row=0, column=2, padx=5, pady=5)

def update_record():
    date_str = revise_date_entry.get().strip()
    try:
        try:
            parsed_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            parsed_date = datetime.datetime.strptime(date_str, "%Y/%m/%d").strftime("%Y-%m-%d")
    except ValueError:
        messagebox.showerror("錯誤", "日期格式錯誤，請使用 YYYY-MM-DD 或 YYYY/MM/DD")
        return

    new_data = [entry.get() for entry in revise_entries]

    for idx in [0,2,3,9,10,11,12,13,14,15]:
        if not new_data[idx]:
            messagebox.showerror("錯誤", f"請填寫必填欄位: {labels[idx]}")
            return
    try:
        for i in [9,10,11,12,13,14,15]:
            new_data[i] = int(new_data[i] if new_data[i].isdigit() else 0)
    except ValueError:
        messagebox.showerror("錯誤", "人數部分輸入數字")
        return

    sql = '''
    UPDATE count SET
    date = ?,  meeting_type=?, time_period=?, week=?, topic=?,
        hymn1=?, hymn2=?, leader=?, translator=?,
        brother=?, sister=?, male_seeker=?, female_seeker=?,
        total_believer=?, total_seeker=?, total=?
        WHERE date=? '''

    database.execute(sql, (*new_data, parsed_date))
    conn.commit()

    messagebox.showinfo("成功", "聚會記錄已更新")

update_button = tk.Button(revise_group, text = "更新" , command=update_record)
update_button.grid(row=0, column=3 , padx=5, pady=5)


tk.Label(revise_group,)
win.mainloop()