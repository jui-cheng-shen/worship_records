import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os
import datetime

year = str(datetime.datetime.now().year)

filename = year +' record.xlsx'

if not os.path.exists(filename):
    wb = Workbook()
    ws = wb.active
    ws.title ="聚會紀錄"
    headers = [
        "日期", "星期" , "主題" , "讚美詩1","讚美詩2" ,  "弟兄人數", "姊妹人數", "慕道者人數", "總人數"
    ]
    ws.append(headers)
    wb.save(filename)


record = []

win =tk.Tk()
win.geometry("500x500")
win.title("record")

group = tk.LabelFrame(win, text="請輸入資料", padx=10, pady=10)
group.pack(padx=10, pady=10)

labels = [
    "日期", "星期" , "主題" , "讚美詩","讚美詩" ,  "弟兄人數", "姊妹人數", "慕道者人數", "總人數"
]


record_input = tk.Entry()
entries = []

for label_text in labels:
    frame = tk.Frame(group)
    frame.pack(fill="x", pady=2)
    lbl = tk.Label(frame, text=label_text, width=15, anchor="w")
    lbl.pack(side="left")
    entry = tk.Entry(frame)
    entry.pack(side="left", fill="x", expand=True)
    entries.append(entry)

def add_record():
    data = [entry.get() for entry in entries]

    if any(value.strip() == "" for value in data):
        messagebox.showerror("錯誤", "請填寫所有欄位")
        return
    date_str = data[0].strip()

    try:
        # 嘗試解析日期
        if "/" in date_str:
            date_obj = datetime.datetime.strptime(date_str, "%Y/%m/%d")
        elif "-" in date_str:
            date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        else:
            messagebox.showerror("錯誤", "日期格式不正確，請使用 YYYY/MM/DD 或 YYYY-MM-DD 格式")
            return
    except ValueError:
        messagebox.showerror("error")
        return

    try :
        brother = int(data[5])
        sister = int(data[6])
        seeker = int(data[7])
        total = brother + sister + seeker
        data[8] = str(total)
    except ValueError:
        messagebox.showerror("error")

    record.append(data)
    store_record(data)

    for entry in entries:
        entry.delete(0, tk.END)

def store_record():
    values = [entry.get() for entry in entries]
    wb = load_workbook(filename)
    ws = wb.active
    ws.append(values)
    wb.save(filename)

btn = tk.Button(win, text = "儲存", command=add_record)
btn.pack(pady=10)

win.mainloop()
