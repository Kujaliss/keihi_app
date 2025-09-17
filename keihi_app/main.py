#経費を各PJごとに計上するアプリ

import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkFont
import pandas as pd
import os
from datetime import datetime
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  #DPIスケーリング有効
except:
    pass

#このスクリプトのあるディレクトリを取得
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, 'data', 'keihi_log.csv')
EXPORT_FOLDER = os.path.join(BASE_DIR, 'exports')


#初期セットアップ
os.makedirs(os.path.join(BASE_DIR, 'data'), exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)


#CSVがなければ作成
if not os.path.exists(DATA_FILE):
    initial_df = pd.DataFrame(columns=["プロジェクト", "使用日", "使用者", "内容", "種別", "金額"])
    initial_df.to_csv(DATA_FILE, index=False)


#データ保存関数
def save_data(pj, date, user, item, category, amount):
    df = pd.read_csv(DATA_FILE)
    new_row = {
        "プロジェクト": pj,
        "使用日": date,
        "使用者": user,
        "内容": item,
        "種別": category,
        "金額": amount
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")


#GUI構築
root = tk.Tk()
root.title("経費計上アプリ")


#フォント指定
default_font = tkFont.Font(family="Yu Gothic", size=10)
root.option_add("*Font", default_font)


#入力項目
tk.Label(root, text="プロジェクト").grid(row=0, column=0)
pj_var = tk.StringVar()
pj_box = ttk.Combobox(root, textvariable=pj_var)
pj_box['values'] = ["PJ-A", "PJ-B", "PJ-C"]
pj_box.grid(row=0, column=1)

tk.Label(root, text="使用日").grid(row=1, column=0)
date_var = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
tk.Entry(root, textvariable=date_var).grid(row=1, column=1)

tk.Label(root, text="使用者").grid(row=2, column=0)
user_var = tk.StringVar()
tk.Entry(root, textvariable=user_var).grid(row=2, column=1)

tk.Label(root, text="内容").grid(row=3, column=0)
item_var = tk.StringVar()
tk.Entry(root, textvariable=item_var).grid(row=3, column=1)

tk.Label(root, text="種別").grid(row=4, column=0)
category_var = tk.StringVar()
category_box = ttk.Combobox(root, textvariable=category_var)
category_box['values'] = ["交通費", "旅費", "交際費", "会議費", "その他"]
category_box.grid(row=4, column=1)

tk.Label(root, text="金額").grid(row=5, column=0)
amount_var = tk.StringVar()
tk.Entry(root, textvariable=amount_var).grid(row=5, column=1)


#Treeview枠作成
tree_frame = tk.Frame(root)
tree_frame.grid(row=9, column=0, columnspan=2, pady=10, sticky="nsew")

# 縦スクロールバー
tree_scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

# 横スクロールバー
tree_scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

# Treeview
tree = ttk.Treeview(
    tree_frame,
    columns=("プロジェクト", "使用日", "使用者", "内容", "種別", "金額"),
    show='headings',
    yscrollcommand=tree_scroll_y.set,
    xscrollcommand=tree_scroll_x.set,
    height=20  # 表示行数
)

# 行高さを調整
style = ttk.Style()
style.configure("Treeview", rowheight=30)

# 列幅設定
col_widths = {
    "プロジェクト": 100,
    "使用日": 100,
    "使用者": 100,
    "内容": 250,
    "種別": 100,
    "金額": 100
}

for col in ("プロジェクト", "使用日", "使用者", "内容", "種別", "金額"):
    tree.heading(col, text=col)
    tree.column(col, width=col_widths[col], anchor="center")

tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# スクロールバーと連動
tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)

# 親ウィンドウのリサイズ対応
root.grid_rowconfigure(9, weight=1)
root.grid_columnconfigure(0, weight=1)


#Treeviewにデータを読み込む関数
def load_data_to_tree():
    #一旦すべて削除
    for row in tree.get_children():
        tree.delete(row)
    #CSV読み込み
    df = pd.read_csv(DATA_FILE)
    for _, row in df.iterrows():
        tree.insert("", tk.END, values=(row["プロジェクト"], row["使用日"], row["使用者"], row["内容"], row["種別"], row["金額"]))


#登録ボタン
def on_register():
    try:
        amount = int(amount_var.get())
    except ValueError:
        messagebox.showerror("エラー", "金額は数字で入力してください。")
        return
    
    save_data(
        pj_var.get(), date_var.get(),user_var.get(),
        item_var.get(), category_var.get(), amount
    )
    messagebox.showinfo("完了", "登録しました！")
    amount_var.set("")      #金額をリセット
    item_var.set("")      #内容をリセット
    load_data_to_tree()   #登録後に画面更新


#月別出力関数
def export_monthly():
    import tkinter.simpledialog as simpledialog

    #年月入力ダイアログ
    year = simpledialog.askstring("年入力", "年を入力してください。（例：2025）")
    month = simpledialog.askstring("月入力", "月を入力してください。（例：09）")

    if not (year and month):
        return
    
    df = pd.read_csv(DATA_FILE)
    
    df['使用日'] = pd.to_datetime(df['使用日'])
    df_filtered = df[(df['使用日'].dt.year == int(year)) & (df['使用日'].dt.month == int(month))]

    if df_filtered.empty:
        messagebox.showinfo("出力結果", "指定月のデータがありません。")
        return
    
    #出力ファイル名
    output_file = os.path.join(EXPORT_FOLDER, f"{year}-{month}_経費一覧.csv")
    df_filtered.to_csv(output_file, index=False, encoding="utf-8-sig")

    messagebox.showinfo("出力完了", f"{output_file}に出力しました！")


#PJ別出力関数
def export_by_pj():
    import tkinter.simpledialog as simpledialog

    #PJ別選択ダイアログ
    pj_name = simpledialog.askstring("プロジェクト入力", "出力したいプロジェクト名を入力してください。（例：PJ-A）")

    if not pj_name:
        return
    
    df = pd.read_csv(DATA_FILE)

    #PJフィルター
    df_filtered = df[df['プロジェクト'] == pj_name]

    if df_filtered.empty:
        messagebox.showinfo("出力結果", f"{pj_name}のデータはありません。")
        return
    
    #出力ファイル名
    output_file = os.path.join(EXPORT_FOLDER, f"{pj_name}_経費一覧.csv")
    df_filtered.to_csv(output_file, index=False, encoding="utf-8-sig")

    messagebox.showinfo("出力完了", f"{output_file}に出力しました！")


#データリセット関数
def reset_all_data():
    if messagebox.askyesno("確認", "本当に全データを削除しますか？"):
        #csv初期化
        initial_df = pd.DataFrame(columns=["プロジェクト", "使用日", "使用者", "内容", "種別", "金額"])
        initial_df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

        #Treeviewクリア
        for row in tree.get_children():
            tree.delete(row)

        messagebox.showinfo("完了", "全データをリセットしました。")


#ボタンの追加
tk.Button(root, text="登録", command=on_register).grid(row=6, column=0, columnspan=2, pady=10)
tk.Button(root, text="月別CSV出力", command=export_monthly).grid(row=7, column=0, columnspan=2, pady=5)
tk.Button(root, text="PJ別CSV出力", command=export_by_pj).grid(row=8, column=0, columnspan=2, pady=5)
tk.Button(root, text="全データリセット", command=reset_all_data, fg="red").grid(row=10, column=0, columnspan=2, pady=5)


#アプリ起動時に読み込み
load_data_to_tree()
root.mainloop()