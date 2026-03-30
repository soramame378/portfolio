import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

def create_file_list():
    # 1. フォルダ選択ダイアログを表示
    folder_path = filedialog.askdirectory()
    if not folder_path:  # キャンセルされた場合
        return

    try:
        file_data = []
        # 2. フォルダ内の走査
        items = os.listdir(folder_path)
        
        if not items:
            raise ValueError("選択されたフォルダは空です。")

        for filename in items:
            filepath = os.path.join(folder_path, filename)
            
            if os.path.isfile(filepath):
                # 更新日時(mtime)を取得
                mtime = os.path.getmtime(filepath)
                date_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                
                file_data.append({
                    "ファイル名": filename,
                    "最終更新日時": date_str,
                    "フルパス": filepath
                })

        if not file_data:
            raise ValueError("フォルダ内にファイルが見つかりませんでした。")

        # 3. Excel出力
        df = pd.DataFrame(file_data)
        output_path = os.path.join(folder_path, "ファイル更新計画一覧.xlsx")
        df.to_excel(output_path, index=False)
        
        messagebox.showinfo("完了", f"Excelを作成しました：\n{output_path}")

    except PermissionError:
        messagebox.showerror("エラー", "Excelファイルが既に開かれているか、アクセス権限がありません。")
    except Exception as e:
        messagebox.showerror("エラー", f"予期せぬエラーが発生しました:\n{str(e)}")

# --- GUIの基本設定 ---
root = tk.Tk()
root.title("ファイル情報取得ツール")
root.geometry("300x150")

label = tk.Label(root, text="フォルダ内の更新日時を抽出します")
label.pack(pady=20)

button = tk.Button(root, text="フォルダを選択して実行", command=create_file_list)
button.pack(pady=10)

root.mainloop()