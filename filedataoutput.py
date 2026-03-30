import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

def create_file_list():
    folder_path = filedialog.askdirectory()
    if not folder_path:
        return

    try:
        file_data = []
        items = os.listdir(folder_path)
        
        if not items:
            raise ValueError("選択されたフォルダは空です。")

        for filename in items:
            # IT事務のこだわりポイント①：一時ファイルや隠しファイルを除外
            if filename.startswith('~$') or filename.startswith('.'):
                continue
                
            filepath = os.path.join(folder_path, filename)
            
            if os.path.isfile(filepath):
                mtime = os.path.getmtime(filepath)
                date_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                
                file_data.append({
                    "ファイル名": filename,
                    "最終更新日時": date_str,
                    "フルパス": filepath
                })

        if not file_data:
            raise ValueError("有効なファイルが見つかりませんでした。")

        # データの構造化
        df = pd.DataFrame(file_data)
        output_path = os.path.join(folder_path, "ファイル更新計画一覧.xlsx")

        # IT事務のこだわりポイント②：Excelの見た目を整える
        # engine='xlsxwriter' を使うと細かい書式設定がしやすいです
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='一覧')
            
            # 列幅の自動調整ロジック
            worksheet = writer.sheets['一覧']
            for i, col in enumerate(df.columns):
                # 列名の長さとデータ内の最大文字数を比較して幅を決定
                column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)

        messagebox.showinfo("完了", f"Excelを作成しました：\n{output_path}")

    except PermissionError:
        messagebox.showerror("エラー", "出力先のExcelが既に開かれています。閉じてから再実行してください。")
    except Exception as e:
        messagebox.showerror("エラー", f"予期せぬエラーが発生しました:\n{str(e)}")

# GUI設定（前回同様）
root = tk.Tk()
root.title("ファイル管理ツール")
root.geometry("350x150")
tk.Label(root, text="フォルダ内の更新日時をExcelに出力します").pack(pady=20)
tk.Button(root, text="フォルダを選択して実行", command=create_file_list, bg="#e1e1e1").pack(pady=10)
root.mainloop()