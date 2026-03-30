import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

def create_file_list():
    # 入力と出力の選択
    input_folder = filedialog.askdirectory(title="調査したいフォルダを選択")
    if not input_folder: return
    output_folder = filedialog.askdirectory(title="保存先を選択")
    if not output_folder: return

    try:
        file_data = []
        for filename in os.listdir(input_folder):
            # 一時ファイル除外
            if filename.startswith('~$') or filename.startswith('.'): continue
            
            filepath = os.path.join(input_folder, filename)
            if os.path.isfile(filepath):
                mtime = os.path.getmtime(filepath)
                date_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                file_data.append({"ファイル名": filename, "更新日時": date_str})

        if not file_data:
            raise ValueError("ファイルが見つかりませんでした。")

        # 保存（装飾なしの標準出力）
        df = pd.DataFrame(file_data)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(output_folder, f"一覧_{timestamp}.xlsx")
        
        # エンジンを指定せず、最も標準的な方法で保存
        df.to_excel(output_path, index=False)
        
        messagebox.showinfo("成功", f"作成完了:\n{output_path}")

    except Exception as e:
        # エラーの「型」と「内容」を正確に表示させる
        import traceback
        error_details = traceback.format_exc()
        print(error_details) # ターミナルに詳細を出す
        messagebox.showerror("エラー発生", f"{str(e)}\n\n詳細はターミナルを確認してください")

# GUI
root = tk.Tk()
root.title("簡易ファイルチェッカー")
root.geometry("300x100")
tk.Button(root, text="実行", command=create_file_list).pack(expand=True)
root.mainloop()