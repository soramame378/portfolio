import os
import pandas as pd
from datetime import datetime

def get_file_list(folder_path):
    file_data = []
    
    # フォルダ内のファイルを取得
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        
        # ディレクトリではなくファイルの場合のみ処理
        if os.path.isfile(filepath):
            # 最終アクセス日時（閲覧日時に近い情報）を取得
            access_time = os.path.getatime(filepath)
            # 読みやすい日付形式に変換
            date_str = datetime.fromtimestamp(access_time).strftime('%Y-%m-%d %H:%M:%S')
            
            file_data.append({
                "ファイル名": filename,
                "最終アクセス日時": date_str
            })
    
    return file_data

# 実行テスト
target_dir = './'  # 現在のフォルダを指定
data = get_file_list(target_dir)

# Excelに出力
df = pd.DataFrame(data)
df.to_excel("ファイル一覧.xlsx", index=False)
print("Excelファイルを作成しました。")