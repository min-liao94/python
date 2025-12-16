import os
import shutil
from pathlib import Path

# 取得當前腳本所在目錄
script_dir = os.path.dirname(os.path.abspath(__file__))
data_dir = os.path.join(script_dir, "data")

# 目標目錄
target_dir = "/home/ivt/docker/file"

if not os.path.isdir(data_dir):
    print(f"找不到 data 資料夾：{data_dir}")
    exit(1)

# 檢查目標目錄是否存在，不存在則建立
if not os.path.exists(target_dir):
    os.makedirs(target_dir)
    print(f"已建立目標資料夾：{target_dir}")

# 遍歷 data 資料夾下的所有工號資料夾
moved_count = 0
skipped_count = 0
error_count = 0

for gonghao_folder in os.listdir(data_dir):
    gonghao_path = os.path.join(data_dir, gonghao_folder)
    
    # 跳過非資料夾的項目
    if not os.path.isdir(gonghao_path):
        continue
    
    # 檢查是否有 files 資料夾
    files_folder = os.path.join(gonghao_path, "files")
    if not os.path.exists(files_folder):
        print(f"跳過：找不到 files 資料夾 {files_folder}")
        continue
    
    # 在 files 資料夾下尋找問卷檔案
    # 支援檔名格式：資訊開頭且 .xlsx 結尾
    for filename in os.listdir(files_folder):
        file_path = os.path.join(files_folder, filename)
        
        # 跳過資料夾
        if os.path.isdir(file_path):
            continue
        
        # 檢查是否符合問卷檔名格式
        if filename.startswith("2025資訊") and filename.endswith(".xlsx"):
            # 目標檔案路徑
            dest_file = os.path.join(target_dir, filename)
            
            # 如果目標檔案已存在，跳過
            if os.path.exists(dest_file):
                print(f"檔案已存在，跳過：{dest_file}")
                skipped_count += 1
                continue
            
            try:
                # 移動檔案
                shutil.move(file_path, dest_file)
                print(f"已移動：{filename} ({gonghao_folder}) -> {dest_file}")
                moved_count += 1
            except Exception as e:
                print(f"移動檔案失敗：{file_path} -> {dest_file} ({e})")
                error_count += 1

print(f"\n處理完成！")
print(f"已移動：{moved_count} 個檔案")
print(f"已跳過：{skipped_count} 個檔案（目標檔案已存在）")
print(f"錯誤：{error_count} 個檔案")

