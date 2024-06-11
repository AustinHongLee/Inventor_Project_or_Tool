# coding=utf-8

import pandas as pd
import os

# 文件路徑
file_directory = os.path.dirname(os.path.abspath(__file__))    
file_name = 'Pipedetail.xlsx'
file_path = os.path.join(file_directory, file_name)

# 讀取Excel文件中的指定工作表和範圍
df = pd.read_excel(file_path, sheet_name='工作表1', usecols='A:X', nrows=35)  # 假設新增的列位於T列之前

# 檢查讀取的表格（前5行）
print(df.head())

# 準備輸出 Select Case 結構的文本
output_text = ""

# 獲取尺寸列的名稱和厚度列的名稱
size_column = df.columns[0]  # 假設尺寸在第1列
thickness_columns = df.columns[1:16]  # 假設厚度從第2列到第16列
new_columns = df.columns[16:]  # 假設新增的列在第17列之後

# 遍歷尺寸和厚度組合生成 Select Case 語句
for i, row in df.iterrows():
    size = row[size_column]
    output_text += f"Case \"{size}\"\n"
    

    # 處理新增的列
    for new_col in new_columns:
        param_name = new_col
        param_value = row[new_col]
        
        if pd.notna(param_value):  # 檢查值是否為NaN
            output_text += f"    {param_name} = {param_value}\n"

# 打印輸出的文本
print(output_text)

# 將輸出文本寫入文件，使用UTF-8編碼
txt_dir = os.path.join(file_directory, 'select_case.txt')
with open(txt_dir, 'w', encoding='utf-8') as file:
    file.write(output_text)

print(f"Select Case 語句已寫入 {txt_dir}")
