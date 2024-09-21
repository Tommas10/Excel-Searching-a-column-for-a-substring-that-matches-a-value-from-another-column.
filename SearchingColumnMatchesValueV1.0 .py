# -*- coding: utf-8 -*-
"""
#Small automation Python script- Searching a column for a substring that matches a value from another column.

#After opening test.xls To fill in the Result column with values ​​of Third_Value column corresponding to values ​​in the #Longer_Value column if the value from the first column was found. 
#If more than one was found, I would like to list them with some delimiter.
#If substring not matches a value from anither column to write no_match. 
#Then save it as Test1.xlsx

#Created by Tommas Huang 
#Created date: 2024-09-021

"""

import pandas as pd  # 引入 pandas 模組，用於數據處理
import os  # 引入 os 模組，用於檔案系統操作

# 修改檔案路徑
file_path = r'D:\01\test.xlsx'  # Set the file path for the Excel file

# Check if the file exists
# 檢查檔案是否存在
if os.path.exists(file_path): 
    # Print file existence message
    print(f"檔案 {file_path} 存在，正在讀取...")  

    try:
        # Read the Excel file into a DataFrame
        # 讀取 Excel 檔案
        df = pd.read_excel(file_path)  
        
        # Print the number of rows read
        # 檢查是否正確讀取資料
        print(f"檔案讀取成功，共有 {len(df)} 行資料。")  

        # Print the column names
        # 列出欄位名稱
        print(f"欄位名稱: {df.columns.tolist()}")  

        # List of required columns
        # 檢查必要的欄位是否存在
        required_columns = ['Value', 'Longer_Value', 'Third_Value'] 
        # Loop through the required columns
        for col in required_columns: 
            # Check if the column is missing
            if col not in df.columns:  
                # Raise an error if a column is missing
                raise ValueError(f"缺少必要的欄位：{col}")  

        # Define a function to match values
        # 定義一個函數來檢查 'Value' 是否在 'Longer_Value' 中
        def match_values(row, df): 
            # Find matches in 'Longer_Value'
            matches = df[df['Longer_Value'].str.contains(str(row['Value']), na=False)]  
            # Check if matches were found
            if not matches.empty:  
                # Return matching 'Third_Value'
                return ', '.join(matches['Third_Value'].tolist()) 
            # Return 'no_match' if no matches found
            return 'no_match'  

        # Apply matching function to each row
        # 對每一行應用匹配函數，並將結果填入 'Result' 欄位
        df['Result'] = df.apply(lambda row: match_values(row, df), axis=1)  

        # Set the output file path
        # 將結果儲存為新的 Excel 檔案
        output_file_path = r'D:\01\Test1.xlsx'  
        # Save the DataFrame to an Excel file
        df.to_excel(output_file_path, index=False) 
        
        # Print success message
        print(f"檔案已儲存為 {output_file_path}") 
    
    # Handle any exceptions that occur
    except Exception as e: 
        # Print error message
        print(f"讀取或處理檔案時發生錯誤: {e}") 

else:
    # Print error message if file does not exist
    print(f"檔案 {file_path} 不存在。請檢查路徑或副檔名是否正確。") 






