'''
Copy VT System variables to corresponding hardware signal.
'''

import pandas as pd


# 读取 source.xlsx 和 target.xlsx 文件中指定的分页
source_sheet_name = '硬線_CEM'
target_sheet_name = '引脚需求匹配'

# 讀取檔案
sourceFile = 'VIU_TR05_信號定義_V04_20240201_source.xlsx'
targetFile = 'VT Configs_CEM_target.xlsx'
df_source = pd.read_excel(sourceFile, sheet_name=source_sheet_name)
df_target = pd.read_excel(targetFile, sheet_name=target_sheet_name)

# 创建一个字典，将 Function Definition 列中的值映射到对应的 VT sysvar和Default value
function_to_vtsysvar = dict(zip(df_target['Function Definition (TAO_CEM)'], df_target['VT sysvar']))
#function_to_defaultvalue = dict(zip(df_target['Function Definition (TAO_CEM)'], df_target['Default value']))
function_to_iotype = dict(zip(df_target['Function Definition (TAO_CEM)'], df_target['I/O Type']))

# 使用 map 函数将 sourceFile的Signal Name 列中的值映射到 targetFile的VT sysvar和Default value 值
df_source['VT Sysvar'] = df_source['Signal Name (Eng)'].map(function_to_vtsysvar)
#df_source['Default Value'] = df_source['Signal Name (Eng)'].map(function_to_defaultvalue)
df_source['I/O Type'] = df_source['Signal Name (Eng)'].map(function_to_iotype)

# 保存到 source.xlsx 文件中指定的分页
with pd.ExcelWriter(sourceFile, engine='openpyxl', mode='a') as writer:
    df_source.to_excel(writer, index=False)

print("The data merge is completed and appended to the last sheet in the source file. ")
