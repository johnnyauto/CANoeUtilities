'''
Generate CAPL scripts fo HI_Sysvar init.
'''

import pandas as pd

# read Excel
sourceData = pd.read_excel('VIU_TR05_信號定義_V04_20240201_source.xlsx', sheet_name='硬線_CEM')

# output content
# navigate each row in the Excel
# Init scripts for IDH signals
output_text = '// IDH HI_Sysvar\n'
for index, row in sourceData.iterrows():
    signal_name = row['Signal Name (Eng)']
    io_type = row['I/O Type']
    in_out = row['IN/OUT']

    if not pd.isna(signal_name) and in_out=='IN' and io_type=='IDH':
        output_text += f'@sysvar::HIO::HI_{signal_name} = 0;\n'

# Init scripts for IDL signals
output_text += '\n// IDL HI_Sysvar\n'
for index, row in sourceData.iterrows():
    signal_name = row['Signal Name (Eng)']
    io_type = row['I/O Type']
    in_out = row['IN/OUT']

    if not pd.isna(signal_name) and (io_type=='IDL' or io_type=='IAN'):
        output_text += f'@sysvar::HIO::HI_{signal_name} = 1;\n'

# create a vssysvar file
with open('HI_Sysvar_init_script.txt', 'w', encoding='utf-8') as f:
    f.write(output_text)

print("HI_Sysvar init sctipts is generated.")
