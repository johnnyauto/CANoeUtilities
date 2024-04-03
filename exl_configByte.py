'''
Generate System variable for config bytes.
'''

import pandas as pd

# read Excel
sourceData = pd.read_excel('VIU_TR05_信號定義_V04_configByte.xlsx', sheet_name='CAN03_Matrix', dtype='str')

# output content
output_text = '<?xml version="1.0" encoding="utf-8"?>\n'
output_text += '<systemvariables version="4">\n'
output_text += '  <namespace name="" comment="" interface="">\n'
output_text += '    <namespace name="config_CEM" comment="" interface="">\n'

# navigate each row in the Excel
for index, row in sourceData.iterrows():
    signal_name = row['Signal Name']
    init_val = row['Default Initialised value']

    if not pd.isna(signal_name):  # if 'signal_name' is not empty
        #modified_signal_name = "cfsv_" + signal_name

        # add modified content into output content
        output_text += f'      <variable anlyzLocal="2" readOnly="false" valueSequence="false" unit="" name="{signal_name}" comment="" bitcount="32" isSigned="true" encoding="65001" type="int" startValue="{init_val}">\n'
        output_text += '      </variable>\n'

# finish the output content
output_text += '    </namespace>\n'
output_text += '  </namespace>\n'
output_text += '</systemvariables>'

# create a vssysvar file
with open('config_CEM.vsysvar', 'w', encoding='utf-8') as f:
    f.write(output_text)

print("config_CEM.vsysvar file is created")
