'''
Generate System variable with predix (HI_/HO_) based on signale name.
'''

import pandas as pd

# read Excel
sourceData = pd.read_excel('VIU_TR05_信號定義_V04_20240201_source.xlsx', sheet_name='硬線_CEM')

# output content
output_text = '<?xml version="1.0" encoding="utf-8"?>\n'
output_text += '<systemvariables version="4">\n'
output_text += '  <namespace name="" comment="" interface="">\n'
output_text += '    <namespace name="HIO" comment="" interface="">\n'

# navigate each row in the Excel
for index, row in sourceData.iterrows():
    signal_name = row['Signal Name (Eng)']
    in_out = row['IN/OUT']

    if not pd.isna(signal_name):  # if 'signal_name' is not empty
        if in_out == "IN":
            prefix = "HI_"
        elif in_out == "OUT":
            prefix = "HO_"
        else:
            prefix = ""

        modified_signal_name = prefix + signal_name

        # add modified content into output content
        output_text += f'      <variable anlyzLocal="2" readOnly="false" valueSequence="false" unit="" name="{modified_signal_name}" comment="" bitcount="32" isSigned="true" encoding="65001" type="int" startValue="0">\n'
        output_text += '      </variable>\n'

# finish the output content
output_text += '    </namespace>\n'
output_text += '  </namespace>\n'
output_text += '</systemvariables>'

# create a vssysvar file
with open('HIO.vsysvar', 'w', encoding='utf-8') as f:
    f.write(output_text)

print("HIO.vsysvar is generated.")
