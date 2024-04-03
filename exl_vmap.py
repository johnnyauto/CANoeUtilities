''' generate *.vmap for CANoe Symbol Mapping '''

import pandas as pd

# read Excel
sourceData = pd.read_excel('VIU_TR05_信號定義_V04_20240201_source.xlsx', sheet_name='硬線_CEM')

'''-----------------------------------------------------------'''
############## output content 1st -- HI_VT_Sysvar ##############
'''-----------------------------------------------------------'''
HI_output_text = '<?xml version="1.0" encoding="utf-8"?>\n'
HI_output_text += '<Mapping Version="5">\n'
HI_output_text += '  <RuntimeGroup Active="True" Name="Static Mapping">\n'
HI_output_text += '    <LogicalGroup Active="True" Name="HI_VT_Sysvar" IsCustomBinding="False">\n'

# navigate each row in the Excel
for index, row in sourceData.iterrows():
    signal_name = row['Signal Name (Eng)']
    vt_sysvar = row['VT Sysvar']
    in_out = row['IN/OUT']
    #reverse_mapping =row['Reverse Mapping']
    reverse_mapping =row['I/O Type']
    
    
    # HI_Sysvar mapping to VT_Sysvar
    if not pd.isna(vt_sysvar):  # if 'vt_sysvar' is not empty
        if in_out == "IN":
            map_source = "HIO::HI_" + signal_name
            map_dest = vt_sysvar

            if reverse_mapping == 'IDL' or reverse_mapping == 'IAN':
                # Reverse Mapping for IDL signal
                HI_output_text += '      <MappingRelation Active="True" DirectionVar1ToVar2="True" Offset="1" Factor="-1" Assignment="OnChange" CyclicTimerValue="10">\n'
            else:
                HI_output_text += '      <MappingRelation Active="True" DirectionVar1ToVar2="True" Offset="0" Factor="1" Assignment="OnChange" CyclicTimerValue="10">\n'
            HI_output_text += '        <ValueObject1 ItemType="symSystemVariable">\n'
            HI_output_text += '          <BusType>-1</BusType>\n'
            HI_output_text += f'          <DatabaseName>{map_source}</DatabaseName>\n'
            HI_output_text += '          <Description />\n'
            HI_output_text += '          <EnvVarName />\n'
            HI_output_text += f'          <FullName>{map_source}</FullName>\n'
            HI_output_text += '          <ID>2</ID>\n'
            HI_output_text += '          <MessageName />\n'
            HI_output_text += f'          <Name>{map_source}</Name>\n'
            HI_output_text += '          <NeedsMessage>False</NeedsMessage>\n'
            HI_output_text += '          <NetworkName />\n'
            HI_output_text += '          <NodeName />\n'
            HI_output_text += '          <SignalName />\n'
            HI_output_text += '          <VariableType>1</VariableType>\n'
            HI_output_text += '        </ValueObject1>\n'
            HI_output_text += '        <ValueObject2 ItemType="symSystemVariable">\n'
            HI_output_text += '          <BusType>-1</BusType>\n'
            HI_output_text += f'          <DatabaseName>{map_dest}</DatabaseName>\n'
            HI_output_text += '          <Description />\n'
            HI_output_text += '          <EnvVarName />\n'
            HI_output_text += f'          <FullName>{map_dest}</FullName>\n'
            HI_output_text += '          <ID>2</ID>\n'
            HI_output_text += '          <MessageName />\n'
            HI_output_text += f'          <Name>{map_dest}</Name>\n'
            HI_output_text += '          <NeedsMessage>False</NeedsMessage>\n'
            HI_output_text += '          <NetworkName />\n'
            HI_output_text += '          <NodeName />\n'
            HI_output_text += '          <SignalName />\n'
            HI_output_text += '          <VariableType>1</VariableType>\n'
            HI_output_text += '        </ValueObject2>\n'
            HI_output_text += '      </MappingRelation>\n'
        else:
            pass

'''
# VT_Sysvar mapping to HI_Sysvar (for reset Sysvar)
for index, row in sourceData.iterrows():
    signal_name = row['Signal Name (Eng)']
    vt_sysvar = row['VT Sysvar']
    in_out = row['IN/OUT']
    reverse_mapping =row['Reverse Mapping']
    
    # VT_Sysvar mapping to HI_Sysvar
    if not pd.isna(vt_sysvar):  # if 'vt_sysvar' is not empty
        if in_out == "IN":
            map_source = vt_sysvar
            map_dest = "HIO::HI_" + signal_name
            
            if reverse_mapping == 'y':
                HI_output_text += '      <MappingRelation Active="True" DirectionVar1ToVar2="True" Offset="1" Factor="-1" Assignment="OnChange" CyclicTimerValue="10">\n'
            else:
                HI_output_text += '      <MappingRelation Active="True" DirectionVar1ToVar2="True" Offset="0" Factor="1" Assignment="OnChange" CyclicTimerValue="10">\n'
            HI_output_text += '        <ValueObject1 ItemType="symSystemVariable">\n'
            HI_output_text += '          <BusType>-1</BusType>\n'
            HI_output_text += f'          <DatabaseName>{map_source}</DatabaseName>\n'
            HI_output_text += '          <Description />\n'
            HI_output_text += '          <EnvVarName />\n'
            HI_output_text += f'          <FullName>{map_source}</FullName>\n'
            HI_output_text += '          <ID>2</ID>\n'
            HI_output_text += '          <MessageName />\n'
            HI_output_text += f'          <Name>{map_source}</Name>\n'
            HI_output_text += '          <NeedsMessage>False</NeedsMessage>\n'
            HI_output_text += '          <NetworkName />\n'
            HI_output_text += '          <NodeName />\n'
            HI_output_text += '          <SignalName />\n'
            HI_output_text += '          <VariableType>1</VariableType>\n'
            HI_output_text += '        </ValueObject1>\n'
            HI_output_text += '        <ValueObject2 ItemType="symSystemVariable">\n'
            HI_output_text += '          <BusType>-1</BusType>\n'
            HI_output_text += f'          <DatabaseName>{map_dest}</DatabaseName>\n'
            HI_output_text += '          <Description />\n'
            HI_output_text += '          <EnvVarName />\n'
            HI_output_text += f'          <FullName>{map_dest}</FullName>\n'
            HI_output_text += '          <ID>2</ID>\n'
            HI_output_text += '          <MessageName />\n'
            HI_output_text += f'          <Name>{map_dest}</Name>\n'
            HI_output_text += '          <NeedsMessage>False</NeedsMessage>\n'
            HI_output_text += '          <NetworkName />\n'
            HI_output_text += '          <NodeName />\n'
            HI_output_text += '          <SignalName />\n'
            HI_output_text += '          <VariableType>1</VariableType>\n'
            HI_output_text += '        </ValueObject2>\n'
            HI_output_text += '      </MappingRelation>\n'
        else:
            pass
'''

# finish the output content
HI_output_text += '    </LogicalGroup>\n'
HI_output_text += '  </RuntimeGroup>\n'
HI_output_text += '</Mapping>'

# create a *.vmap file
with open('VTsysvar_HI_mapping.vmap', 'w', encoding='utf-8') as f:
    f.write(HI_output_text)

print("VTsysvar_HI_mapping.vmap is created")


'''-----------------------------------------------------------'''
############## output content 2nd -- HO_VT_Sysvar ##############
'''-----------------------------------------------------------'''
HO_output_text = '<?xml version="1.0" encoding="utf-8"?>\n'
HO_output_text += '<Mapping Version="5">\n'
HO_output_text += '  <RuntimeGroup Active="True" Name="Static Mapping">\n'
HO_output_text += '    <LogicalGroup Active="True" Name="HO_VT_Sysvar" IsCustomBinding="False">\n'

# VT_Sysvar mapping to HO_Sysvar
for index, row in sourceData.iterrows():
    signal_name = row['Signal Name (Eng)']
    vt_sysvar = row['VT Sysvar']
    in_out = row['IN/OUT']
    #default_value = row['Default Value']
    #reverse_mapping =row['Reverse Mapping']

    if not pd.isna(vt_sysvar):  # if 'vt_sysvar' is not empty
        if in_out == "OUT":
            map_source = vt_sysvar
            map_dest = "HIO::HO_" + signal_name

            '''
            #if default_value == 1:
            if reverse_mapping == 'y':
                # for VT2516A LSD config
                print('LSD')
                HO_output_text += '      <MappingRelation Active="True" DirectionVar1ToVar2="True" Offset="1" Factor="-1" Assignment="OnChange" CyclicTimerValue="10">\n'
            else:
                HO_output_text += '      <MappingRelation Active="True" DirectionVar1ToVar2="True" Offset="0" Factor="1" Assignment="OnChange" CyclicTimerValue="10">\n'
            '''
            HO_output_text += '      <MappingRelation Active="True" DirectionVar1ToVar2="True" Offset="0" Factor="1" Assignment="OnChange" CyclicTimerValue="10">\n'
            HO_output_text += '        <ValueObject1 ItemType="symSystemVariable">\n'
            HO_output_text += '          <BusType>-1</BusType>\n'
            HO_output_text += f'          <DatabaseName>{map_source}</DatabaseName>\n'
            HO_output_text += '          <Description />\n'
            HO_output_text += '          <EnvVarName />\n'
            HO_output_text += f'          <FullName>{map_source}</FullName>\n'
            HO_output_text += '          <ID>2</ID>\n'
            HO_output_text += '          <MessageName />\n'
            HO_output_text += f'          <Name>{map_source}</Name>\n'
            HO_output_text += '          <NeedsMessage>False</NeedsMessage>\n'
            HO_output_text += '          <NetworkName />\n'
            HO_output_text += '          <NodeName />\n'
            HO_output_text += '          <SignalName />\n'
            HO_output_text += '          <VariableType>1</VariableType>\n'
            HO_output_text += '        </ValueObject1>\n'
            HO_output_text += '        <ValueObject2 ItemType="symSystemVariable">\n'
            HO_output_text += '          <BusType>-1</BusType>\n'
            HO_output_text += f'          <DatabaseName>{map_dest}</DatabaseName>\n'
            HO_output_text += '          <Description />\n'
            HO_output_text += '          <EnvVarName />\n'
            HO_output_text += f'          <FullName>{map_dest}</FullName>\n'
            HO_output_text += '          <ID>2</ID>\n'
            HO_output_text += '          <MessageName />\n'
            HO_output_text += f'          <Name>{map_dest}</Name>\n'
            HO_output_text += '          <NeedsMessage>False</NeedsMessage>\n'
            HO_output_text += '          <NetworkName />\n'
            HO_output_text += '          <NodeName />\n'
            HO_output_text += '          <SignalName />\n'
            HO_output_text += '          <VariableType>1</VariableType>\n'
            HO_output_text += '        </ValueObject2>\n'
            HO_output_text += '      </MappingRelation>\n'
        else:
            pass

# finish the output content
HO_output_text += '    </LogicalGroup>\n'
HO_output_text += '  </RuntimeGroup>\n'
HO_output_text += '</Mapping>'

# create a *.vmap file
with open('VTsysvar_HO_mapping.vmap', 'w', encoding='utf-8') as f:
    f.write(HO_output_text)

print("VTsysvar_HO_mapping.vmap is created")