# 假设sig欄位的字串列表保存在一個名為data的變數中
data = ['sig1_HUT', 'sig2', 'sig3_HUT', 'sig4']

for string in data:
    if '_HUT' in string:
        new_string = string.replace('_HUT', '')
        print(new_string)
