import pandas as pd
import numpy as np
from openpyxl.utils.datetime import to_excel


df = pd.read_excel('long form to short form data conversion.xlsx', header=[0, 1])

df.dropna(axis="columns", how='all', inplace=True)
title = ['Test date', 'Energy throughput/KWh', 'Capacity/Ah', 'SOH/%', 'DCIR / mOhm', 'pack66 Ri increase/%']
aa = list(df)  # obtain all headers list
# obtain the four big headers
packs = []
for pack in list(df):
    if pack[0] not in packs:
        packs.append(pack[0])


dfs = []
counts = []
my_style1 = '''
            background-color: yellow
            '''
my_style2 = '''
            background-color: green
            '''
# light blue
my_style3 = '''
            background-color: #00B0F0   
            '''
x = []
y = []
for i, pack in enumerate(packs):
    df_temp = df[pack][title]
    df_temp.dropna(axis='index', how='all', inplace=True)
    df_temp.dropna(axis='columns', how='all', inplace=True)
    df_temp.loc[:, 'pack serial'] = pack
    if i == 0:
        counts.append(len(df_temp))
    else:
        counts.append(counts[i - 1] + len(df_temp))
    x.append(df_temp.loc[:, 'Energy throughput/KWh'].values)
    y.append(df_temp.loc[:, 'SOH/%'].values)
    dfs.append(df_temp)
df = pd.concat(dfs, ignore_index=True)  # combine all csv files together


def highlight_max(s):
    is_max = s == s.max()
    l = []
    for v in is_max:
        if v:
            l.append('background-color:yellow')
        else:
            l.append('')
    #     print(l)
    return l


def highlight_min(s):
    is_min = s == s.min()
    l = []
    for v in is_min:
        if v:
            l.append('background-color:yellow')
        else:
            l.append('')
    #     print(l)
    return l


s1 = df.loc[0:counts[0], 'Energy throughput/KWh']
a = s1[s1 == s1.max()].index

s2 = df.loc[counts[0]:counts[1] - 1, 'SOH/%']
b = s2[s2 == s2.min()].index

s3 = df.loc[counts[1]:counts[2] - 1, 'Energy throughput/KWh']
c = s3[s3 == s3.max()].index
import xlsxwriter

cols = df.shape[1]
writer = pd.ExcelWriter("Task2.xlsx", datetime_format='mmm-d-yyyy', engine='xlsxwriter')
workbook = writer.book
bg_format = workbook.add_format({'bold': False, 'bg_color': '#00B0F0',
                                 'align': 'center', 'valign': 'vcenter', 'font_color': 'black',
                                 'font_size': 10, 'text_wrap': True})
s_name = 'output'
df.style.applymap(lambda x: my_style1, subset=pd.IndexSlice[a, 'Test date']) \
    .applymap(lambda x: my_style2, subset=pd.IndexSlice[b, 'Test date']) \
    .applymap(lambda x: my_style1, subset=pd.IndexSlice[c, 'Test date']) \
    .applymap(lambda x: my_style3, subset=pd.IndexSlice[0, 'pack serial']) \
    .apply(highlight_min, axis=0, subset=pd.IndexSlice[counts[0]:counts[1] - 1, 'SOH/%']) \
    .apply(highlight_max, axis=0, subset=pd.IndexSlice[counts[1]:counts[2] - 1, 'Energy throughput/KWh']) \
    .to_excel(writer, sheet_name=s_name, index=False)

sheet_list = ['Sheet1']
sheet = sheet_list[0]
fmt = workbook.add_format({"font_name": u"Calibri"})
bold = workbook.add_format({
    'bold': False,  # Font Bold
    'border': 1,  # width of cell borders
    'align': 'center',  #
    'valign': 'vcenter',  #
    'text_wrap': True,  # Auto wrap
})
worksheet1 = writer.sheets[s_name]

worksheet1.set_row(0, 20, fmt)
worksheet1.set_column(0, cols, 20, fmt)
worksheet1.conditional_format('A1:G1', {'type': 'no_blanks', 'format': bg_format})

worksheet1.conditional_format('H2:K2', {'type': 'blanks', 'format': bg_format})
worksheet1.conditional_format('A2:F94', {'type': 'no_blanks', 'format': bold})

# plot line chart
color = ['purple', 'blue', 'red', 'green', ]
import matplotlib.pyplot as plt

fig = plt.figure(figsize=(15, 6))
ab = len(x)
for i in range(len(x)):
    label = packs[i]
    plt.plot(list(x[i]), list(y[i]), c=color[i], label=label)
plt.legend()
plt.xlabel('Energy throughput/KWh')
plt.ylabel('SOH/%')
plt.title('SOH/% vs. Energy throughput/KWh')
plt.show()

fig.savefig('foo.png')

worksheet1.insert_image(4, 7, 'foo.png')
writer.save()
