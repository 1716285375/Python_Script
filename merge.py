import pandas as pd
import time


def OpenFile(filename):
    fileData = pd.DataFrame(pd.read_excel(filename, sheet_name=0))
    return fileData


print("##########################################################\n")
print('|------------FUNCTION：合并带日期数据的EXCEL表------------|\n')
print('##########################################################\n')
# 文件数量
num = input('请输入要合并的文件数量：')
filenum = int(num)
filename = []
for i in range(filenum):
    name = input('请输入要合并的文件名：')
    filename.append(name)

out = input('请输入合并后的文件名：')
print('合并后的文件为{outcome}\n'.format(outcome=out))

# 导入文件的数据
data = []           # 数据
columns_len = []    # 表中列数
row_len = []       # 表中行数
print(filename)
for i in range(filenum):
    tmpdata = OpenFile(filename[i]).dropna()
    columns_len.append((len(tmpdata.columns)))
    # print(columns_len)
    row_len.append(len(tmpdata))
    # print(row_len)
    for j in range(row_len[i]):
        col = []
        # print(pd.to_datetime(tmpdata.iloc[j, 0]))
        for k in range(columns_len[i]):
            # 指定选择的列数的数据转化为excel时间格式
            if k > 0:
                col.append(tmpdata.iloc[j, k])
            else:
                col.append(pd.to_datetime(tmpdata.iloc[j, k]))
        data.append(col)


# define colunm name
colums = []
for m in range(columns_len[i]):
    if m == 0:
        index = 'Date'
    else:
        index = 'col {number}'.format(number=m)
    colums.append(index)

newdata = pd.DataFrame(data, columns=colums).dropna()
# 将时间格式转换为字符串格式
# newdata[colums[0]] = newdata[colums[0]].apply(lambda x: x.strftime('%Y'))
newdata.to_excel(out, index=False)
print("##########################################################\n")
print('|----------------   MERGE BEGINNING !   ----------------|\n')
print('##########################################################\n')
print("\n#########################################################\n")

for x in range(0, 100, 9):
    t = float(x) / 1000
    time.sleep(t)
    print('|---------------------合并进度：{num}%--------------------|\n'.format(num=x))

print('|------------------- 合并进度: 100%--------------------|\n')
print("\n##########################################################\n")
print("\n##########################################################\n")
print('|---------------   MERGE SUCCESSFULL !  -----------------|\n')
print('##########################################################\n')
