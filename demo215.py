# _*_ Author:JackZhang9
# _*_ Time:20230215
'''读取excel中的数据，并保存为新excel'''
import pandas as pd
import os
import numpy as np

pd.set_option('display.max_columns',None)
file_dir='文件目录名'
if not os.path.exists(file_dir):
    os.mkdir(file_dir)

file_name='excel文件名.xlsx'
data_read=pd.ExcelFile(file_name)  # 读取文件对象,是一个个sheet
data1=pd.read_excel(data_read)  # 只能读取第一张sheet
# print(data_read.sheet_names)

degrees=set() # 所有指定column的集合
for sheet_name in data_read.sheet_names:
    data=pd.read_excel(data_read,sheet_name=sheet_name)  # 循环读取每一张sheet
    degrees=degrees.union(set(data['columns_name']))  # 每一个sheet里面的某个series做并集操作，然后赋值给这个集合

# print(degrees)
'''用一个新的excel表格保存新的sheet'''
for degree in degrees:
    file_path=os.path.join(file_dir,'{}.xlsx'.format(degree))
    data_write=pd.ExcelWriter(file_path)
    for sheet_name in data_read.sheet_names:
        sheet_data=pd.read_excel(data_read,sheet_name=sheet_name)  # 每一张sheet的数据
        data_sub=sheet_data[sheet_data['column']==degree]   # 提取符合某一column的所有行的dataframe
        data_sub=data_sub.astype('str')
        # data=np.random.randint(0,100,(10,10))
        # data_sub=pd.DataFrame(data)
        data_sub.to_excel(data_write,sheet_name=sheet_name)
    data_write.save()
    data_write.close()
