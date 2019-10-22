import numpy as np
import pandas as pd
path=input('请输入bom.xls文件完整路径如：c:/bom.xls(注意是反斜杆/) : ')
path2=input('请输入查询条件.xlsx文件的完整路径如：c:/查询条件.xlsx : ')
df=pd.read_excel(path,sheet_name='BOM多级展开') #读取主要的数据
df1=pd.read_excel(path,sheet_name='素材件')
df2=pd.read_excel(path,sheet_name='辅料')
dftj=pd.read_excel(path2)  #读取查询条件数据，包含产品代码和数量
cpdm=list(dftj.产品代码) #获取产品代码并转换成list类型
qs=list(dftj.数量)  #获取所需数量 并转换成list
df_1=df[['产品代码','产品名称','序号','材料代码','材料名称','材料单位','产品用量','材料属性']]  #提取出所需结果模板
df_empty=pd.DataFrame()  #创建一个空的数据框 以存取查询结果
for x,q in zip(cpdm,qs):  #使用zip方法使两个list可以一起迭代
 df_chaxun=df_1[df_1['产品代码']==x]  #循环提取产品代码对应的结果
 df_chaxun['成品需求']=q  #新建一列循环写入 查询条件里的数量
 df_chaxun['产品实际用量']=df_chaxun['产品用量'].apply(lambda x:x*q) #新建一列循环写入产品的实际用量 产品用量*查询条件里面的数量
 df_empty=df_empty.append(df_chaxun) #将查询和计算的结果追加到 刚刚新建的空数据框
print(df_empty)
path3=input('请输入查询结果.xlsx完整路径如：c:/查询结果.xlsx : ') #交互输入需要导出的查询结果路径和文件名
df_empty.to_excel(path3,encoding='utf8',index=False) #导出查询结果 将index忽略掉
print('查询结果已导出')
input('输入任意键退出程序')  
