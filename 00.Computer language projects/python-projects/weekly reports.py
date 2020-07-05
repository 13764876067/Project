'''

自动化办公:

使用pandas操作excel,word生成周报图表实现自动化办公.
'''

'''
pandas

DataFrame

Series

'''

import pandas as pd
import matplotlib.pyplot as plt
'''
读写数据表
'''

'''

#创建DataFram 数据帧,相当于一个sheet
df = pd.DataFrame({
    'id':[1,2,3],
    'name':['lucy','jack','rose'],
    'age':[20,22,16]
})

df = df.set_index('id')
print(df)
df.to_excel('people.xlsx')

print('Done!')

'''


# pd.read_excel

people = pd.read_excel('people.xlsx',sheet_name='Sheet1')
print(people.columns)

p1 = people.sort_values(by='id',ascending=False)
print(p1)

'''
绘制图表
'''
people = pd.read_excel('people.xlsx',sheet_name='Sheet1')
p1 = people.sort_values(by='age',ascending=False)
plt.bar(p1['name'],p1.age,color='orange')
plt.title('people-age',fontsize=16)
plt.xlabel('name')
plt.ylabel('age')

plt.xticks(p1.name,rotation='90')
plt.tight_layout()
plt.show()


p1['total']=p1['id']+p1['age']
p2=p1.sort_values(by='total',ascending=False)
p2.plot.barh(x='name',y=['id','age'],stacked=True)
plt.tight_layout()
plt.show()

'''
from docx import Document
document = Document()
document.save('new.docx')
'''
