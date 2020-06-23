import xlwt
import pandas as pd
import xlwings as xw

temp = {"a":[1,2,3], "b":[1,2,3]}
data = pd.DataFrame(temp)
print(data)
print(data.iloc[1,1])



'''
workbook = xw.Book(r'C:\\Users\87010\Desktop\兰桂骐工作文件\爬虫\实验创建添加表格.xls')
sheet1 = workbook.sheets.add(name = 'xlwings试验添加表格1')
sheet2 = workbook.sheets.add(name = 'xlwings试验添加表格2')

sheet1.range("A1").value = data.values
workbook.save()
workbook.close()
'''