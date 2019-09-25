from pymongo import MongoClient
import xlwt
import pandas as pd

def load_from_MongoDB(dbname,colname):
    conn = MongoClient('127.0.0.1', 27017)
    db = conn[dbname]
    # collist = db.list_collection_names()
    # print(collist)
    tech_mgdb = db[colname]
    return tech_mgdb.find()

def export_excel(file,file_name):
   #将字典列表转换为DataFrame
   pf = pd.DataFrame(list(file))
   #指定字段顺序
   order = ['Title','URL','Summary','Category','Investment Status','Medical Field',
            'Patent Status','Development Status','Medical Center','Inventors','Content']
   pf = pf[order]
   #将列名替换为中文
   columns_map = {
       'Title':'标题名',
       'URL':'网址',
       'Summary':'摘要',
       'Category':'Category',
       'Investment Status':'Investment Status',
       'Medical Field':'Medical Field',
       'Patent Status':'Patent Status',
       'Development Status':'Development Status',
       'Medical Center':'Medical Center',
       'Inventors':'Inventors',
       'Content':'Content'
   }
   pf.rename(columns = columns_map,inplace = True)
   #指定生成的Excel表格名称
   excel_file = pd.ExcelWriter('%s.xlsx' % file_name)
   #替换空单元格
   #pf.fillna(' ',inplace = True)
   #输出
   pf.to_excel(excel_file,encoding = 'utf-8',index = False)
   #保存表格
   excel_file.save()




if __name__ == '__main__':
    dbname = 'Technologies爬取库'
    colname = '科技类别信息_201909'
    loadfile = load_from_MongoDB(dbname,colname)
    file_name = 'Technologies'
    export_excel(loadfile, file_name)


