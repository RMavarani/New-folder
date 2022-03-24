
import pandas as pd
class search:
    col =None
    row = None
    name = None
    sheet = None
    value=None
    def __init__(self,col= None,row = None, name= None,sheet = None):
        self.col = col
        self.row = row
        self.name = name
        self.sheet = sheet
    def __str__(self):
        return "name:"+str(self.name)+" col:"+str(self.col)+" row:"+str(self.row)+" sheet:"+str(self.sheet)+" value:"+str(self.value)

    def get_value_location(self):
        if self.name=="job" :
            index = 1
        else:
            index = 2
        return index+ self.col
    def get_col(self):
        return self.col
    
    def get_row(self):
        return self.row

    def get_name(self):
        return self.name 

    def get_sheet(self):
        return self.sheet 
    
def read_my_xls(filename,Title):
    wb=pd.read_excel(filename, sheet_name= None, index_col=0)
    dic ={}
    column_names=[]
    rownumber=0   
    
    for col in wb["sheet1"]:
        column_names.insert(rownumber,col)
        x=0        
        for rows in wb["sheet1"][col]:
                
                y=rownumber
                for tit in Title:
                    if rows == tit:
                        record=search(y,x,rows,"sheet1")
                        print (record.name, record.col, record.row, record.sheet)
                        dic[rows]=record
                x+=1
        rownumber += 1
    return wb,dic,column_names
                        
def get_values(temp,wb):
    return wb[temp.get_sheet()][column_names[temp.get_value_location()]].iloc[temp.get_row()]
                    
dic ={}                    
column_names=[]
Title=['a','job','c','b','d']
                                    
result= read_my_xls("./excample.xlsx",Title)
wb=result[0]
dic=result[1]
column_names=result[2]

for tit in Title:
    dic[tit].value=get_values(dic[tit],wb)
    print(dic[tit])


