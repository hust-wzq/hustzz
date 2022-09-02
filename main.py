#


# pandas读取表格第二到五列
import pandas as pd
df = pd.read_excel(r'C:\Users\王子祺\Desktop\截止2022年4月15日尚未领取合同研究生196人.xlsx')
df = df.iloc[:,1:6]




# pandas读取xslx表格前10行
import pandas as pd
df = pd.read_excel(r'C:\Users\王子祺\Desktop\截止2022年4月15日尚未领取合同研究生196人.xlsx')
print(df.head(10))
df2 = df.iloc[:,]

# 传参数
import pandas as pd
df = pd.read_excel(r'C:\Users\王子祺\Desktop\截止2022年4月15日尚未领取合同研究生196人.xlsx',sheet_name='Sheet1')
print(df.head(10))
df2 = df.iloc[:,]

class BuildRobot():
    def __init__(self,armcount,headcount):
        self.armcount = armcount
        self.headcount = headcount













