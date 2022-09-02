import pandas as pd
import sqlalchemy




conn = sqlalchemy.create_engine('mysql+pymysql://yerdon:Qaz84759@sh-cynosdbmysql-grp-9v8niq3a.sql.tencentcdb.com:21749/hustzz')
sql = " SELECT 学号, 姓名, 单位名称, 身份证件类型, 身份证号 FROM GS_status "
sql1 = " SELECT 学号, 姓名, 摘要 FROM gszz_liushui WHERE 账务年 = '2021' AND 统计口径 = '学校助学金' "
df = pd.read_sql(sql, conn)
df_x = pd.read_sql(sql1, conn)
if '于2021' in df_x['摘要'].values:
    df_x['发放人'] = df_x['摘要'].str.split('于2021', 2).str[0]
else:
    if '学院' in df_x['摘要'].values:
        df_x['发放人'] = df_x['摘要'].str.split('发放', 2).str[0].str.split('学院', 2).str[1]
    else:
        df_x['发放人'] = df_x['摘要'].str.split('发放', 2).str[0].str.split('中心', 2).str[1]


df1 = pd.read_excel(r'C:\Users\王子祺\Desktop\匹配\2022年5月学校助学金导入数据.xlsx', dtype=object)
df1 = pd.merge(df1, df.loc[:, ['身份证号', '身份证件类型', '学号']], on='学号', how='left')


df_2021硕 = pd.read_excel(r'C:\Users\王子祺\Desktop\匹配\2022年5月国家助学金导入数据-2021硕.xlsx', dtype=object)
df_2021硕 = pd.merge(df_2021硕, df.loc[:, ['身份证号', '身份证件类型', '学号']], on='学号', how='left')