import pandas as pd
import pymysql, sqlalchemy


conn = sqlalchemy.create_engine('mysql+pymysql://yerdon:Qaz84759@sh-cynosdbmysql-grp-9v8niq3a.sql.tencentcdb.com:21749/hustzz')
doctor_sql = """ SELECT x.XH, x.NJ, x.LQLBMC, x.SXLBMC, x.XZ, y.是否少骨, y.`学籍异动`, y.补发月数, y.原始指标类型, y.修正指标类型, y.硕博贯通, y.出国联培 FROM (SELECT a.XH, a.NJ,a.LQLBMC, a.SXLBMC, a.XZ, XJZTMC FROM GS AS a WHERE a.XH REGEXP 'D') AS x LEFT JOIN zxj_info AS y ON x.XH = y.学号 """
doctor_data = pd.read_sql(doctor_sql, conn)

doctor_data['NJ'] = doctor_data['NJ'].fillna('0').astype(int)   # 年级
doctor_data['XZ'] = doctor_data['XZ'].fillna('0').astype(float) # 学制
doctor_data['补发月数'] = doctor_data['补发月数'].fillna('0').astype(int)   # 补发月数
doctor_data = doctor_data.fillna('')    # 填充空值，否则nan和任何值运算都是nan


def dx(df):
    lqlb = df['LQLBMC']  # 录取类别
    issg = df['是否少骨']  # 是否少骨
    isqj = df['是否强军']  # 是否强军

    if lqlb == '非定向':  # 非定向
        return '3'
    elif lqlb == '定向':  # 定向
        if issg == '是':  # 少骨
            return '1'
        elif isqj == '是':  # 强军
            return '2'
        else:
            return '0'
    else:
        return '0'


def xjyc(df, current_ym):
    xjyd = df['学籍异动']

  
    bq = df['补发月数']

    c_y = int(current_ym.split('-')[0])
    c_m = int(current_ym.split('-')[1])

    if xjyd == '':  # 学籍异动单元格为空
        return '1'  # 没有学籍异动
    else:  # 学籍异动单元格有值
        if bq > 0:  # 有补发月数
            if '休学' in xjyd:  # 休学
                xjyd = xjyd.replace('休学', '', xjyd.count('休学') - 1)  # 保留最后一个休学
                fxny = xjyd[xjyd.find('休学') - 6:xjyd.find('休学')]  # 休学结束时间
                # print(fxny)
                fxny = int(fxny[:4]) * 12 + int(fxny[5:])  # 休学结束时间，注意中间有个.
                if fxny < c_y * 12 + c_m:  # 休学结束时间小于当前时间
                    return '1'  # 没有学籍异动
                else:
                    return '0'  # 有学籍异动
            else:  # 没有休学（应该是档案导致的补发月数）
                return '1'  # 没有学籍异动
        else:
            return '0'  # 有学籍异动


def nx(df, current_ym):
    bq = df['补发月数']
    nj = df['NJ']
    xz = df['XZ']

    bynx = int(nj) * 12 + int(bq) + float(xz) * 12 + 8
    # print(bynx)

    c_y = int(current_ym.split('-')[0])
    c_m = int(current_ym.split('-')[1])

    # print(c_y*12+c_m)
    if bynx > c_y * 12 + c_m:  # 补发月数大于当前时间
        return '1'  # 在年限内
    else:  # 补发月数小于当前时间
        return '0'  # 不在年限内

def cejc(df):
    ys = df['原始指标类型']
    xz = df['修正指标类型']
    if xz == '':
        a = ys
    else:
        a = xz
    if a == '基础':
        return '1'
    else:
        return '0'

def xsls(df, current_ym):
    nj = df['NJ']
    sbgt = df['硕博贯通']
    
    c_y = int(current_ym.split('-')[0])
    c_m = int(current_ym.split('-')[1])
    
    if c_y *12 + c_m < 2023*12+9:
        if nj == 2022 and sbgt == '':
            return '0'
        elif nj == 2022 and sbgt != '':
            return '1'
        else:
            return '2'
    else:
        return '2'

codes = pd.read_sql(" select code, gz, xz, reason from doctor_zxj_code ", conn)
code_dict = {}
for code, gz, xz, reason in codes.values:
    code_dict[str(code)] = [gz, xz, reason]


def match_code(code, which='gz'):
    gz = code_dict[code][0]
    xz = code_dict[code][1]
    reason = code_dict[code][2]
    # print(code)
    if which == 'gz':
        return gz
    elif which == 'xz':
        return xz
    else:
        return reason

first = doctor_data.apply(dx, axis=1)
second = doctor_data['SXLBMC'].apply(lambda x: '0' if "非" in x else '1')
third = doctor_data['出国联培'].apply(lambda x: '1' if "出国" in x else '0')

for current_ym in ['2022-9', '2023-1', '2023-3', '2023-9']:
    fourth = doctor_data.apply(xsls, current_ym=current_ym, axis=1)
    fifth = doctor_data['学籍异动'].apply(lambda x: '1' if x == '' else '0')
    sixth = doctor_data.apply(nx, axis=1, current_ym=current_ym)
    seventh = doctor_data.apply(cejc, axis=1)
    code = first + second + third + fourth + fifth + sixth + seventh

    doctor_data[current_ym+'_国助'] = code.apply(match_code, which='gz')
    doctor_data[current_ym + '_校助'] = code.apply(match_code, which='xz')
    doctor_data[current_ym + '_理由'] = code.apply(match_code, which='reason')


doctor_data.to_excel(r'D:\Users\yerdon\Desktop\D.xlsx', index=False)

