import pandas as pd, pymysql, numpy as np
from sqlalchemy import *


class XueYeJiang:
    def __init__(self, y1=9, y2=7, y3=5, y4=3, y5=1, b1=18210, b2=18210, b3=22210):
        # 初始化数据库连接
        self.conn = create_engine(
            'mysql+pymysql://yerdon:Qaz84759@sh-cynosdbmysql-grp-9v8niq3a.sql.tencentcdb.com:21749/hustzz')

        # 初始化表格
        self.writer = pd.ExcelWriter("static/res.xlsx")

        # 1.主表
        self.data = pd.DataFrame()  # 生成主表
        self.init_df()  # 读学籍库
        self.stu_qlt(y1, y2, y3, y4, y5)  # 算生源质量
        self.data.to_excel(self.writer, sheet_name='1.学籍信息', index=False)

        # 2.生成生源类型统计表
        self.sylxtj = self.data.pivot_table(values='XH', aggfunc=['count'], index=['DWBH', 'DWMC'],
                                            columns=['NJ', 'sylx'])
        self.sylxtj.to_excel(self.writer, sheet_name='2.生源类型统计')

        # 3.生成分数、人数表
        self.fsrs = self.data.pivot_table(values='sydf', aggfunc=['sum', 'count'], index=['DWBH', 'DWMC'],
                                          columns=['NJ'])
        self.fsrs.to_excel(self.writer, sheet_name='3.分数人数')

        # 4.生成Z-Score、预算表，
        self.z_score = self.get_z_score()
        self.z_score.to_excel(self.writer, sheet_name='4.得分、预算')

        # 5.生成微调基础
        self.modify_basis = self.modify_basis()
        self.modify_basis.to_excel(self.writer, sheet_name='5.调整基础')

        # 输入预算
        self.budget = {'2020': b1, '2021': b2, '2022': b3}

        # 6.调整
        self.data = self.modify_basis.fillna(0)
        self.res = Modify.modify(data=self.data, budget=self.budget)
        self.res.to_excel(self.writer, sheet_name='6.调整结果')

        # 7.博士生预算
        self.doc_budget = self.get_doc_budget()
        self.doc_budget.to_excel(self.writer, sheet_name='7.博士生预算')

        # 保存，关闭数据库连接
        self.writer.save()
        self.conn.dispose()

    def init_df(self):
        # 筛选条件在这里改，目前只筛选学籍状态正常的全日制硕士
        sql = " SELECT DWBH, DWMC, SXLBMC, NJ, KSFS, XH, BKBYDW FROM GS WHERE XJZTMC='正常' AND SXLBMC = '全日制硕士研究生' "
        self.data = pd.read_sql(sql, self.conn)

    def stu_qlt(self, y1, y2, y3, y4, y5):
        sql = " SELECT * FROM gaoxiaoleibie "  # 这张表是自己创建的，在hustzz下，只有名称和类别两个字段，用于和本科毕业单位进行匹配
        df = pd.read_sql(sql, self.conn)

        # 转成字典减少进出df的次数，这一步在比较大的df中应该效果明显
        syzl = {}
        for name, type in zip(df['名称'], df['类别']):
            syzl[name] = type

        self.data['sylx'] = self.data['BKBYDW'].str.replace(" ", "").apply(lambda x: syzl[x] if x in syzl else "双非") + \
                            self.data['KSFS'].str.replace(" ", "").apply(
                                lambda x: "推免" if x == '推荐免试' else "统考")  # 生源类型，把211/985字段和推免/统考字段合并，避免多重标题
        self.data['sydf'] = self.data['sylx'].apply(Weight.syzldf,
                                                    args=(y1, y2, y3, y4, y5))  # 生源得分，apply的是下面的函数，就是去年的数字

    def get_z_score(self):
        """
        从分数人数表中获取学院年级得分，除以人数得出人均分数，用Z-Score公式计算
        """
        df = pd.DataFrame()
        for grade in ['2020', '2021', '2022']:
            df[grade + 'score'] = self.fsrs['sum'][grade] / self.fsrs['count'][
                grade]  # 分数除以人数,因为是从self.fsrs里取出来的series，自身是带索引的，就是带学院之类的信息
            df[grade + 'score'] = df[grade + 'score'].apply(
                lambda x: (x - df[grade + 'score'].mean()) / np.std(df[grade + 'score']))  # Z-Score公式
            df[grade + 'budget'] = df[grade + 'score'].astype('float') + 8  # 用Z-Score加上人均8千，得到初步预算
        return df

    def modify_basis(self):
        df = pd.DataFrame()
        for grade in ['2020', '2021', '2022']:
            # 总人数
            df[grade + '总人数'] = self.fsrs['count'][grade]
            # 推免人数
            if grade == '2022':
                df[grade + '推免人数'] = self.sylxtj['count'][grade]['211推免'] + self.sylxtj['count'][grade]['985推免'] + \
                                     self.sylxtj['count'][grade]['双非推免']
            else:
                df[grade + '推免人数'] = self.sylxtj['count'][grade]['211推免'] + self.sylxtj['count'][grade]['985推免']
            # 初步金额
            df[grade + '初步金额'] = df[grade + '总人数'] * self.z_score[grade + 'budget']
        return df

    def get_doc_budget(self):
        sql = """ 
            SELECT DWBH AS 院系代码, DWMC AS 院系名称, NJ AS 年级, 0 AS `直博生`, 0 AS `招生人数`, 0 AS `优博名额`, COUNT(*) AS `资助人数（人）/金额（万元）` FROM `GS` WHERE SXLBMC='全日制博士研究生' AND XJZTMC='正常' AND NJ IN ('2020', '2021', '2022') GROUP BY DWBH, DWMC, NJ 
            UNION ALL
            SELECT a.*, (a.`直博生`+a.`优博名额`) AS `资助人数（人）/金额（万元）` FROM(SELECT DWBH AS `院系代码`, DWMC AS `院系名称`, NJ AS `年级`, SUM(CASE WHEN XZ='5' AND KSFS='本科直博' AND XJZTMC='正常' THEN 1 ELSE 0 END) AS `直博生`, COUNT(*) AS `招生人数`, ROUND(COUNT(*)*0.3, 2) AS `优博名额` FROM `GS` WHERE SXLBMC='全日制博士研究生' AND NJ IN ('2018', '2019') GROUP BY DWBH, DWMC, NJ) AS a 
            UNION ALL
            SELECT b.*, IF(b.`直博生` >= b.`优博名额`, b.`优博名额`, b.`直博生`) AS `资助人数（人）/金额（万元）` FROM(SELECT DWBH AS `院系代码`, DWMC AS `院系名称`, NJ AS `年级`, SUM(CASE WHEN XZ='5' AND KSFS='本科直博' AND XJZTMC='正常' THEN 1 ELSE 0 END) AS `直博生`, COUNT(*) AS `招生人数`, ROUND(COUNT(*)*0.3, 2) AS `优博名额` FROM `GS` WHERE SXLBMC='全日制博士研究生' AND NJ ='2017' GROUP BY DWBH, DWMC, NJ) AS b 
            """
        df = pd.read_sql(sql, self.conn).pivot_table(index=['院系代码', '院系名称'], columns='年级', aggfunc=np.sum)
        return df


class Weight:
    @staticmethod
    def syzldf(x, y1, y2, y3, y4, y5):
        if x == '985推免':
            x = int(y1)
        elif x == '211推免':
            x = int(y2)
        elif x == '985统考':
            x = int(y3)
        elif x == '211统考':
            x = int(y4)
        elif x == '双非推免':
            x = int(y5)
        else:
            x = 0
        return x


class Modify:
    @staticmethod
    def amount_exhaustion(total_stu, rec_stu):
        """
        total_stu:某学院某年级受助总人数
        rec_stu:某学院某年级推免人数,新生包括双非推免，老生只考虑211以上推免
        """
        ae = []
        total_stu, rec_stu = int(total_stu), int(rec_stu)
        for i in range(total_stu + 1):
            for j in range(total_stu + 1):
                if i + j <= total_stu - rec_stu:  # 三等、二等相加不能挤占推免生的一等，这样也可以节约计算资源
                    for k in range(rec_stu, total_stu + 1):
                        if i + j + k == total_stu:
                            ae.append(i * 4 + j * 8 + k * 10)
        return np.array(ae)

    @staticmethod
    def get_closet(amount, total_stu, rec_stu):
        """
        amount_exhaustion方法穷举了一个学院某年级的受助金额的所有分配情况，本方法将某学院实际算出来的金额贴合到其中最近的一项
        amount_list: 一个学院某年级受助金额的可能的分配方案
        amount: 一个学院在Z-Score模型中算出来的值，或者说还未调整完的值
        """
        array = Modify.amount_exhaustion(total_stu, rec_stu)
        idx = (np.abs(array - amount)).argmin()  # 找到在穷举列中最接近输入金额的值的索引
        return array[idx]  # 用索引返回该值

    @staticmethod
    def modify(data, budget):
        df = pd.DataFrame()
        for grade in ['2020', '2021', '2022']:
            rate = data[grade + '初步金额'].sum() / float(budget[grade])  # 用当年的学院总初步金额除以预算，乘回金额
            # 第一次贴合之后把预算跟总预算对比一边
            df[grade + '预算'] = data[grade + '初步金额'].astype('float') / rate
            df[grade + '预算'] = data.apply(
                lambda x: Modify.get_closet(x[grade + '初步金额'], x[grade + '总人数'], x[grade + '推免人数']), axis=1)

            df[grade + '总人数'] = data[grade + '总人数']
            df[grade + '推免人数'] = data[grade + '推免人数']
            rate = df[grade + '预算'].sum() / float(budget[grade])
            # 第二次贴合效果很好
            df[grade + '预算'] = df[grade + '预算'] / rate
            df[grade + '预算'] = df.apply(
                lambda x: Modify.get_closet(x[grade + '预算'], x[grade + '总人数'], x[grade + '推免人数']), axis=1)
            df[grade + '预算'] = df[grade + '预算'] * 1000
            df[grade + '最终人均'] = df[grade + '预算'] / df[grade + '总人数']
        return df
