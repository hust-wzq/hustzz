import easygui as g
import rsa as r
import pandas as pd
import tkinter as tk
import os, sys, time, sqlalchemy, codecs, pymysql, webbrowser, docx


class ExTools(tk.Tk):
    def __init__(self):
        super(ExTools, self).__init__()

        # 标题和大小
        self.title('表单工具包')
        self.geometry('650x700')

        # 表单加密解密区
        lb1 = tk.Label(self, text='表单加密解密', font=('微软雅黑', 13))
        lb1.place(x=50, y=10)
        b1 = tk.Button(self, text='生成密钥对', width=15, height=2, command=self.create_rsa)
        b2 = tk.Button(self, text='加密表单内容', width=15, height=2, command=self.crypt_excel)
        b3 = tk.Button(self, text='解密表单内容', width=15, height=2, command=self.decrypt_excel)
        b1.place(x=50, y=55)
        b2.place(x=50, y=110)
        b3.place(x=50, y=165)

        # 数据库应用
        lb2 = tk.Label(self, text='数据库应用', font=('微软雅黑', 13))
        lb2.place(x=250, y=10)
        b4 = tk.Button(self, text='一键上传(非现存表）', width=15, height=2, command=self.one_key_upload)
        b5 = tk.Button(self, text='更新数据(需主键)', width=15, height=2, command=self.one_key_update)
        b6 = tk.Button(self, text='整张表下载', width=15, height=2, command=self.simple_download)
        b7 = tk.Button(self, text='SQL语句下载', width=15, height=2, command=self.sql_download)
        b4.place(x=250, y=55)
        b5.place(x=250, y=110)
        b6.place(x=250, y=165)
        b7.place(x=250, y=220)

        # Excel合并拆分
        lb3 = tk.Label(self, text='合并拆分', font=('微软雅黑', 13))
        lb3.place(x=450, y=10)
        b8 = tk.Button(self, text='表格合并', width=15, height=2, command=Table.tables_append)
        b9 = tk.Button(self, text='表格拆分', width=15, height=2, command=Table.tables_spilit)
        b10 = tk.Button(self, text='Word拆分', width=15, height=2, command=WordSplit)
        b8.place(x=450, y=55)
        b9.place(x=450, y=110)
        b10.place(x=450, y=165)

        # 帮助区
        b11 = tk.Button(self, text='使用说明书', width=15, height=2, command=self.users_manual)
        b11.place(x=450, y=220)

        # 底部输出区
        self.text = tk.Text(self, wrap='word')
        self.scr = tk.Scrollbar()  # 滚动条
        self.scr.pack(side=tk.RIGHT, fill=tk.Y)
        self.scr.config(command=self.text.yview)
        self.text.config(yscrollcommand=self.scr.set)
        self.text.place(x=10, y=300)
        self.text.tag_configure('stderr', foreground='#b22222')
        self.text.see('1000.0')
        # 用自己写的类篡改了命令行的输出，放到了界面里面
        sys.stdout = TextRedirector(self.text, 'stdout')
        sys.stderr = TextRedirector(self.text, 'stderr')

    @staticmethod
    def users_manual():
        webbrowser.open('https://s.alphalawyer.cn/1r1Cpe')

    @staticmethod
    def get_dirname(f: str) -> str:
        """
        读文件上级目录
        :param f: 文件名字符串
        :return: 上级目录字符串
        """
        return os.path.dirname(f) + '\\'  # 加上'\\'方便使用

    @staticmethod
    def read_xls(f: str):
        """
        读取excel或csv文件
        :param f: 文件名
        :return: DataFrame格式的数据表
        """
        return pd.read_excel(f, dtype=object) if '.xls' in f else pd.read_csv(f, dtype=object)

    @staticmethod
    def chose_cols(dataframe: any) -> list:
        """
        选择一列或多列，可以用于加密、下载等功能上
        :param dataframe: 被选择的表格
        :return: 被选取的列名
        """
        if len(dataframe.columns) > 1:  # 处理一列或多列的表格
            if g.ccbox('加解密一列或多列', choices=['单列', '多列']):
                return [g.choicebox('要加解密的列', choices=list(dataframe.columns))]  # 框在list中，同多选保持一致格式
            else:
                return g.multchoicebox('要加解密的列', choices=list(dataframe.columns))  # 多选返回的是list
        else:
            return [dataframe.columns[0]]  # 框在list中，同多选保持一致格式

    @staticmethod
    def read_pub_key():
        pub_key_file = g.fileopenbox('读取加密用的公钥')
        with open(pub_key_file, 'rb') as pub:
            p = pub.read()
        return r.PublicKey.load_pkcs1(p)

    @staticmethod
    def read_pri_key():
        pri_key_file = g.fileopenbox('读取解密用的私钥')
        with open(pri_key_file, 'rb') as pri:
            p = pri.read()
        return r.PrivateKey.load_pkcs1(p)

    @staticmethod
    def str2bytes(dt):
        """
        这个功能必不可少，用于转换字节符号中的两段反斜杠
        :param dt: dataframe中被加密的字符串，本应该是bytes但是用pd读取时变成了string
        :return: 复原成bytes
        """
        new_bytes = bytes(dt[2:-1], encoding='utf-8')
        original = codecs.escape_decode(new_bytes, 'hex_escape')
        return original[0]

    @staticmethod
    def crypt_cols(dt: any, pub_key: any) -> bytes:
        """
        加密数据
        :param pub_key: 公钥
        :param dt:被加密数据
        :return: 被加密后数据
        """
        dt = str(dt).encode('utf8')  # 将文本数据转换成字节，不然无法用rsa加密
        return r.encrypt(dt, pub_key)  # 返回加密后的数据

    @staticmethod
    def decrypt_cols(dt: any, pri_key: any) -> str:
        """
        加密数据
        :param pri_key: 读取私钥
        :param dt:被解密数据
        :return: 解密后数据
        """
        dt = r.decrypt(dt, pri_key)  # 使用私钥进行解密
        return str(dt, encoding='utf-8')  # 返回解密后的数据

    def crypt_excel(self):
        """
        加密表单中的一列或多列
        :return: 无
        """
        try:
            filename = g.fileopenbox('要加密的文件')
            dirpath = self.get_dirname(filename)
            df = self.read_xls(filename)
            cols_chosen = self.chose_cols(df)
            pub_key = self.read_pub_key()
            for col in cols_chosen:
                df[col] = df[col].apply(self.crypt_cols, args=(pub_key,))
            df.to_excel(dirpath + '加密后.xlsx', index=False, encoding='utf_8_sig')
            os.system('explorer.exe /n, {}'.format(dirpath))  # 输出完成后打开保存的文件夹
            print(time.asctime()[11:20] + '：已输出，并已为您自动打开文件夹')
        except TypeError:
            print(time.asctime()[11:20] + '：用户终止')

    def decrypt_excel(self):
        try:
            filename = g.fileopenbox('要解密的文件')
            dirpath = self.get_dirname(filename)
            df = self.read_xls(filename)
            cols_chosen = self.chose_cols(df)
            pri_key = self.read_pri_key()
            for col in cols_chosen:
                df[col] = df[col].apply(self.str2bytes)
                df[col] = df[col].apply(self.decrypt_cols, args=(pri_key,))
            df.to_excel(os.path.dirname(filename) + '\\解密后.xlsx', index=False, encoding='utf_8_sig')
            os.system('explorer.exe /n, {}'.format(dirpath))  # 输出完成后打开保存的文件夹
            print(time.asctime()[11:20] + '：已输出，并已为您自动打开文件夹')
        except TypeError:
            print(time.asctime()[11:20] + '：用户终止')

    @staticmethod
    def create_rsa():
        """
        创建一对rsa的公钥和私钥，其中公钥用于加密，私钥用于解密
        :return: 无
        """
        try:
            save_path = g.diropenbox('密码对保存位置') + '\\'
            pub_key, pri_key = r.newkeys(1024)

            pub = pub_key.save_pkcs1()
            pri = pri_key.save_pkcs1()

            with open(save_path + '公钥（加密）.pem', 'wb+') as f:
                f.write(pub)
            with open(save_path + '私钥（解密）.pem', 'wb+') as f:
                f.write(pri)
            os.system('explorer.exe /n, {}'.format(save_path))  # 输出完成后打开保存的文件夹
            print(time.asctime()[11:20] + '：已输出，并已为您自动打开文件夹')
        except TypeError:
            print(time.asctime()[11:20] + '：用户终止')

    @staticmethod
    def chose_db():
        engine = sqlalchemy.create_engine(
            'mysql+pymysql://yerdon:Qaz84759@sh-cynosdbmysql-grp-9v8niq3a.sql.tencentcdb.com:21749/sys')
        # 通过检测连接的方式来获取数据库的列表，等价于show databases;
        insp = sqlalchemy.inspect(engine)
        full_dbs = insp.get_schema_names()
        # 数据库中要把系统库去掉，系统库不能操作
        choices = [i for i in full_dbs if i not in ['information_schema', 'mysql', 'performance_schema', 'sys']]
        db = g.choicebox(msg='选择数据库', title='选择数据库', choices=choices)
        return sqlalchemy.create_engine(
            'mysql+pymysql://yerdon:Qaz84759@sh-cynosdbmysql-grp-9v8niq3a.sql.tencentcdb.com:21749/{}'.format(db))

    def one_key_upload(self):
        try:
            file = g.fileopenbox('打开要上传的表格')
            os.chdir(os.path.dirname(file))
            file = file.replace(os.path.dirname(file) + '\\', '')
            if '.xlsx' in file:
                df = pd.read_excel(file, dtype=object)
                file_name = file[:-5]
            elif '.xls' in file:
                df = pd.read_excel(file, dtype=object)
                file_name = file[:-4]
            elif '.csv' in file:
                df = pd.read_csv(file, dtype=object)
                file_name = file[:-3]
            engine = self.chose_db()
            df.to_sql(file_name, engine, index=False)
            print('操作成功，您写入名为《{}》的表格'.format(file_name))
        except ValueError:
            print(time.asctime()[11:20] + ':存在同名表，请先在数据库中操作')
        except TypeError:
            print(time.asctime()[11:20] + '：用户终止')

    def one_key_update(self):
        try:
            # 获取要上传的文件名、修改工作目录并去掉文件名后面的格式
            file = g.fileopenbox('打开要上传的表格')
            os.chdir(os.path.dirname(file))
            file = file.replace(os.path.dirname(file) + '\\', '')
            if '.xlsx' in file:
                df = pd.read_excel(file, dtype=object)
                file_name = file[:-5]
            elif '.xls' in file:
                df = pd.read_excel(file, dtype=object)
                file_name = file[:-4]
            elif '.csv' in file:
                df = pd.read_csv(file, dtype=object)
                file_name = file[:-3]
            engine = self.chose_db()    # 选取数据库
            insp = sqlalchemy.inspect(engine)   # 重置检查
            if len(insp.get_table_names()) > 0:
                choices = insp.get_table_names() + ['新建表格']
                tb_name = g.choicebox(msg='选择一张现存的表覆盖，或者新建一张表', title='选取表格', choices=choices)
            else:
                tb_name = '新建表格'

            if tb_name == '新建表格':
                tb_name = g.enterbox(msg='输入表格名称', title='表格名称', default=file_name)
                df.to_sql(tb_name, engine, index=False, if_exists='replace')
                print('操作成功，您写入名为《{}》的表格'.format(tb_name))
            else:
                if g.ccbox(msg='请务必确保这张表格已经设置主键'.format(tb_name), title='覆盖操作', choices=['继续', '取消']):
                    conn = engine.connect()
                    df.to_sql('temp', engine, index=False, if_exists='replace')
                    sql1 = " REPLACE INTO %s SELECT * FROM temp " % tb_name
                    conn.execute(sql1)
                    sql2 = " DROP Table If EXISTS temp "
                    conn.execute(sql2)
                    print(time.asctime()[11:20] + '：您写入了{}表格，如果原来存在这个表格，那么已经被替换，如需恢复，请联系管理员'.format(tb_name))
                else:
                    print(time.asctime()[11:20] + '：取消操作')
        except TypeError:
            print(time.asctime()[11:20] + '：用户终止')

    def simple_download(self):
        try:
            engine = self.chose_db()
            insp = sqlalchemy.inspect(engine)  # 重置检查
            if len(insp.get_table_names()) == 0:
                print(time.asctime()[11:20] + '：该数据库中没有表单')
            else:
                choices = insp.get_table_names()
                tb_name = g.choicebox(msg='要下载的表格', title='选取表格', choices=choices)
                sql = """ SELECT * FROM {} """.format(tb_name)
                df = pd.read_sql(sql, engine)
                if g.ccbox('是否预览前二十项？', choices=['预览', '直接下载']):
                    print(df.head(20))
                    if g.ccbox('是否下载？', choices=['下载', '取消']):
                        save_path = g.diropenbox('保存位置')
                        file_name = save_path + '\\' + tb_name + '.xlsx'
                        df.to_excel(file_name, index=False, encoding='utf_8_sig')
                        os.system('explorer.exe /n, {}'.format(save_path))
                        print(time.asctime()[11:20] + '：保存在{}文件夹'.format(save_path))
                    else:
                        print(time.asctime()[11:20] + '：用户取消下载')
                else:
                    save_path = g.diropenbox('保存位置')
                    file_name = save_path + '\\' + tb_name + '.xlsx'
                    df.to_excel(file_name, index=False, encoding='utf_8_sig')
                    os.system('explorer.exe /n, {}'.format(save_path))
                    print(time.asctime()[11:20] + '：保存在{}文件夹'.format(save_path))
        except TypeError:
            print(time.asctime()[11:20] + ':用户终止')
        except Exception:
            print(time.asctime()[11:20] + ':用户终止')

    def sql_download(self):
        try:
            engine = self.chose_db()
            sql = g.textbox('输入或粘贴查询语句(百分号查询和python冲突，会报错，换成LEFT或RIGHT）')
            df = pd.read_sql(sql, engine)
            df2 = pd.DataFrame([sql])
            if g.ccbox('是否预览前二十项？', choices=['预览', '直接下载']):
                print(df.head(20))
                if g.ccbox('是否下载？', choices=['下载', '取消']):
                    save_path = g.filesavebox('保存文件')
                    file_name = save_path + '.xlsx'
                    xl = pd.ExcelWriter(file_name)
                    df.to_excel(xl, sheet_name='查询结果', index=False, encoding='utf_8_sig')
                    df2.to_excel(xl, sheet_name='SQL语句', index=False, encoding='utf_8_sig')
                    xl.save()
                    print(time.asctime()[11:20] + '：保存为{}.xlsx'.format(save_path))
                else:
                    print(time.asctime()[11:20] + '：用户取消下载')
            else:
                save_path = g.filesavebox('保存文件')
                file_name = save_path + '.xlsx'
                xl = pd.ExcelWriter(file_name)
                df.to_excel(xl, sheet_name='查询结果', index=False, encoding='utf_8_sig')
                df2.to_excel(xl, sheet_name='SQL语句', index=False, encoding='utf_8_sig')
                xl.save()
                print(time.asctime()[11:20] + '：保存为{}.xlsx'.format(save_path))
        except TypeError:
            print(time.asctime()[11:20] + ':用户终止')

    @staticmethod
    def get_sheet_name(file):
        sheets = pd.read_excel(file, sheet_name=None)
        return list(sheets.keys())

    @staticmethod
    def mk_dir(path_s):
        if os.path.exists(path_s):  # 如果存在拆分文件夹
            print('拆分文件夹已存在，正在清理拆分文件夹')
            for i in os.listdir(path_s):  # 文件夹下每一个文件
                os.remove(path_s + i)  # 删除
        else:  # 如果存在拆分文件夹
            os.mkdir(path_s)  # 创建一个拆分文件夹
            print('正在创建拆分文件夹')


class TextRedirector(object):
    """
    篡改输出用的
    """
    def __init__(self, widget, tag='stdout'):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, str, (self.tag,))  # (self.tag,) 是设置配置
        self.widget.configure(state='disabled')


class Table:
    @staticmethod
    def tables_spilit():
        file = g.fileopenbox('要分割的excel')
        if file is None:
            print(time.asctime() + ':用户取消')
        else:
            try:
                dir_path = ExTools.get_dirname(file) + '\\拆分\\'      # 加上拆分文件夹
                ExTools.mk_dir(dir_path)                               # 清空、创建文件夹
                sheets = ExTools.get_sheet_name(file)                  # 获取sheets列表
                ref_list = []                                       # 因为不同sheet里面的学院可能不一致，我们需要一张全表
                default_value = [1, 1]
                for sheet in sheets:
                    print(time.asctime() + '：正在读取%s内容，请稍候' % sheet)
                    headers = g.multenterbox(msg='%s表格有几行表头，标题在哪' % sheet,
                                             fields=['表头行数', '标题在第几行'], values=default_value)
                    default_value = headers
                    headers_num = int(headers[0]) - 1
                    title_line = int(headers[1]) - 1

                    df = pd.read_excel(file, sheet_name=sheet, dtype=object, header=None)  # 读数据
                    df_head = df.loc[:headers_num, :]  # 表头部分
                    df = df.loc[headers_num + 1:, :]  # 内容部分

                    ref = g.choicebox('选择一列作为拆分依据', title='拆分依据', choices=list(df_head.loc[title_line, :]))
                    for i in range(len(list(df_head.loc[title_line, :]))):
                        if list(df_head.loc[title_line, :])[i] == ref:
                            df.rename(columns={i: ref}, inplace=True)
                            df_head.rename(columns={i: ref}, inplace=True)
                    """
                    for i in range(len(df[title])):
                        m = str(df[title][i]).replace('/', '')   # 去掉学院名称里不能用于命名文件的字符
                        df[title][i] = m.replace(' ', '')        # 去掉学院名称里的空格
                    """
                    df[ref] = df[ref].astype('str')     # 分类依据的格式统一设置为str，避免不同格式被分成不一样的文件
                    df_ref = df.drop_duplicates(subset=ref, keep='first', inplace=False)  # 学院名称去重

                    for reference in list(df_ref[ref]):     # 单sheet中的分类依据
                        df_temp = df.loc[df[ref] == reference]  # 按照依据检索
                        df_temp = df_temp.append(df_head)   # 加上表头部分
                        df_temp = df_temp.sort_index()      # 排序
                        df_temp.to_excel(dir_path + reference + '_' + sheet + '.xlsx', index=False, header=None)  # 输出到拆分文件夹中
                        if reference not in ref_list:       # 因为每张sheet中的分类依据可能彼此之间不是子集，需要取他们的并集
                            ref_list.append(reference)      # 历遍多张sheet后得到完整的依据表

                for reference in ref_list:      # 现在是完整的
                    writer = pd.ExcelWriter(dir_path + reference + '.xlsx')     # 创建一个可以重复写入的excel对象
                    for sheet in sheets:
                        print('已合并%s_%s内容' % (reference, sheet))
                        if os.path.exists(dir_path + reference + '_' + sheet + '.xlsx'):
                            df = pd.read_excel(dir_path + reference + '_' + sheet + '.xlsx', dtype=object)
                            df.to_excel(writer, sheet, index=False)
                            os.remove(dir_path + reference + '_' + sheet + '.xlsx')
                        else:
                            pass
                    writer.save()
                os.system('explorer.exe /n, {}'.format(ExTools.get_dirname(file)))
            except TypeError:
                print(time.asctime() + ':用户取消')

    @staticmethod
    def tables_append():
        try:
            files = g.fileopenbox(title='要合并的文件', multiple=True)
            dir_path = ExTools.get_dirname(files[0])
            if files is None:
                print(time.asctime() + ':未选择文件')
            elif len(files) == 1:
                print(time.asctime() + ':只选择了一个文件，合并失败')
            else:
                try:
                    writer = pd.ExcelWriter(dir_path + '\\合并结果.xlsx')
                    all_sheets = []     # 有些院系的sheet数量是不全的，得从所有文件中遍历一遍取得全集
                    print(time.asctime() + ':正在找sheets的全集')
                    for file in files:
                        sheets = ExTools.get_sheet_name(file)
                        for sheet in sheets:
                            if sheet not in all_sheets:
                                all_sheets.append(sheet)
                    print(time.asctime() + ':已找到sheets的全集，开始合并')

                    for sheet in all_sheets:    # 这样是完整的
                        headers = g.enterbox(msg='%s表格有几行表头' % sheet)
                        headers_num = int(headers) - 1

                        for i, file in enumerate(files):    # 先找一张有这个sheet的表格
                            if sheet in ExTools.get_sheet_name(file):
                                print(time.asctime() + ':已找到有{}sheet的第一个文件，开始取其表头'.format(sheet))
                                df = pd.read_excel(file, sheet_name=sheet, dtype=object)
                                print(time.asctime() + ':正从{}的{}取其表头'.format(file, sheet))
                                df_head = df.loc[:headers_num, :]  # 表头部分
                                df = df.loc[headers_num:, :]
                                break

                        for file in files[i +1:]:  # 剩余文件
                            try:
                                df_next = pd.read_excel(file, sheet_name=sheet, dtype=object)
                                df_next = df_next.loc[headers_num:, :]
                                df = df.append(df_next)  # 取每个表的内容部分
                            except ValueError:
                                print(time.asctime() + ':{}没有{}sheet，跳过'.format(file, sheet))

                        df = df_head.append(df)  # 如果有而且进行了合并，就加上表头
                        df.to_excel(writer, sheet, index=False)
                    writer.save()
                    os.system('explorer %s' % dir_path)
                    print(time.asctime() + ':合并完成,存储在{}'.format(dir_path))
                except PermissionError:
                    print(time.asctime() + '：文件被占用，请先关闭文件')
        except TypeError:
            print(time.asctime() + ':用户取消')


class WordSplit:
    def __init__(self):
        self.file = g.fileopenbox()
        old = docx.Document(self.file)
        dir_name = os.path.split(self.file)[0]
        self.save_path = dir_name + "\\拆分\\"
        ExTools.mk_dir(self.save_path)

        paras = old.paragraphs
        self.texts = [p.text for p in paras]
        title_index = int(g.enterbox("通知标题在第几行")) - 1
        self.school_index = int(g.enterbox("院系在第几行")) - 1
        self.title_text = self.texts[title_index]
        self.total = self.texts.count(self.title_text)
        self.paragraphs_per_file = int(len(self.texts) / self.total)
        self.split()

    @staticmethod
    def delete_paragraph(paragraph):
        p = paragraph._element
        p.getparent().remove(p)
        # p._p = p._element = None
        paragraph._p = paragraph._element = None

    def split(self):
        for i in range(self.total):
            new = docx.Document(self.file)
            """
            以法学院为例，法学院前面的都不要，假如前面有3个学院，每个学院15段，那就是3*15一共45段，在45段范围内逐次删除一段
            """
            for j in range(i * self.paragraphs_per_file):
                p = new.paragraphs[0]
                self.delete_paragraph(p)
            """
            还是以法学院为例，删完前面的，然后从十五段开始往后数的所有段落都不要,直接打个9999，反正超范围了就停
            """
            for j in range(self.paragraphs_per_file, 9999):
                try:
                    p = new.paragraphs[self.paragraphs_per_file]
                    self.delete_paragraph(p)
                except IndexError:
                    break
            new_name = self.texts[i * self.paragraphs_per_file + self.school_index]
            new_name = new_name.replace(" ", "").replace("：", "").replace("\n", "")
            print("正在拆分{}".format(new_name))
            new.save(self.save_path + "\\" + new_name + ".docx")
        print("拆分完毕")


if __name__ == '__main__':
    win = ExTools()
    win.mainloop()
