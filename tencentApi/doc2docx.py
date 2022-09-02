from win32com import client as wc
import os


class TransDocToDocx:
    def __init__(self):
        self.word = wc.Dispatch('Word.Application')

    def trans(self, file):
        # 打开旧word 文件
        old_name = file
        doc = self.word.Documents.Open(old_name)
        # 保存为新word 文件,其中参数 12 表示的是docx文件
        rename = os.path.splitext(old_name)[0] + '.docx'
        doc.SaveAs(rename, 12)
        # 关闭word文档
        doc.Close()
        self.word.Quit()
        return rename
