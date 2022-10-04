import pyautogui as pag
import pandas as pd
from time import sleep
import pyperclip as ppc


file = r"E:\助管文件夹\代码克隆\hustzz\事务流水\新建 Microsoft Excel 工作表.xlsx"
df = pd.read_excel(file)

for sid in df['XH']:
    pag.moveTo()
    pag.click()
    sleep(.2)
    pag.doubleClick()
    ppc.copy(sid)
    pag.hotkey('ctrl', 'a')
    sleep(.2)
    pag.hotkey('ctrl', 'v')
    sleep(.2)
    pag.press('enter')
    pag.moveTo()
    pag.click()
    pag.moveTo()
    pag.click()
    pag.moveTo()
    pag.click()
    sleep(.2)