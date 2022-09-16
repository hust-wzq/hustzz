from time import sleep
import pandas as pd     # 用于读表
import pyautogui as p     # 图形自动化
import pyperclip     # 用于剪切板交互


sleep(5)        # 缓冲时间
df = pd.DataFrame(columns=['data'])
for i in df['data']:
    pyperclip.copy(i)       # 把i传入剪切板
    p.moveTo(960, 540, 0.5)     # 横纵坐标+移动时间
    sleep(.2)                       # 给页面加载等事项的预留时间
    p.click()                       # 模拟点击
    p.hotkey('ctrl', 'a')           # 模拟热键
    p.hotkey('ctrl', 'v')           # 跟剪切板交互，模拟粘贴
    p.press('enter')                # 跟键盘单一按键交互
    

def func(x):
    name = x['name']


df.apply(func, axis=1)