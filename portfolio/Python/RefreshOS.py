import pyautogui as pg
import os
import time


folder_path = input('Before running this program be sure you are LoggedOn to OneStream\n\n'
                    'Enter the path to the working folder: ')

files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]

for file in files:
    os.startfile(f'{folder_path}\{file}')
    time.sleep(10)
    pg.moveTo(750, 500, 0.2)
    pg.click(clicks=2, interval=0.5)
    with pg.hold('alt'):
        pg.press(['x', 'f', 'y', '2'])
    time.sleep(15)
    pg.hotkey('ctrl', 's')
    time.sleep(3)
    pg.hotkey('alt', 'f4')


