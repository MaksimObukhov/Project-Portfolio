import pyautogui as pg
import os
import time


folder_path = input('Enter the path to the working folder: ')
month = input('Enter the month of the forecast: ')

files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]

delay = [7, 10, 5, 5]
for file in files:
    os.startfile(f'{folder_path}\{file}')
    time.sleep(7)

    Vert = 750
    for i in range(4):
        pg.click(127, 373, clicks=2, interval=0.5)
        pg.moveTo(1886, 750, 0.2)
        pg.click(clicks=2, interval=0.5)

        if i == 3:
            Vert = 677

        pg.moveTo(Vert, 261, 0.2)
        pg.doubleClick()
        pg.typewrite(month)

        pg.moveTo(1246, 824, 0.2)
        pg.click()

        time.sleep(delay[i])

        if i == 3:
            pg.hotkey('ctrl', 'pgdn')
            pg.hotkey('ctrl', 'pgdn')
            break

        pg.hotkey('ctrl', 'pgdn')
        pg.moveTo(127, 373, 0.2)
        pg.click(clicks=2, interval=0.5)

    pg.hotkey('ctrl', 's')
    time.sleep(3)
    pg.hotkey('alt', 'f4')





