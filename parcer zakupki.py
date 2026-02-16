from tkinter import *
from tkinter import ttk
import tkinter as tk 
from tkinter.messagebox import showerror, showwarning, showinfo
from tkinter import font

import datetime
import os
import requests
from lxml import html
import time
from fake_useragent import UserAgent
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import column_index_from_string
from enum import Enum
import webbrowser
import subprocess

class ButtonStatus(Enum):
    Start = 1
    End = 2

g_szVersion = '0.3.0'
g_szTitleName = 'Parce zakupki.gov.ru'
g_szExelPathRead = None
g_Button = None
g_ButtonStatus = ButtonStatus.Start

g_EntryName = None
g_EntryTableName = None
g_EntryLabel = None

g_Root = None

def GUI_ButtonClick():
    global g_Button
    global g_ButtonStatus
    global g_EntryName
    global g_EntryTableName

    if g_ButtonStatus == ButtonStatus.Start:
        g_Button["text"] = "Идет поиск"

        ParceUrl = g_EntryName.get()
        TableName = g_EntryTableName.get()

        g_ButtonStatus = ButtonStatus.End
        Parce_Start(ParceUrl, TableName)

def cmd_paste(self):
    global g_Root
    widget = g_Root.focus_get()
    if isinstance(widget, ttk.Entry) or isinstance(widget, tk.Text):
        widget.event_generate("<<Paste>>")

def cmd_cut(self):
    global g_Root
    widget = g_Root.focus_get()
    if isinstance(widget, ttk.Entry) or isinstance(widget, tk.Text):
        widget.event_generate("<<Cut>>")

def cmd_copy(self):
    global g_Root
    widget = g_Root.focus_get()
    if isinstance(widget, ttk.Entry) or isinstance(widget, tk.Text):
        widget.event_generate("<<Copy>>")

def cmd_selectall(self):
    self.widget.select_range(0, 'end')
    self.widget.icursor('end')

def GUI_Click_Text(URL):
    webbrowser.open(URL)

def GUI_OpenResultFolder():
    g_szPath_Script = os.getcwd()
    global g_szExelPathRead
    g_szExelPathRead = g_szPath_Script + '\\Результаты поиска'
    if not os.path.exists(g_szExelPathRead):
        os.mkdir(g_szExelPathRead)
    subprocess.Popen(r'explorer ' + g_szExelPathRead)

def GUI_KeyBind(e):
    if e.keycode == 86 and e.keysym != 'v':
        cmd_paste(e)
    elif e.keycode == 67 and e.keysym != 'c':
        cmd_copy(e)
    elif e.keycode == 88 and e.keysym != 'x':
        cmd_cut(e)
    elif e.keycode == 65 and e.keysym != 'a':
        cmd_selectall(e)

def GUI_Start():
    global g_Root
    g_Root = Tk()
    g_Root.resizable(False, False)

    global g_szTitleName
    g_Root.title(g_szTitleName)
    g_Root.geometry("300x150")

    HyperLinkFontStyle = font.Font(family = "Arial", size = 12, underline = True)
    text = tk.Label(g_Root, fg="blue", text="Ссылка на поиск с фильтрами", cursor="hand2", font=HyperLinkFontStyle)
    text.bind("<Button-1>", lambda e: GUI_Click_Text('https://zakupki.gov.ru/epz/contract'))
    text.pack(anchor=CENTER)

    global g_EntryName
    g_EntryName = ttk.Entry()
    g_EntryName.pack(anchor=CENTER, pady=1, fill=X)


    text = tk.Label(g_Root, fg="blue", text="Название таблицы", cursor="hand2", font=HyperLinkFontStyle)
    text.bind("<Button-1>", lambda e: GUI_OpenResultFolder())
    text.pack(anchor=CENTER)

    global g_EntryTableName
    g_EntryTableName = ttk.Entry()
    g_EntryTableName.pack(anchor=CENTER, pady=1, fill=X)

    global g_EntryLabel
    g_EntryLabel = ttk.Label()
    g_EntryLabel.pack(anchor=CENTER, pady=1)

    global g_Button
    g_Button = Button(text="Начать поиск", command=GUI_ButtonClick)
    g_Button.pack(anchor=CENTER, pady=1)

    text = tk.Label(g_Root, text=f'v {g_szVersion}', font=("Arial", 10)) 
    text.place(x=250, y=120)

    g_EntryTableName.bind('<Control-KeyPress>', GUI_KeyBind)
    g_EntryName.bind('<Control-KeyPress>', GUI_KeyBind)
    g_Root.mainloop()

def GetInfoByID(ID):
    try:
        url = 'https://zakupki.gov.ru/epz/contract/contractCard/common-info.html?reestrNumber=' + str(ID)
        szRequest = requests.get(url, headers={'User-Agent': UserAgent().chrome})
        print('Try parce: ' + url)
        if (szRequest.status_code == 429 or szRequest.status_code == 404):
            print('Error request: ' + str(szRequest.status_code))
            return 0

        # with open('test1.html', 'w', encoding='utf-8') as output_file:
        #     output_file.write(szRequest.text)
        
        tree = html.fromstring(szRequest.content)

        preKey = '//h2[normalize-space(text())="Информация о поставщиках"]/..//'
        name = IsKeyCopy(preKey + 'td[@class="tableBlock__col tableBlock__col_first text-break"]/text()', tree)
        if (len(name) > 0):
            name = name[0].strip()
    
        inn = IsKeyCopy(preKey + 'td[@class="tableBlock__col tableBlock__col_first text-break"]//section[span[contains(text(),"ИНН:")]]/span[2]/text()', tree)
        if (len(inn ) > 0):
            inn = inn[0].strip()

        address = IsKeyCopy(preKey + 'td[3]/text()', tree)
        if (len(address) > 0):
            address = address[0].strip()

        contact_a = ''
        contact_b = ''

        contact_a = IsKeyCopy(preKey + 'td[5]/text()', tree)
        if (len(contact_a) > 0 and (contact_a[0].strip()).find('субъект') != -1):
            contact_a = []
        if (len(contact_a) < 1 or not (contact_a[0].strip())):
            contact_a = IsKeyCopy(preKey + 'td[4]/text()', tree)
        if (len(contact_a) > 0 and (contact_a[0].strip())):
            texts = [t.strip() for t in contact_a if t.strip()]
            contact_a = texts[0]
            if (len(texts) > 1):
                contact_b = texts[1]
        else:
            contact_a = ''
        return {'name': name, 
                'inn': inn,
                'address': address,
                'contact_a': contact_a,
                'contact_b': contact_b,
                'url': url}
    except Exception as szError:
        print(str(szError)) 
        return 0

def IsKeyCopy(key, btree):
    vkey = btree.xpath(key)
    if (len(vkey) == 0):
        return ''
    return vkey

def GetPagesCount(URL):
    try: 
        szRequest = requests.get(URL, headers={'User-Agent': UserAgent().chrome})
        tree = html.fromstring(szRequest.content)
        iPagesCount = tree.xpath('//li[contains(@class, "page")][last()]/a/span[@class="link-text"]/text()')
        if (iPagesCount):
            if (len(iPagesCount) == 1):
                return int(iPagesCount[0])
    except Exception:
            return 1
    return 1

def ToExcel(Data, TableName):
    DT = pd.DataFrame(data=Data)
    DT.columns = ['Название', 'ИНН', 'Адресс', 'Контакт 1', 'Контакт 2', 'URL']

    aNumbers = []
    for i in range(1, len(Data)+1):
        aNumbers.append(i)
    DT.insert(0, '№', aNumbers)
    
    g_szPath_Script = os.getcwd()
    
    global g_szExelPathRead
    g_szExelPathRead = g_szPath_Script + '\\Результаты поиска'
    if not os.path.exists(g_szExelPathRead):
        os.mkdir(g_szExelPathRead)
    g_szExelPathRead += '\\' + str(TableName) + '.xlsx'
    
    DT.to_excel(g_szExelPathRead, index=False)
    
    wb = load_workbook(g_szExelPathRead)
    ws = wb.active

    adjusted_width = [0, 4, 100, 12, 84, 16, 24, 10]
    thin = Side(border_style="thin", color="000000")
    for column_cells in ws.columns:
        column = column_cells[0].column_letter

        for cell in column_cells:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        ws.column_dimensions[column].width = adjusted_width[column_index_from_string(column)]
    wb.save(g_szExelPathRead)

    print(f'Save file to {g_szExelPathRead}')

def ResetSearch():
    global g_ButtonStatus
    g_ButtonStatus = ButtonStatus.Start
    g_Button["text"] = "Начать поиск"

def Parce_Start(URL, TableName):
    try:
    
        global g_szTitleName

        if (URL.find('https://zakupki.gov.ru/epz/contract/') == -1):
            showerror(title=g_szTitleName, message="Не правильная ссылка")
            ResetSearch()
            return

        if (TableName == ''):
            TableName = str(datetime.datetime.now().strftime("%Y.%m.%d %H_%M_%S"))

        URL = re.sub(r'(recordsPerPage=)_\d+', r'\1_100', URL)
        URL = re.sub(r'(pageNumber=)\d+', r'\g<1>1', URL)

        iMaxPage = GetPagesCount(URL)
        if (iMaxPage < 1):
            showinfo(title=g_szTitleName, message="Не нашел страниц!")
            ResetSearch()
            return

        IDs = []
        DATA = []

        iCount = 1
        if (iMaxPage == 1):
            szRequest = requests.get(URL, headers={'User-Agent': UserAgent().chrome})

            tree = html.fromstring(szRequest.content)
            titles = tree.xpath('//div[contains(@class, "registry-entry__header-mid__number")]/*/text()')
            IDs = []
            for tile in titles:
                text = tile.strip()
                if (text[0] != '№'):
                    continue
                text = text[2:]
                IDs.append(text)

            for ID in IDs:
                time.sleep(0.1)
                ID = str(ID)

                data = GetInfoByID(ID)
                if (data == 0):
                    continue
                DATA.append(data)
                iCount += 1
        else:
            for i in range(1, iMaxPage):
                URLp = URL.replace('pageNumber=1', 'pageNumber=' + str(i))

                szRequest = requests.get(URLp, headers={'User-Agent': UserAgent().chrome})

                tree = html.fromstring(szRequest.content)
                titles = tree.xpath('//div[contains(@class, "registry-entry__header-mid__number")]/*/text()')
                IDs = []
                for tile in titles:
                    text = tile.strip()
                    if (text[0] != '№'):
                        continue
                    text = text[2:]
                    IDs.append(text)

                URL_Count = 1
                for ID in IDs:
                    time.sleep(0.1)
                    ID = str(ID)
                    data = GetInfoByID(ID)
                    print(f'Page {i}/{iMaxPage-1} URL {URL_Count}/{len(IDs)}')
                    URL_Count += 1
                    if (data == 0):
                        continue
                    DATA.append(data)
        if (len(DATA) <= 1):
            showinfo(title=g_szTitleName, message="Поиск завершен, но не успешно!")
            ResetSearch()
            return
        
        seen = set()
        subDATA = []
        for ID in DATA:
            if ID['inn'] not in seen:
                if ID['inn'] != "":
                    seen.add(ID['inn'])
                subDATA.append(ID)
        DATA = subDATA                

        ToExcel(DATA, TableName)

        global g_szExelPathRead
        showinfo(title=g_szTitleName, message="Поиск успешно завершен!\nРезультаты лежат тут\n" + g_szExelPathRead)
        ResetSearch()

    except Exception as szError:
        showerror(title=g_szTitleName, message="Ошибка " + str(szError))
        ResetSearch()
        return

def Main():
    if (int(datetime.datetime.now().strftime("%Y%m%d%H%M%S")) > 20260312084754):
        showerror(title=g_szTitleName, message="Лицензия не действительна, " \
        "обратись к сис. админу для продления!")
        return
    # Data = []
    # Data.append(GetInfoByID(3861200531325000057))
    # ToExcel(Data, '123')
    GUI_Start()

if __name__ == "__main__":
    Main()