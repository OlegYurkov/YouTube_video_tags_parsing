

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import load_workbook

headers = {
    'authority': 'www.youtube.com',
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,'
              'application/signed-exchange;v=b3;q=0.9',
    'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 '
                  'Safari/537.36',
}

def get_keywords(url: str) -> list:
    keywords = []
    req = requests.get(url=url, headers=headers)
    soup = BeautifulSoup(req.text, 'lxml')
    tags = soup.find_all('meta', property="og:video:tag")
    for tag in tags:
        keywords.append(tag['content'])
    return keywords

def save_to_excel(keywords: list):
    try:
        # Спробуємо відкрити наявний файл Excel
        wb = load_workbook("keywords.xlsx")
        ws = wb.active
    except FileNotFoundError:
        # Якщо файл не знайдено, створимо новий
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Keywords"])  # Додаємо заголовок

    # Визначаємо наступний рядок для запису
    next_row = ws.max_row + 1

    # Записуємо ключові слова в комірки
    for keyword in keywords:
        cell = ws.cell(row=next_row, column=1, value=keyword)
        next_row += 1

    # Зберігаємо файл Excel
    wb.save("keywords.xlsx")
    print(f"Результати додані до файлу 'keywords.xlsx'")

def main():
    url = input('Введіть посилання на відео: ')
    if "youtube.com" not in url or 'playlists' in url:
        print('Введіть правильне посилання')
        return
    keywords = get_keywords(url)
    if not keywords:
        print('\nУ цьому відео немає ключових слів')
    else:
        print(f'\nЗнайдені ключові слова:\n{"-"*23}')
        for keyword in keywords:
            print(keyword)
        save_to_excel(keywords)

if __name__ == "__main__":
    ma