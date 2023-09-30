# pip install beautifulsoup4
# pip install openpyxl
# pip install pytube

import openpyxl
from pytube import Playlist, YouTube
from bs4 import BeautifulSoup

def get_video_links_from_playlist(playlist_url, output_file):
    try:
        # Створюємо об'єкт Playlist і завантажуємо інформацію про плейлист
        playlist = Playlist(playlist_url)

        # Отримуємо список посилань на відео у плейлисті
        video_links = playlist.video_urls

        # Зберігаємо посилання у файл
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Video Links"
            ws.append(["Video Links"])

            for link in video_links:
                ws.append([link])

            wb.save(output_file)
            print(f"Посилання на відео збережено у файл '{output_file}'")
        except Exception as e:
            print(f"Виникла помилка при збереженні у файл: {str(e)}")

        return video_links
    except Exception as e:
        print(f"Сталася помилка при отриманні посилань на відео: {str(e)}")
        return []

def get_video_tags(video_url, output_file):
    try:
        yt = YouTube(video_url)
        watch_html = yt.watch_html
        soup = BeautifulSoup(watch_html, 'html.parser')
        keywords_tag = soup.find('meta', {'name': 'keywords'})
        if keywords_tag:
            keywords = keywords_tag['content'].split(',')
            return [tag.strip() for tag in keywords]
        else:
            return []
    except Exception as e:
        print(f"Відбулася помилка при отриманні тегів: {str(e)}")
        return []

def remove_duplicate_tags(tags_list):
    return list(set(tags_list))

def process_video_links(playlist_url, input_file, output_file):
    video_links = get_video_links_from_playlist(playlist_url, input_file)

    if video_links:
        tags_list = []

        for link in video_links:
            tags = get_video_tags(link, output_file)
            tags_list.extend(tags)

        # Видаляємо теги, що повторюються
        tags_list = remove_duplicate_tags(tags_list)

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Keywords"
            ws.append(["Keywords"])

            for tag in tags_list:
                ws.append([tag])

            wb.save(output_file)
            print(f"Мітки збережені у файл '{output_file}'")
        except Exception as e:
            print(f"Виникла помилка при збереженні у файл: {str(e)}")

def main():
    playlist_url = input('Виникла помилка при збереженні файлу: ')
    input_file = "video_links.xlsx"
    output_file = "keywords.xlsx"

    process_video_links(playlist_url, input_file, output_file)

if __name__ == "__main__":
    main()