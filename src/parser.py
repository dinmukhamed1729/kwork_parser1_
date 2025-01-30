import os
import re
import logging
from urllib.parse import urlsplit, urljoin, unquote

import pandas as pd
import requests
from bs4 import BeautifulSoup


class Parser:
    def __init__(self, session, save_path="./"):
        """Конструктор парсера."""
        self.session = session
        self.save_path = save_path
        self.base_url = "https://ecu-firmware-files.ru"
        self.title = None
        if not os.path.exists(save_path):
            os.makedirs(save_path)

    def get_page_content(self, url):
        """Запрашивает HTML-код страницы с обработкой ошибок."""
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            return response.text
        except requests.exceptions.Timeout:
            logging.error(f"⏳ Таймаут при загрузке {url}")
        except requests.exceptions.RequestException as e:
            logging.error(f"❌ Ошибка сети {url}: {e}")
        return None

    def find_thread_links(self, html):
        """Находит ссылки на темы."""
        soup = BeautifulSoup(html, "html.parser")
        return [
            self.base_url + a["href"].split("/unread")[0]
            for div in soup.find_all("div", class_="structItem-title")
            if (a := div.find("a", href=True))
        ]

    def extract_thread_id(self, thread_url):
        """Извлекает ID темы из URL."""
        match = re.search(r"/threads/([^/]+)(?:/|$)", thread_url)
        return match.group(1) if match else None

    def extract_articles(self, html):
        """Извлекает статьи из темы."""


        soup = BeautifulSoup(html, "html.parser")
        self.title = soup.find("h1", class_="p-title-value")
        container = soup.find("div", class_="block-body js-replyNewMessageContainer")

        if not container:
            return []

        return container.find_all("article", class_="message")

    def parse_thread(self, thread_url):
        """Парсит страницу темы и извлекает данные."""
        html = self.get_page_content(thread_url)
        if not html:
            return None
        articles = self.extract_articles(html)

        titles, authors, texts = [], [], []

        for article in articles:

            author = article.find("a", class_="username")
            description = article.find("div", class_="bbCodeBlock-expandContent") or article.find("div", class_="bbWrapper")


            author = author.get_text(strip=True) if author else "Автор не найден"
            text = description.get_text("\n", strip=True) if description else "Текст не найден"

            authors.append(author)
            texts.append(text)

        title = self.title.get_text(strip=True) if self.title else "Заголовок не найден"
        # Сбор ссылок на вложенные файлы

        soup = BeautifulSoup(html, "html.parser")
        attachments = [
            {"name": link["title"], "url": link["href"]}
            for link in soup.select("ul.attachmentList a[href]")
            if link.get("title")
        ]

        # Поиск кнопки "Скачать"
        download_button = soup.select_one(".p-title-pageAction a.button--cta")
        if download_button and (download_url := download_button.get("href")):
            attachments.append({"name": "", "url": download_url})

        title_tag = soup.select_one(".p-title .p-title-value")
        dir_name = title_tag.text.strip() if title_tag else ""

        return {
            "title": title,
            "author": authors,
            "br_text": texts,
            "attachments": attachments,
            "dir_name": dir_name,
            "thread_url": thread_url,
        }

    def parse_forum(self, forum_url):
        """Парсит все страницы форума и сохраняет данные."""
        page_number = 1
        results = []

        while forum_url:
            html = self.get_page_content(forum_url)
            if not html:
                logging.error(f"Не удалось загрузить {forum_url}.")
                break

            thread_links = self.find_thread_links(html)
            for thread_url in thread_links:
                if thread_data := self.parse_thread(thread_url):
                    results.append(thread_data)
                    self.save_thread_data(thread_data)

            # Переход на следующую страницу
            soup = BeautifulSoup(html, "html.parser")
            next_page = soup.select_one("a.pageNav-jump--next")
            forum_url = self.base_url + next_page["href"] if next_page else None
            page_number += 1

        print(f"Парсинг завершен. Обработано {page_number - 1} страниц, найдено {len(results)} тем.")
        return results

    def save_thread_data(self, data):
        """Сохраняет данные темы: файлы, текст, отчет в Excel."""
        thread_folder = os.path.join(self.save_path, data["dir_name"])
        os.makedirs(thread_folder, exist_ok=True)

        # 1️⃣ Сохранение текстового файла с описанием
        self.save_text_file(data, thread_folder)

        # 2️⃣ Поиск и скачивание вложенных файлов
        self.download_attachments(data, thread_folder)

        # 3️⃣ Добавление в `report.xlsx`
        self.update_report(data)

    def save_text_file(self, data, thread_folder):
        """Сохраняет текстовый файл с описанием темы."""
        txt_path = os.path.join(thread_folder, f"{data['dir_name']}.txt")
        with open(txt_path, "a", encoding="utf-8") as f:
            f.write(f"Ссылка на тему: {data['thread_url']}\n")
            for author, text in zip(data["author"], data["br_text"]):
                f.write(f"Название: {data['title']}\n")
                f.write(f"Автор: {author}\n")
                f.write(f"Описание:\n{text}\n")
                f.write("\n" * 5)

        print(f"✅ Текстовый файл с описанием сохранен: {txt_path}")

    def download_attachments(self, data, thread_folder):
        """Загружает все вложенные файлы в тему."""
        for attachment in data["attachments"]:
            file_url = urljoin(self.base_url, attachment["url"])
            file_name = attachment["name"]

            try:
                self.download_file(file_url, thread_folder, file_name)
            except requests.RequestException as e:
                logging.error(f"❌ Ошибка скачивания {file_url}: {e}")

    def download_file(self, file_url, thread_folder, file_name=""):
        """Скачивает файл по ссылке и сохраняет его на диск."""
        response = self.session.get(file_url, stream=True)
        response.raise_for_status()

        # Проверка Content-Type и Content-Disposition для различия файлов и HTML-страниц
        content_type = response.headers.get("Content-Type", "")

        if "text/html" in content_type:
            print("+++++++++++++++++++++++++++++++")
            file_info = self.get_file_info(file_url)
            print(file_info)
            for file in file_info:
                self.download_file(file["file_url"], thread_folder, file["file_name"])
            return

        # Получаем корректное имя файла
        file_name = self.get_filename_from_headers(response.headers, file_url, file_name)

        if file_name == "reply":
            print(f"⚠️ Пропущен файл с именем 'reply': {file_url}")
            return

        file_path = os.path.join(thread_folder, file_name)
        file_path = self.ensure_unique_file_path(file_path)

        with open(file_path, "wb") as f:
            for chunk in response.iter_content(1024):
                f.write(chunk)

        print(f"✅ Файл {file_name} скачан и сохранен в {file_path}")

    def get_filename_from_headers(self, headers, file_url, file_name):
        """Извлекает имя файла из заголовков или URL, поддерживая кириллицу."""
        content_disposition = headers.get("Content-Disposition")

        if not file_name:
            # Проверяем `filename*` с UTF-8 кодировкой
            if content_disposition and "filename*" in content_disposition:
                file_name_encoded = content_disposition.split("filename*=")[1]
                encoding, file_name = file_name_encoded.split("''", 1)
                file_name = unquote(file_name)

            # Проверяем `filename`
            elif content_disposition and "filename=" in content_disposition:
                file_name = content_disposition.split("filename=")[1].strip('"')

            # Если нет заголовка, берем имя из URL
            else:
                file_name = os.path.basename(urlsplit(file_url).path)
                file_name = unquote(file_name)  # Декодируем кириллицу

        return file_name

    def ensure_unique_file_path(self, file_path):
        """Проверяет наличие файла с таким же именем и добавляет номер, если файл уже существует."""
        counter = 1
        original_file_path = file_path
        while os.path.exists(file_path):
            name, ext = os.path.splitext(original_file_path)
            file_path = os.path.join(os.path.dirname(original_file_path), f"{name} - {counter}{ext}")
            counter += 1
        return file_path

    def update_report(self, data):
        """Обновляет отчет в Excel."""
        excel_path = os.path.join(self.save_path, "report.xlsx")

        df = pd.read_excel(excel_path) if os.path.exists(excel_path) else pd.DataFrame(
            columns=["№", "Статус", "Название темы", "Ссылка на тему"]
        )

        if data["thread_url"] not in df["Ссылка на тему"].values:
            df = pd.concat([df, pd.DataFrame([{
                "№": len(df) + 1,
                "Статус": "Скачан",
                "Название темы": data["title"],
                "Ссылка на тему": data["thread_url"]
            }])], ignore_index=True)

        df.to_excel(excel_path, index=False)
        print(f"📊 Отчет обновлен: {excel_path}")

    def get_file_info(self, page_url):
        """Извлекает имя файла и ссылку для скачивания с HTML страницы."""
        try:
            response = self.session.get(page_url)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")
            file_info = []

            file_items = soup.select(".block-body .block-row")

            for item in file_items:
                file_name_tag = item.select_one(".contentRow-title")
                file_link_tag = item.select_one(".contentRow-extra a")

                if file_name_tag and file_link_tag:
                    file_name = file_name_tag.text.strip()
                    file_url = urljoin(self.base_url, file_link_tag["href"])

                    file_info.append({"file_name": file_name, "file_url": file_url})

            return file_info

        except requests.RequestException as e:
            print(f"Ошибка при получении страницы: {e}")
            return []

#std_28097344_9AWAEU41_combiloader.bin