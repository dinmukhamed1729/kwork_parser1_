import tkinter as tk
from tkinter import messagebox, filedialog
import logging
from parser import Parser
from tkinter.ttk import Progressbar
import openpyxl

class UrlInputPage:
    def __init__(self, root, session):
        self.result_text = None
        self.path_entry = None
        self.url_entry = None
        self.progress = None
        self.root = root
        self.session = session

        # Поле для каталога скачивания
        self.save_path = tk.StringVar()

        # Создаем объект парсера и передаем путь
        self.parser = Parser(session)

        self.create_ui()

    def create_ui(self):
        tk.Label(self.root, text="Введите URL форума:").grid(row=0, column=0, padx=10, pady=5)
        self.url_entry = tk.Entry(self.root, width=50)
        self.url_entry.grid(row=0, column=1, padx=10, pady=5)
        self.progress = Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, pady=10)
        # Поле выбора каталога
        tk.Label(self.root, text="Каталог для скачивания:").grid(row=1, column=0, padx=10, pady=5)
        self.path_entry = tk.Entry(self.root, textvariable=self.save_path, width=40)
        self.path_entry.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(self.root, text="Выбрать", command=self.select_directory).grid(row=1, column=2, padx=5, pady=5)

        # Кнопки управления
        tk.Button(self.root, text="Парсить форум", command=self.parse_forum).grid(row=2, column=0, columnspan=3,
                                                                                  pady=10)

        # Поле для вывода результатов
        self.result_text = tk.Text(self.root, height=10, width=70)
        self.result_text.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

    def update_progress(self, value):
        self.progress["value"] = value
        self.root.update_idletasks()

    def select_directory(self):
        """Выбор каталога для сохранения скачанных файлов."""
        directory = filedialog.askdirectory()
        if directory:
            self.save_path.set(directory)  # Устанавливаем путь
            self.parser.save_path = directory  # Передаем в парсер
            logging.info(f"Каталог сохранения установлен: {directory}")

    def parse_forum(self):
        """Запускает парсинг форума."""
        url = self.url_entry.get()
        save_path = self.save_path.get()

        if not url:
            messagebox.showwarning("Ошибка", "Введите URL форума!")
            return

        if not save_path:
            messagebox.showwarning("Ошибка", "Выберите каталог для сохранения файлов!")
            return

        logging.info(f"Запуск парсинга форума: {url}")
        results = self.parser.parse_forum(url)

        if results:
            self.result_text.delete("1.0", tk.END)
            for thread in results:
                self.result_text.insert(tk.END, f"{thread['title']} - {thread['author']}\n")

            messagebox.showinfo("Успех", f"Найдено {len(results)} тем.")
        else:
            messagebox.showwarning("Ошибка", "Темы не найдены.")
