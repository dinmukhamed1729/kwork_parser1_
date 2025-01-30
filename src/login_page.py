import tkinter as tk
from tkinter import messagebox
import requests
import logging


class LoginPage:
    def __init__(self, root, on_success):
        self.root = root
        self.on_success = on_success
        self.session = requests.Session()

        self.create_ui()

    def create_ui(self):
        tk.Label(self.root, text="Логин:").grid(row=0, column=0, padx=10, pady=10)
        self.login_entry = tk.Entry(self.root)
        self.login_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.root, text="Пароль:").grid(row=1, column=0, padx=10, pady=10)
        self.password_entry = tk.Entry(self.root, show="*")
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Button(self.root, text="Войти", command=self.login).grid(row=2, column=0, columnspan=2, pady=10)

    def login(self):
        login = self.login_entry.get()
        password = self.password_entry.get()

        logging.info("Начинаем процесс входа...")

        login_url = "https://ecu-firmware-files.ru/login"
        post_url = "https://ecu-firmware-files.ru/login/login"

        payload = {"login": login, "password": password, "remember": "1"}
        headers = {"User-Agent": "Mozilla/5.0"}

        try:
            response = self.session.get(login_url, headers=headers)
            logging.info(f"Получен ответ от страницы входа: {response.status_code}")

            if response.status_code != 200:
                messagebox.showerror("Ошибка", f"Ошибка входа: {response.status_code}")
                return

            login_response = self.session.post(post_url, data=payload, headers=headers)
            logging.info(f"Ответ на вход: {login_response.status_code}")

            if login_response.status_code == 200 and "/account/" in login_response.text:
                logging.info("Вход выполнен успешно!")
                messagebox.showinfo("Успех", "Вход выполнен успешно!")
                self.on_success(self.session)  # Переход к URL-вводу
            else:
                messagebox.showerror("Ошибка", "Ошибка входа. Проверьте логин и пароль.")
        except Exception as e:
            logging.exception("Ошибка входа")
            messagebox.showerror("Ошибка", str(e))
