import tkinter as tk
from login_page import LoginPage


class EcuFirmwareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ECU Firmware App")

        # Открываем страницу логина при старте
        self.show_login_page()

    def show_login_page(self):
        """Открывает страницу логина."""
        self.clear_window()
        LoginPage(self.root, self.show_url_input_page)

    def show_url_input_page(self, session):
        """Открывает страницу ввода URL после успешного входа."""
        from url_input_page import UrlInputPage  # Импорт только при необходимости
        self.clear_window()
        UrlInputPage(self.root, session)

    def clear_window(self):
        """Очищает текущее содержимое окна перед загрузкой новой страницы."""
        for widget in self.root.winfo_children():
            widget.destroy()


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = EcuFirmwareApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Ошибка при запуске приложения: {e}")
        input("Нажмите Enter, чтобы закрыть...")
