from tkinter import *
from tkinter import ttk


class DetailsEditorDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Редактирование основной надписи")
        self.geometry("400x200")

        self.result = None

        self._center_window()
        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")

        # --- Секция замены кода ---
        ttk.Label(main_frame, text="Заменить код группы", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, columnspan=3, sticky=W, pady=(0, 10)
        )

        self.var_code = BooleanVar(value=True)
        self.var_filename = BooleanVar(value=True)

        ttk.Checkbutton(main_frame, text="Обозначение", variable=self.var_code).grid(row=1, column=0, sticky=W)
        ttk.Checkbutton(main_frame, text="Имя файла", variable=self.var_filename).grid(row=2, column=0, sticky=W)

        ttk.Label(main_frame, text="Искомая строка:").grid(row=1, column=1, sticky=E, padx=10)
        self.old_value = ttk.Entry(main_frame, width=25)
        self.old_value.grid(row=1, column=2, padx=5)
        self.old_value.focus_set()

        ttk.Label(main_frame, text="Новая строка:").grid(row=2, column=1, sticky=E, padx=10)
        self.new_value = ttk.Entry(main_frame, width=25)
        self.new_value.grid(row=2, column=2, padx=5)

        # --- Секция автора ---
        author_frame = ttk.Frame(main_frame)
        author_frame.grid(row=4, column=0, columnspan=3, sticky="ew")

        author_frame.columnconfigure(0, weight=0)
        author_frame.columnconfigure(1, weight=2)
        author_frame.columnconfigure(2, weight=1)

        ttk.Label(author_frame, text="Заменить автора", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, columnspan=3, sticky=W, pady=(10, 5)
        )

        ttk.Label(author_frame, text="Автор:").grid(row=1, column=0, sticky=E, padx=(0, 5))
        self.author = ttk.Entry(author_frame, width=25)
        self.author.grid(row=1, column=1, sticky="w", padx=5)

        # --- Кнопки ---
        button_frame = ttk.Frame(main_frame, padding=(0, 10))
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)

        ttk.Button(button_frame, text="Применить", command=self.on_ok).grid(row=0, column=0, padx=10)
        ttk.Button(button_frame, text="Отмена", command=self.destroy).grid(row=0, column=1, padx=10)

        entries = [self.old_value, self.new_value, self.author]

        def focus_next(current):
            i = entries.index(current)
            next_entry = entries[(i + 1) % len(entries)]
            next_entry.focus_set()

        def focus_prev(current):
            i = entries.index(current)
            prev_entry = entries[(i - 1) % len(entries)]
            prev_entry.focus_set()

        for entry in entries:
            entry.bind("<Down>", lambda e, cur=entry: focus_next(cur))
            entry.bind("<Up>", lambda e, cur=entry: focus_prev(cur))

    def _center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def on_ok(self):
        self.result = {'old_value': self.old_value.get(),
                       'new_value': self.new_value.get(),
                       'author': self.author.get(),
                       'need_code': self.var_code.get(),
                       'need_filename': self.var_filename.get()}
        self.destroy()
