from tkinter import *
from tkinter import ttk, scrolledtext, filedialog, messagebox
import os
from app.BlueprintsEditorDialog import BlueprintsEditorDialog
from app.BlueprintsEditor import BlueprintsEditor
from app.DetailsEditorDialog import DetailsEditorDialog
from app.DetailsEditor import DetailsEditor


class MainWindow(Tk):
    def __init__(self):
        super().__init__()
        self.title("Редактор свойств Компас 3D")
        self.geometry("800x600")

        self.file_type = StringVar(value="details")
        self.file_listbox = None
        self.files = []

        self._create_widgets()
        self._center_window()

    def log_message(self, message: str):
        self.log_text.config(state='normal')
        self.log_text.insert(END, message + '\n')
        self.log_text.see(END)
        self.log_text.config(state='disabled')
        self.update()

    def _center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _create_widgets(self):
        # Выбор типа файла
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill=X)

        ttk.Label(top_frame, text="Тип файлов:").pack(side=LEFT, padx=5)
        ttk.Radiobutton(top_frame, text="Детали", variable=self.file_type, value="details",
                        command=self.clear_files).pack(side=LEFT, padx=5)
        ttk.Radiobutton(top_frame, text="Чертежи", variable=self.file_type, value="blueprints",
                        command=self.clear_files).pack(side=LEFT, padx=5)

        # Кнопки управления файлами
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill=X)

        ttk.Button(btn_frame, text="Добавить файлы", command=self.add_files
                   ).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Очистить список", command=self.clear_files
                   ).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Обработать", command=self.process_files
                   ).pack(side=LEFT, padx=5)

        # Список выбранных файлов
        list_frame = ttk.Frame(self, padding=10)
        list_frame.pack(fill=BOTH, expand=True)

        ttk.Label(list_frame, text="Выбранные файлы:").pack(anchor=W)

        files_scrollbar = ttk.Scrollbar(list_frame)
        files_scrollbar.pack(side=RIGHT, fill=Y)

        self.file_listbox = Listbox(list_frame, yscrollcommand=files_scrollbar.set, height=10)
        self.file_listbox.pack(fill=BOTH, expand=True)
        files_scrollbar.config(command=self.file_listbox.yview)

        # Лог выполнения
        log_frame = ttk.LabelFrame(self, text="Лог выполнения", padding=10)
        log_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state='disable')
        self.log_text.pack(fill=BOTH, expand=True)

        # Статус бар
        self.status_var = StringVar(value='Готов')
        self.status_bar = Label(self, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        self.status_bar.pack(side=BOTTOM, fill=X)

    def add_files(self):
        file_type = self.file_type.get()
        filetype_map = {
            "details": [
                ("КОМПАС детали", "*.m3d *.a3d"),
                ("Все файлы", "*.*"),
            ],
            "blueprints": [
                ("КОМПАС чертежи", "*.cdw *.spw"),
                ("Все файлы", "*.*"),
            ],
        }
        file_types = filetype_map.get(file_type, [("Все файлы", "*.*")])
        files = filedialog.askopenfilenames(
            title="Выберите файлы КОМПАС 3D",
            filetypes=file_types
        )

        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.file_listbox.insert(END, os.path.basename(file))

        self.status_var.set(f"Добавлено файлов: {len(self.files)}")

    def clear_files(self):
        self.files.clear()
        self.file_listbox.delete(0, END)
        self.status_var.set("Список очищен")

    def process_files(self):
        if not self.files:
            messagebox.showwarning("Внимание", "Не выбраны файлы для обработки!")
            return
        file_type = self.file_type.get()
        if file_type == "blueprints":
            dialog = BlueprintsEditorDialog(self)
            dialog.grab_set()
            dialog.wait_window()

            if dialog.result:
                editor = BlueprintsEditor(self.files, dialog.result, self)
                editor.run()

        elif file_type == "details":
            dialog = DetailsEditorDialog(self)
            dialog.grab_set()
            self.wait_window(dialog)

            if dialog.result:
                editor = DetailsEditor(self.files, dialog.result, self)
                editor.run()


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
