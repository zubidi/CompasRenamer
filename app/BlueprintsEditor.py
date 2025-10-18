import pythoncom
from win32com.client import Dispatch, gencache
from pathlib import Path
import os


def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    app = api.Application  # Получаем основной интерфейс
    app.Visible = True  # Показываем окно пользователю (если скрыто)
    app.HideMessage = const.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы
    return module, api, const, app


class BlueprintsEditor:
    STAMP_IDS = {
        'Обозначение': 2,
        'Организация': 9,
        'Имя файла': 44,
        'Разработчик': 110,
        'Проверяющий': 111,
        'Дата разработки': 130,
        'Дата проверки': 131,
    }

    def __init__(self, files: list[str], data: dict, main_window):
        self.files = files
        self.data = data
        self.main_window = main_window  # ссылка на окно для логирования

        self.docs = []
        self.docs_api = []
        self.stamps = []

    def run(self):
        self.main_window.log_message("Запуск обработки чертежей...")
        module7, api7, const7, app7 = get_kompas_api7()

        stamps_data = self._load_stamps(app7, module7)
        self._process_code_replacement(stamps_data)
        self._update_personnel_fields(stamps_data)
        self._apply_to_documents(stamps_data)
        self._close_all_docs()
        self._rename_files()
        self.main_window.log_message("Обработка завершена.\n")
        self.main_window.status_var.set(f"Обработано файлов: {len(self.files)}")

    def _load_stamps(self, app, module):
        self.main_window.log_message("Загрузка данных штампов...")
        stamps_data = {key: [] for key in self.STAMP_IDS}
        self.stamp_files = []

        for file in self.files:
            try:
                doc = app.Documents.Open(PathName=file, Visible=False, ReadOnly=False)
                self.docs.append(doc)
                self.docs_api.append(module.IKompasDocument1(doc))

                for i in range(doc.LayoutSheets.Count):
                    stamp = doc.LayoutSheets.Item(i).Stamp
                    self.stamps.append(stamp)
                    self.stamp_files.append(file)

                    for name, idx in self.STAMP_IDS.items():
                        try:
                            val = stamp.Text(idx).Str
                        except Exception:
                            val = ""
                        stamps_data[name].append(val)

                self.main_window.log_message(f"Загружен штамп: {os.path.basename(file)}")

            except Exception as e:
                self.main_window.log_message(f"Ошибка загрузки {file}: {e}")

        return stamps_data

    def _process_code_replacement(self, stamps_data):
        old_code = self.data.get('old_code')
        new_code = self.data.get('new_code')

        if not old_code or not new_code:
            return

        field_flags = {
            'Обозначение': self.data.get('need_code', False),
            'Организация': self.data.get('need_org', False),
            'Имя файла': self.data.get('need_filename', False)
        }

        for field, enabled in field_flags.items():
            if not enabled:
                continue

            for i, value in enumerate(stamps_data[field]):
                if old_code in value:
                    stamps_data[field][i] = value.replace(old_code, new_code)

            self.main_window.log_message(f"Обновлено поле '{field}' — заменено '{old_code}' на '{new_code}'")

    def _update_personnel_fields(self, stamps_data):
        replacements = {
            'Разработчик': self.data.get('developer'),
            'Проверяющий': self.data.get('checker'),
            'Дата разработки': self.data.get('date_dev'),
            'Дата проверки': self.data.get('date_rev'),
        }

        for field, new_value in replacements.items():
            if new_value:
                stamps_data[field] = [new_value] * len(stamps_data[field])
                self.main_window.log_message(f"Обновлено поле '{field}' → {new_value}")

    def _apply_to_documents(self, stamps_data):
        self.main_window.log_message("Применение изменений в документах...")
        for i, stamp in enumerate(self.stamps):
            file = self.stamp_files[i]
            try:
                for field, idx in self.STAMP_IDS.items():
                    new_text = stamps_data[field][i]
                    stamp.Text(idx).Str = new_text
                stamp.Update()

                author = self.data.get('author')
                if author:
                    doc_index = self.files.index(file)
                    self.docs_api[doc_index].Author = author

                self.main_window.log_message(f"Изменения применены: {os.path.basename(file)}")
            except Exception as e:
                self.main_window.log_message(f"Ошибка при обновлении {file}: {e}")

    def _rename_files(self):
        old_code = self.data.get('old_code')
        new_code = self.data.get('new_code')
        if not old_code or not new_code or not self.data.get('need_filename', False):
            return

        self.main_window.log_message("Переименование файлов...")
        for i, file in enumerate(self.files):
            try:
                path = Path(file)
                new_name = path.name.replace(old_code, new_code)
                if new_name != path.name:
                    new_path = path.with_name(new_name)
                    os.rename(path, new_path)
                    self.main_window.log_message(f"Файл переименован: {path.name} → {new_name}")
                    self.files[i] = str(new_path)

                    self.main_window.file_listbox.delete(i)
                    self.main_window.file_listbox.insert(i, new_name)

            except Exception as e:
                self.main_window.log_message(f"Ошибка при переименовании {file}: {e}")

    def _close_all_docs(self):
        self.main_window.log_message("Закрытие всех документов...")
        for doc in self.docs:
            try:
                doc.Close(1)  # сохранить при закрытии
            except Exception:
                pass