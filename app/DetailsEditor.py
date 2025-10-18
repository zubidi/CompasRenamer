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


class DetailsEditor:
    def __init__(self, files: list[str], data: dict, main_window):
        self.files = files
        self.data = data
        self.main_window = main_window

        self.docs3d = []
        self.docs_api = []
        self.docs1 = []

    def run(self):
        self.main_window.log_message("Запуск обработки 3D-деталей...")
        module7, api7, const7, app7 = get_kompas_api7()

        try:
            self._load_docs(app7, module7)
            self._update_codes()
            self._update_authors()
        except Exception as e:
            self.main_window.log_message(f"Ошибка выполнения: {e}")
        finally:
            self._close_all_docs()
            self._rename_files()
            self.main_window.log_message("Обработка завершена.")
            self.main_window.status_var.set(f"Обработано файлов: {len(self.files)}")

    def _load_docs(self, app, module):
        self.main_window.log_message("Загрузка документов...")
        for file in self.files:
            try:
                self.main_window.log_message(f"Открытие: {file}")
                doc3d = app.Documents.Open(PathName=file, Visible=False, ReadOnly=False)
                self.docs3d.append(doc3d)
                self.docs_api.append(module.IKompasDocument3D(doc3d))
                self.docs1.append(module.IKompasDocument1(doc3d))
                self.main_window.log_message(f"Загружено: {os.path.basename(file)}")
            except Exception as e:
                self.main_window.log_message(f"Ошибка загрузки {file}: {e}")

    def _update_codes(self):
        old_code = self.data.get("old_value")
        new_code = self.data.get("new_value")

        if not old_code or not new_code or not self.data.get("need_code", False):
            self.main_window.log_message("Обновление обозначений пропущено (не заданы параметры).")
            return

        self.main_window.log_message("Обновление обозначений...")
        for i, doc_api in enumerate(self.docs_api):
            try:
                part = doc_api.TopPart
                old_marking = part.Marking or ""
                new_marking = old_marking.replace(old_code, new_code)

                if old_marking != new_marking:
                    part.Marking = new_marking
                    part.Update()
                    doc_api.Save()
                    self.main_window.log_message(
                        f"{os.path.basename(self.files[i])}: {old_marking} → {new_marking}"
                    )
                else:
                    self.main_window.log_message(
                        f"{os.path.basename(self.files[i])}: без изменений."
                    )
            except Exception as e:
                self.main_window.log_message(f"Ошибка при обновлении {self.files[i]}: {e}")

    def _update_authors(self):
        author = self.data.get("author")
        if not author:
            self.main_window.log_message("Обновление авторов пропущено (не задан параметр).")
            return

        self.main_window.log_message(f"Обновление авторов на '{author}'...")
        for i, doc1 in enumerate(self.docs1):
            try:
                doc1.Author = author
                self.main_window.log_message(f"{os.path.basename(self.files[i])}: Автор установлен.")
            except Exception as e:
                self.main_window.log_message(f"Ошибка при изменении автора {self.files[i]}: {e}")

    def _rename_files(self):
        old_code = self.data.get("old_value")
        new_code = self.data.get("new_value")
        if not old_code or not new_code or not self.data.get("need_filename", False):
            self.main_window.log_message("Переименование файлов пропущено.")
            return

        self.main_window.log_message("Переименование файлов...")
        for i, file in enumerate(self.files):
            try:
                path = Path(file)
                new_name = path.name.replace(old_code, new_code)
                if new_name != path.name:
                    new_path = path.with_name(new_name)
                    os.rename(path, new_path)
                    self.files[i] = str(new_path)

                    self.main_window.file_listbox.delete(i)
                    self.main_window.file_listbox.insert(i, new_name)

                    self.main_window.log_message(f"{path.name} → {new_name}")
                else:
                    self.main_window.log_message(f"{path.name}: без изменений.")
            except Exception as e:
                self.main_window.log_message(f"Ошибка при переименовании {file}: {e}")

    def _close_all_docs(self):
        self.main_window.log_message("Закрытие всех документов...")
        for doc in self.docs3d:
            try:
                doc.Close(1)  # 1 — сохранить при закрытии
            except Exception as e:
                self.main_window.log_message(f"Не удалось закрыть документ: {e}")
