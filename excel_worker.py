import openpyxl
import datetime
import os
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    datefmt='%d.%m.%Y %H:%M:%S'
)
logger = logging.getLogger("excel_worker")


class ExcelSingleFileWorker:
    def __init__(self, filepath: str, date: datetime.date = None):
        self.filepath = filepath
        self.report_sheet = "Report"
        self.archive_date = date or self._get_previous_month_date()
        self.current_month_sheet = self.archive_date.strftime("%m.%Y")

        if not os.path.exists(self.filepath):
            self._create_initial_workbook()
            logger.info(f"Создан новый Excel-файл и лист '{self.report_sheet}'")

        self.wb = openpyxl.load_workbook(self.filepath)

    def _get_previous_month_date(self) -> datetime.date:
        today = datetime.date.today()
        first_of_this_month = today.replace(day=1)
        last_month = first_of_this_month - datetime.timedelta(days=1)
        return last_month

    def _create_initial_workbook(self):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        wb.create_sheet(self.report_sheet)
        wb.save(self.filepath)

    def _sanitize_cell(self, cell):
        """ Обрабатываем ячейку, преобразуя её в строку или оставляем как есть.
        ПРИМЕЧАНИЕ: Этот метод больше не используется для архивации,
        так как мы копируем данные напрямую с сохранением форматирования. """
        if isinstance(cell, (datetime.datetime, datetime.date)):
            return cell.strftime('%d.%m.%Y')
        elif isinstance(cell, float):  # если число с плавающей точкой
            # Просто возвращаем исходное значение, без преобразований
            return cell
        elif isinstance(cell, int):  # если целое число
            return cell  # возвращаем как есть
        elif cell is None:  # если пустая ячейка
            return ""
        return str(cell).strip()  # в остальных случаях, просто преобразуем в строку

    def _create_or_replace_sheet(self, name: str):
        if name in self.wb.sheetnames:
            logger.warning(f"⚠️ Лист '{name}' уже существует. Перезапишем.")
            self.wb.remove(self.wb[name])
        else:
            logger.info(f"✅ Создан новый лист '{name}'.")
        return self.wb.create_sheet(name)

    def archive_full_report(self):
        if self.report_sheet not in self.wb.sheetnames:
            logger.warning(f"❌ Лист '{self.report_sheet}' не найден.")
            return

        source_ws = self.wb[self.report_sheet]
        archive_ws = self._create_or_replace_sheet(self.current_month_sheet)

        # переместим архивный лист сразу после "Report" (если Report есть)
        try:
            report_index = self.wb.sheetnames.index(self.report_sheet)
            self.wb._sheets.remove(archive_ws)
            self.wb._sheets.insert(report_index + 1, archive_ws)
        except Exception as e:
            logger.warning(f"⚠️ Не удалось переставить лист {self.current_month_sheet}: {e}")

        # Копируем данные напрямую без преобразования
        rows_copied = 0

        # Получаем максимальный размер данных на исходном листе
        max_row = source_ws.max_row
        max_col = source_ws.max_column

        # Копируем ширину столбцов
        for col_idx in range(1, max_col + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            if col_letter in source_ws.column_dimensions:
                archive_ws.column_dimensions[col_letter].width = source_ws.column_dimensions[col_letter].width

        # Сохраняем заголовки до удаления листа Report
        headers = None
        try:
            first_row = next(source_ws.iter_rows(min_row=1, max_row=1))
            headers = [cell.value for cell in first_row]
        except StopIteration:
            # Если лист пуст, создадим пустой список заголовков
            headers = []
            logger.warning("⚠️ Лист Report пуст, создаем пустой лист.")

        # Копируем все значения ячеек сначала
        for row_idx in range(1, max_row + 1):
            for col_idx in range(1, max_col + 1):
                # Получаем ячейку из исходного листа
                source_cell = source_ws.cell(row=row_idx, column=col_idx)
                # Создаем ячейку в целевом листе с тем же значением
                target_cell = archive_ws.cell(row=row_idx, column=col_idx)
                # Копируем значение как есть
                target_cell.value = source_cell.value

                # Копируем формат числа - это самое важное для сохранения форматирования чисел
                try:
                    target_cell.number_format = source_cell.number_format
                except Exception as e:
                    logger.warning(f"⚠️ Не удалось скопировать формат числа: {e}")

            rows_copied += 1

        logger.info(f"📥 Перенесено {rows_copied} строк из '{self.report_sheet}' в '{self.current_month_sheet}'.")

        # Теперь удаляем исходный лист Report
        self.wb.remove(source_ws)

        # И создаем новый
        new_report_ws = self.wb.create_sheet(self.report_sheet)

        # Добавляем заголовки, если они есть
        if headers:
            new_report_ws.append(headers)

        # перемещаем Report в начало
        try:
            self.wb._sheets.remove(new_report_ws)
            self.wb._sheets.insert(0, new_report_ws)
            logger.info(f"🧹 Лист '{self.report_sheet}' пересоздан и перемещён в начало.")
        except Exception as e:
            logger.warning(f"⚠️ Не удалось переместить лист '{self.report_sheet}' в начало: {e}")

        self._save()

        logger.info(f"📥 Перенесено {rows_copied} строк из '{self.report_sheet}' в '{self.current_month_sheet}'.")

        headers = [cell.value for cell in next(source_ws.iter_rows(min_row=1, max_row=1))]
        self.wb.remove(source_ws)

        new_report_ws = self.wb.create_sheet(self.report_sheet)
        new_report_ws.append(headers)

        # перемещаем Report в начало
        self.wb._sheets.remove(new_report_ws)
        self.wb._sheets.insert(0, new_report_ws)

        logger.info(f"🧹 Лист '{self.report_sheet}' пересоздан и перемещён в начало.")

        self._save()

    def _save(self):
        self.wb.save(self.filepath)
        logger.info(f"💾 Изменения сохранены в файле '{self.filepath}'.")