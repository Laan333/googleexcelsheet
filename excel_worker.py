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
        """ Обрабатываем ячейку, преобразуя её в строку или оставляем как есть. """
        if isinstance(cell, (datetime.datetime, datetime.date)):
            return cell.strftime('%d.%m.%Y')
        if isinstance(cell, float):  # если число с плавающей точкой
            sanitized_value = f"{cell:.6f}".replace(",", ".")
            logger.debug(f"Санитизировано число с плавающей точкой: {cell} -> {sanitized_value}")
            return sanitized_value  # Преобразуем в строку с точкой
        if isinstance(cell, int):  # если целое число
            logger.debug(f"Санитизировано целое число: {cell}")
            return str(cell)  # оставляем как строку целое число
        return cell  # оставляем все остальные типы как есть

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

        rows_copied = 0
        for row in source_ws.iter_rows(values_only=True):
            sanitized = [self._sanitize_cell(cell) for cell in row]
            logger.debug(f"Перенос строки: {row} -> {sanitized}")  # Логирование значений
            archive_ws.append(sanitized)
            rows_copied += 1

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