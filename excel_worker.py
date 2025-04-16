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
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.report_sheet = "Report"
        self.current_month_sheet = datetime.datetime.today().strftime("%m.%Y")

        if not os.path.exists(self.filepath):
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            wb.create_sheet(self.report_sheet)
            wb.save(self.filepath)
            logger.info(f"Создан новый Excel-файл и лист '{self.report_sheet}'")

        self.wb = openpyxl.load_workbook(self.filepath)

    def _sanitize_cell(self, cell):
        if isinstance(cell, (datetime.datetime, datetime.date)):
            return cell.strftime('%d.%m.%Y')
        elif cell is None:
            return ""
        return str(cell).strip()

    def archive_full_report(self):
        if self.report_sheet not in self.wb.sheetnames:
            logger.warning(f"❌ Лист '{self.report_sheet}' не найден.")
            return

        source_ws = self.wb[self.report_sheet]

        # создаём или получаем лист текущего месяца
        if self.current_month_sheet not in self.wb.sheetnames:
            archive_ws = self.wb.create_sheet(self.current_month_sheet)
            logger.info(f"✅ Создан новый лист '{self.current_month_sheet}' для переноса.")
        else:
            logger.warning(f"⚠️ Лист '{self.current_month_sheet}' уже существует. Перезапишем.")
            self.wb.remove(self.wb[self.current_month_sheet])
            archive_ws = self.wb.create_sheet(self.current_month_sheet)

        rows_copied = 0
        for row in source_ws.iter_rows(values_only=True):
            sanitized = [self._sanitize_cell(cell) for cell in row]
            archive_ws.append(sanitized)
            rows_copied += 1

        logger.info(f"📥 Перенесено {rows_copied} строк из '{self.report_sheet}' в '{self.current_month_sheet}'.")

        # Копируем заголовки из старого листа
        headers = [cell.value for cell in next(source_ws.iter_rows(min_row=1, max_row=1))]

        # удаляем старый лист
        self.wb.remove(source_ws)

        # создаём новый чистый Report
        new_ws = self.wb.create_sheet(self.report_sheet)
        new_ws.append(headers)
        logger.info(f"🧹 Лист '{self.report_sheet}' пересоздан, оставлены только заголовки.")

        # сохраняем
        self.wb.save(self.filepath)
        logger.info(f"💾 Изменения сохранены в файле '{self.filepath}'.")
