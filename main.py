import time

from excel_worker import ExcelSingleFileWorker
from year_configuration import MonthlyTrigger
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def main():
    logger.info('Запуск скрипта')
    while True:
        trigger = MonthlyTrigger()
        if trigger.check_and_trigger():
            manager = ExcelSingleFileWorker(filepath='test_source.xlsx')
            manager.archive_full_report()
            logger.info('Задача выполнена.')
        time.sleep(60)

if __name__ == '__main__':
    main()
