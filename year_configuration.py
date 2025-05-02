import datetime
import os

class MonthlyTrigger:
    def __init__(self, filename="last_trigger.txt"):
        self.filename = filename

    def _get_saved_date(self):
        if not os.path.exists(self.filename):
            return None
        with open(self.filename, 'r') as f:
            return f.read().strip()

    def _save_today_date(self):
        today_str = datetime.date.today().isoformat()
        with open(self.filename, 'w') as f:
            f.write(today_str)

    def check_and_trigger(self):
        now = datetime.datetime.now()
        today = now.date()
        time_now = now.strftime('%H:%M')

        if today.day != 2:
            return False  # не 2-е число — пошёл нах*й

        #if time_now != "06:00":
            #return False  # не то время — идём курить

        last_trigger_date = self._get_saved_date()
        if last_trigger_date == today.isoformat():
            return False  # уже выполнялось сегодня — не дублируем

        self._save_today_date()
        return True