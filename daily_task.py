import schedule
import time
import subprocess

def run_task(days):
    subprocess.run(["python", "product_request_export.py", "--headless", f"--days={days}"])

schedule.every().day.at("07:50").do(run_task, days=1)
schedule.every().day.at("12:50").do(run_task, days=2)

while True:
    schedule.run_pending()
    time.sleep(30)
