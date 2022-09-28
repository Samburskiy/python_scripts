##планировщик задач

from apscheduler.schedulers.blocking import BlockingScheduler
import xlwings as xw
import time as t

def open_xl(filepath):
    a = 0
    while a < 3:
        try:
            wb = xw.Book(filepath)
            break
        except Exception as err:
            a+=1
            print('Ошибка ', filepath.split('\\')[-1], err)
            t.sleep(300) #пробует еще два раза запустить с интервалом в 5 минут

sched = BlockingScheduler()

sched.add_job(open_xl, 'cron', day_of_week = '0-4', hour = 15, minute = 30, misfire_grace_time = 10, args = [filepath1])
sched.add_job(open_xl, 'cron', day_of_week = '0-4', hour = 15, minute = 40, misfire_grace_time = 10, args = [filepath2])
sched.add_job(open_xl, 'cron', day_of_week = '0-4', hour = 13, misfire_grace_time = 10, args = [filepath3])
sched.add_job(open_xl, 'cron', day_of_week = '0-4', hour = 12, minute = 35, misfire_grace_time = 10, args = [filepath4])

print('Планировщик запущен')

sched.start()