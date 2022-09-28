import pymssql as ms
import xlwings as xw
import time as t
import datetime as dt
import sys as s
from pywintypes import com_error

#sql-команда
command = '''
select 'procedure' as partition, t1.procedure_name, t1.procedure_end_time, emp_mail
from t1 
left join t2
on  t1.procedure_name = t2.procedure_name
and t2.procedure_start_date = cast(getdate() as date)
and t2.procedure_partition = 'finish_check'
where t1.procedure_weekday = datepart(w, getdate()-1)
and t1.procedure_end_time <= cast(getdate() as time)
and t2.procedure_name is null
union
select 'report' as partition, t1.report_name, t1.report_daytime, emp_mail
from 
t1 
left join
(select * from table where log_type = 'Done') t2
on  t1.report_id = t2.report_id
and t1.report_weekday = t2.report_weekday
and t1.report_daytime = t2.report_daytime
and t2.stage_start_dttm > dateadd(dd, -3, getdate())
where t1.report_weekday = datepart(w, dateadd(dd, -1, getdate()))
and stage_start_dttm is null
and cast(t1.report_daytime as time) <= cast(getdate() as time)
and act_flg = 1
'''

#подключение
conn = ms.connect(params)
conn.autocommit(True)
cursor = conn.cursor()

print('Connection done')

msg = ''

#работаем с 8 до 19, каждые 5 минут; при нахождении записей понимаем, что что-то не запустилось - отправляем письмо макросом VBA
while 8 <= dt.datetime.now().hour <= 19:
    try:
        cursor.execute(command)
        result = cursor.fetchall()
        if len(result)>0:
            dct = {}
            addrs = set()
            print('wow')
            for row in result:
                addrs.add(row[3])
            for addr in addrs:
                names = []
                for row in result:
                    if row[3] == addr:
                        names.append(row[0]+' '+row[1].encode('latin1').decode('cp1251')+' в '+row[2].strftime('%H:%M'))
                dct[addr] = ';'.join(names)
            a = []
            for i in dct:
                a.append(i + '*' + dct[i])
            msg = '$'.join(a)

            app = xw.App()
            wb = app.books.open(filepath)
            mcr = wb.macro('Mail')
            mcr(msg)
            wb.close()
            app.quit()
            
            t.sleep(3000)
        else:
            t.sleep(300)
    except Exception as e:
           tm = t.asctime(t.localtime()).split( )
           print('Ошибка', s.exc_info()[0], tm[3], tm[2], tm[1], tm[4], tm[0])
           t.sleep(300)
           continue
