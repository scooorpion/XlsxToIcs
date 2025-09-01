import pandas as pd
from datetime import datetime
from icalendar import Calendar, Event

file_path = '學生課表20250901104200.xlsx'  # 修改为你的文件路径
df = pd.read_excel(file_path)

cal = Calendar()

for index, row in df.iterrows():    
    date = datetime.strptime(str(row['日期']), '%Y-%m-%d')
    start_time = datetime.strptime(str(row['開始時間']), '%H:%M')
    end_time = datetime.strptime(str(row['結束時間']), '%H:%M')
    
    start_datetime = datetime.combine(date, start_time.time())
    end_datetime = datetime.combine(date, end_time.time())

    event = Event()
    event.add('summary', f"{row['科目名稱']} ({row['班別名稱']})")
    event.add('dtstart', start_datetime)
    event.add('dtend', end_datetime)
    event.add('location', row['課室'])
    event.add('description', f"Teacher: {row['教師']}")

    cal.add_component(event)

with open('class_schedule.ics', 'wb') as f:
    f.write(cal.to_ical())