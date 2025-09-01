import pandas as pd
from datetime import datetime
from icalendar import Calendar, Event

# 读取课表数据
file_path = '/Users/fengxiao/Downloads/學生課表20250901104200.xlsx'  # 修改为你的文件路径
df = pd.read_excel(file_path)

# 创建一个新的日历
cal = Calendar()

# 遍历课表数据，创建 ICS 事件
for index, row in df.iterrows():
    # 解析时间
    date = datetime.strptime(str(row['日期']), '%Y-%m-%d')
    start_time = datetime.strptime(str(row['開始時間']), '%H:%M')
    end_time = datetime.strptime(str(row['結束時間']), '%H:%M')

    # 生成事件的开始和结束时间
    start_datetime = datetime.combine(date, start_time.time())
    end_datetime = datetime.combine(date, end_time.time())

    # 创建事件
    event = Event()
    event.add('summary', f"{row['科目名稱']} ({row['班別名稱']})")
    event.add('dtstart', start_datetime)
    event.add('dtend', end_datetime)
    event.add('location', row['課室'])
    event.add('description', f"Teacher: {row['教師']}")

    # 将事件添加到日历
    cal.add_component(event)

# 保存为 .ics 文件
with open('class_schedule.ics', 'wb') as f:
    f.write(cal.to_ical())