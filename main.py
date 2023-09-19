# Read data from xls

import pandas as pd

uri = r'D:\Work\Home-2023\may-project\project-mei\ข้อมูลย้อนหลัง10ปี\ผลการตรวจวัดคุณภาพอากาศ\SO2 2009-2020\SO2 2010-2020\ประตูผา station.xlsx'
df = pd.read_excel(uri,  skiprows=1, skipfooter=10)

print(df.to_markdown())
df['Date/Time'] = pd.to_datetime(df['Date/Time'])
df.set_index('Date/Time', inplace=True)