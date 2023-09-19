import numpy as np
import pandas as pd
import tensorflow as tf
from sklearn.preprocessing import MinMaxScaler
from sklearn.metrics import mean_squared_error
from keras.models import Sequential
from keras.layers import LSTM, Dense

uri = r'D:\Work\Home-2023\may-project\project-mei\ข้อมูลย้อนหลัง10ปี\ผลการตรวจวัดคุณภาพอากาศ\NO2 2009-2019\Banhuafai.xlsx'
df = pd.read_excel(uri, skiprows=1, skipfooter=10)

df['Date/Time'] = pd.to_datetime(df['Date/Time'])
df.set_index('Date/Time', inplace=True)
df = df.iloc[:, :-6]

df = df.replace(['F', 'C', 'D', 'A', 'P'], [None, None, None, None, None ])
df = df.ffill().bfill()

# Interpolate missing values linearly
df.interpolate(method='linear', inplace=True)


print(df.to_markdown())