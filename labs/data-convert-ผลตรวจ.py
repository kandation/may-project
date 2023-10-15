import numpy as np
import pandas as pd
import re
import os
import matplotlib.pyplot as plt
import datetime

pd.options.display.max_columns = None


def find_headers(dataframe):
    # Search Date/time header
    header_idx = -1
    for ri, rows in enumerate(dataframe.values):
        for ri, row in enumerate(rows):
            is_header = 'Date/Time'.lower() in str(row).strip().lower()
            if is_header:
                header_idx = ri + 1
                break

    # Check if a valid header index was found
    if header_idx != -1:
        print("Header found at row index:", header_idx)
        # Set the header and skip rows accordingly
        dataframe = pd.DataFrame(dataframe.values[header_idx:], columns=dataframe.iloc[header_idx - 1])
        dataframe.columns = ['Date/Time'] + [str(col).strip() for col in dataframe.columns[1:]]
        dataframe = dataframe.reset_index(drop=True)
    else:
        print("Header not found in the DataFrame")
    return dataframe


def find_footer(dataframe):
    row_index = None
    for i, row in dataframe.iterrows():
        if 'Station Down (D)=' in str(row['Date/Time']):
            row_index = i
            break

    print('Remove footer at row index:', row_index)

    # Remove rows from the bottom up to and including the found row
    if row_index is not None:
        print("Removing rows from index:", row_index, -1)
        dataframe = dataframe.iloc[:row_index]

        # CLean NaN last 10 rows
        df_no10 = dataframe.iloc[:-10]
        df_cleaned = dataframe.tail(10).dropna(subset=['Date/Time'], how='any')
        dataframe = pd.concat([df_no10, df_cleaned], ignore_index=True)

    else:
        print("Text 'Station Down (D)=' not found")

    # Reset the index if needed
    dataframe = dataframe.reset_index(drop=True)
    return dataframe


def cleansing(dataframe: pd.DataFrame):
    str_fill = ['F', 'C', 'D', 'A', 'P', '-', 'O']
    dataframe = dataframe.replace(str_fill, [np.NaN] * str_fill.__len__())

    # # Use a regular expression to filter columns with the pattern 'xx:xx'
    # time_pattern = r'\d+:\d+(:\d+)*'
    # time_columns = [col for col in dataframe.columns if re.search(time_pattern, str(col))]
    # print(time_columns)
    # print('Time columns:', time_columns.__len__())
    # if len(time_columns) == 24:  # Check if you have 24 time columns
    #     start_time = '1990-01-01 01:00:00'
    #     end_time = '1990-01-02 00:59:59'
    #     time_range = pd.date_range(start=start_time, end=end_time, freq='1H')
    #     time_columns = [time.strftime('%H:%M:%S') for time in time_range.time]
    # else:
    #     raise ValueError('The number of time columns is not 24')

    start_time = '1990-01-01 01:00:00'
    end_time = '1990-01-02 00:59:59'
    time_range = pd.date_range(start=start_time, end=end_time, freq='1H')
    time_columns = [time.strftime('%H:%M:%S') for time in time_range.time]

    dataframe = dataframe.iloc[:, :25]
    dataframe.columns = ['Date/Time'] + time_columns

    print('Cleansing dataframe')
    print(dataframe.isna().sum())
    print(dataframe.head(10).to_markdown())
    print(dataframe.tail(10).to_markdown())
    return dataframe


def convert_thai_date_to_datetime(thai_date):
    month_mapping = {
        'มกราคม': 'January',
        'กุมภาพันธ์': 'February',
        'มีนาคม': 'March',
        'เมษายน': 'April',
        'พฤษภาคม': 'May',
        'มิถุนายน': 'June',
        'กรกฏาคม': 'July',
        'สิงหาคม': 'August',
        'กันยายน': 'September',
        'ตุลาคม': 'October',
        'พฤศจิกายน': 'November',
        'ธันวาคม': 'December'
    }
    spx = thai_date.strip().split()
    print(spx)
    if not spx: return
    day, thai_month, thai_year = spx
    month = month_mapping.get(thai_month, 'Unknown')
    # Subtract 543 from the Buddhist year to get the Gregorian year
    gregorian_year = int(thai_year) - 543
    date_string = f"{day} {month} {gregorian_year}"
    return datetime.datetime.strptime(date_string, "%d %B %Y")


def to_serialize(dataframe: pd.DataFrame):
    # Convert Thaidate to Datetime
    dataframe['Date/Time'] = dataframe['Date/Time'].apply(convert_thai_date_to_datetime)

    print('To serialize dataframe')
    print(dataframe.dtypes)
    print(dataframe.isna().sum())
    print(dataframe.head(10).to_markdown())

    dataframe.to_excel('test.xlsx', index=False)
    # dataframe = dataframe[['Datetime'] + [col for col in dataframe.columns if col != 'Datetime']]
    melted_df = pd.melt(dataframe, id_vars=['Date/Time'], var_name='Time', value_name='Value')

    print('Melted dataframe')
    print(melted_df.isna().sum())
    print(melted_df.head(10).to_markdown())
    print(melted_df.tail(10).to_markdown())

    melted_df.rename(columns={'Date/Time': 'Datetime'}, inplace=True)

    melted_df['_time_dt'] = pd.to_datetime(melted_df['Time'], format='%H:%M:%S', errors='coerce').dt.time

    melted_df['Datetime'] += pd.to_timedelta(melted_df['_time_dt'].astype(str))
    melted_df.loc[melted_df['Time'] == '00:00:00', 'Datetime'] += pd.DateOffset(days=1)

    print('Melted dataframe 2_______')
    print(melted_df.tail(10).to_markdown())
    print(melted_df.dtypes)
    print(melted_df['Datetime'].isna().sum())

    return melted_df


def sort_series(dataframe: pd.DataFrame):
    df_sorted = dataframe.sort_values(by='Datetime')
    df_sorted.reset_index(drop=True, inplace=True)
    melted_df = df_sorted

    # Extract years from the Datetime column
    melted_df['Year'] = melted_df['Datetime'].dt.year
    melted_df['Value'] = pd.to_numeric(melted_df['Value'], errors='coerce')
    melted_df['ValueLn'] = melted_df['Value'].interpolate(method='linear')

    return melted_df


def plot_all(dataframe: pd.DataFrame, file_path: str, name: str):
    plt.rcParams['font.family'] = 'Tahoma'
    plt.figure(figsize=(12, 6))
    plt.plot(dataframe['Datetime'], dataframe['Value'], linestyle='-', lw=0.5)
    plt.xlabel('Datetime')
    plt.ylabel('Value')
    plt.title(f'Time Series Plot {name}')
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Show the plot
    plt.savefig(file_path, dpi=300)


def plot_seasonal(dataframe: pd.DataFrame, file_path: str, name: str):
    unique_years = dataframe['Year'].unique()
    # Plot seasonality for each year
    plt.figure(figsize=(24, 12))
    # First subplot on the left
    plt.rcParams['font.family'] = 'Tahoma'
    for year in unique_years:
        year_data = dataframe[dataframe['Year'] == year]
        plt.subplot(2, 1, 1)
        plt.plot(year_data['Datetime'], year_data['ValueLn'], label=f'Year {year}', lw=0.5)
        plt.ylabel('Value')
        plt.subplot(2, 1, 2)
        plt.plot(year_data['Datetime'], year_data['Value'], label=f'Year {year}', lw=0.5)
        plt.ylabel('Value')

    plt.xlabel('Datetime')

    plt.title('Seasonality Plot Each Year for ' + name)
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.legend(loc='upper right')

    plt.tight_layout()

    # Show the plot
    plt.savefig(file_path)


def parse_file_name(name):
    file = os.path.basename(name)
    file = file.split('.')[0]
    return file


def main(directory):
    dir_name = os.path.basename(directory)
    dir_name = dir_name.replace(' ', '_')

    dir_out = f'outputs/{dir_name}'
    print('Base directory:', dir_name)
    files = [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.xlsx')]

    for file_uri in files:
        f_name = parse_file_name(file_uri)
        xls = pd.ExcelFile(file_uri)
        print('File:', file_uri)
        print('Sheet nums', xls.sheet_names)
        dfs = []

        for sheet_name in xls.sheet_names:
            # Read data from the current sheet
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df = find_headers(df)
            df = find_footer(df)
            dfs.append(df)

        concatenated_df = pd.concat(dfs, ignore_index=True)
        df = concatenated_df

        df_length = len(df)
        print('Dataframe size:', df_length)
        if df_length < 15:
            print('ERROR - Dataframe size is less than 15', df_length, file_uri)
            with open('error.txt', 'a', encoding='utf-8') as f:
                f.write(f'{file_uri}\n')
            continue
        print('Concatenated dataframe')
        print(df.head(10).to_markdown())
        print(df.tail(10).to_markdown())

        df = cleansing(df)
        df = to_serialize(df)
        df = sort_series(df)

        dir_out_plot = f'{dir_out}/plot'
        os.makedirs(dir_out_plot, exist_ok=True)
        selected_col = ['Year', 'Datetime', 'Value', 'ValueLn']

        df[selected_col].to_csv(f'{dir_out}/{f_name}.csv', index=False)
        plot_all(df, f'{dir_out_plot}/{f_name}.png', f_name)
        plot_seasonal(df, f'{dir_out_plot}/{f_name}_seasonal.jpg', f_name)

        print('Saved to:', f'{dir_out}/{f_name}.csv')
        print(df.head(10).to_markdown())
        print(df.tail(10).to_markdown())


if __name__ == '__main__':
    # Get all files in the directory
    directories = [
        r'D:\Work\Home-2023\may-project\project-mei\ข้อมูลย้อนหลัง10ปี\ผลการตรวจวัดอุตุนิยมวิทยา\2010-2020'

    ]

    for directory in directories:
        recur_dirs = os.listdir(directory)
        recur_dirs = [ 'WD', 'WS']

        base_dir = os.path.basename(directory)
        print('Base directory:', base_dir)
        print('Directory:', recur_dirs)
        for dir in recur_dirs:
            join_path = os.path.join(directory, dir)
            print('Join path:', join_path)
            main(join_path)

    # main(r'D:\Work\Home-2023\may-project\project-mei\ข้อมูลย้อนหลัง10ปี\ผลการตรวจวัดอุตุนิยมวิทยา\2010-2020\RH')
