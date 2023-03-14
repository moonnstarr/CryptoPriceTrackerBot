import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import datetime as dt
import time
from binance import Client
import xlwings as xw
import openpyxl

class CryptoTracker:

    def __init__(self):
        self.root = tk.Tk()

        self.root.geometry("400x300")
        self.root.title("Crypto Tracker")

        self.bg = tk.PhotoImage(file='bg.png')

        # Show image using label
        self.bg_label = tk.Label(self.root, image=self.bg)
        self.bg_label.place(x=0, y=0)

        # input file path variable
        self.label_file_path = tk.Label(self.root, text='')

        self.select_button = tk.Button(self.root, text="Select File", font=("Arial", 18),
                                       command=self.select_file, bg='lightblue')
        self.select_button.pack(padx=10, pady=10)

        # output file save location variable
        self.output_save_location = tk.Label(self.root, text='')

        self.update_button = tk.Button(self.root, text="Update & Save", font=("Arial", 18),
                                       command=self.update_and_save, bg='lightblue')
        self.update_button.pack(padx=10, pady=10)

        # Progress bar
        self.my_progress = ttk.Progressbar(self.root, orient='horizontal',
                                           length=300, mode='indeterminate')
        self.my_progress.pack(pady=20)

        self.root.mainloop()

    def select_file(self):
        file_path = filedialog.askopenfilename(title='Select Input File',
                                               filetypes=(("xlsx files", "*.xlsx"), ("All Files", "*.*"))
                                               )
        self.label_file_path["text"] = file_path
        return None

    def update_and_save(self):

        # progress bar starting
        self.my_progress.start(10)
        self.root.update_idletasks()

        # Save location
        save_folder = filedialog.askdirectory(title="Select a folder to Save")
        self.output_save_location["text"] = save_folder
        save_location = self.output_save_location["text"]
        save_folder_location = r"{}".format(save_location)

        excel_file_path = self.label_file_path["text"]

        try:
            excel_file_name = r"{}".format(excel_file_path)
            input_df = pd.read_excel(excel_file_name)

        except ValueError:
            tk.messagebox.showerror("Information", "The file you have entered is invalid")
            self.my_progress.stop()
            return None
        except FileNotFoundError:
            tk.messagebox.showerror("Information", f"No such file as {excel_file_path}")
            self.my_progress.stop()
            return None

        # Clients API
        api = "hYDk9rJChtP51A310z7irr8VPE7SQZ7FUAMBPiyyduRs6gxqOUMxljXyO1cMXyTx"
        secret = "JfIw3c2nwzuVsffD8TOzQ1HLoLobUGQUsgOoWG8fXYtHPy0Vi0qcTCaC1I5NoENM"
        client = Client(api, secret)

        # GET HISTORICAL CRYPTO OHLC VALUES
        def get_hist_crypto_data():
            history_df = pd.DataFrame()
            for i, row in input_df.iterrows():
                df_row = input_df.iloc[i]
                symbol = str(df_row['coin'])
                date = str(df_row['date'])

                # creating start date and end date
                date_temp = dt.datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
                start_date = date_temp - dt.timedelta(days=1)
                end_date = date_temp + dt.timedelta(days=7)

                # converting the datetime into seconds of timestamp
                start_date_in_sec = str(
                    int(time.mktime(dt.datetime.strptime(str(start_date), "%Y-%m-%d %H:%M:%S").timetuple())))
                end_date_in_sec = str(
                    int(time.mktime(dt.datetime.strptime(str(end_date), "%Y-%m-%d %H:%M:%S").timetuple())))

                historical = client.get_historical_klines(symbol,
                                                          interval='1d',
                                                          start_str=start_date_in_sec,
                                                          end_str=end_date_in_sec)
                historical_df = pd.DataFrame(historical)
                historical_df['symbol'] = symbol
                history_df = pd.concat([history_df, historical_df], ignore_index=True, axis=0).drop_duplicates()

            return history_df

        try:
            hist_df = get_hist_crypto_data()
            hist_df.drop_duplicates()
        except:
            tk.messagebox.showerror("Information", "Data extraction process breakdown!")
            self.my_progress.stop()
            return None

        def process_data():
            # rename columns
            hist_df.columns = (['opentime',
                                'Open',
                                'High',
                                'Low',
                                'Close',
                                'Volume',
                                'Close time',
                                'Quote asset volume',
                                'Number of trades',
                                'TB base asset volume',
                                'TB quote asset volume',
                                'Unused field, ignore',
                                'symbol'])

            # convert the open time to date
            hist_df['opentime'] = pd.to_datetime(hist_df['opentime'] / 1000, unit='s')

        try:
            process_data()
            hist_df = hist_df.drop(columns=['Volume', 'Close time', 'Quote asset volume', 'Number of trades',
                                            'TB base asset volume', 'TB quote asset volume', 'Unused field, ignore'])
        except:
            tk.messagebox.showerror("Information", "Data processing error!")
            self.my_progress.stop()
            return None


        def add_yeterday_date_columns():
            concat_df = pd.DataFrame()
            for i, row in input_df.iterrows():
                input_df_row = input_df.loc[i]
                date = str(input_df_row['date'])
                date_temp = dt.datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
                yesterday = date_temp - dt.timedelta(days=1)

                new_date_series = pd.Series([yesterday])
                concat_df = pd.concat([concat_df, new_date_series], ignore_index=True, axis=0)
                input_df['Yesterday'] = concat_df[0]
            return input_df

        def add_date_columns(days_to_add, column_name):
            concat_df = pd.DataFrame()
            for i, row in input_df.iterrows():
                input_df_row = input_df.loc[i]
                date = str(input_df_row['date'])
                date_temp = dt.datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
                new_date = date_temp + dt.timedelta(days=days_to_add)

                new_date_series = pd.Series([new_date])
                concat_df = pd.concat([concat_df, new_date_series], ignore_index=True, axis=0)
                input_df[column_name] = concat_df[0]
            return input_df

        try:
            add_yeterday_date_columns()

            add_date_columns(1, 'day 1')
            add_date_columns(2, 'day 2')
            add_date_columns(3, 'day 3')
            add_date_columns(4, 'day 4')
            add_date_columns(5, 'day 5')
            add_date_columns(6, 'day 6')
            add_date_columns(7, 'day 7')
        except:
            tk.messagebox.showerror("Information", "Invalid date")
            self.my_progress.stop()
            return None

        def merge_data(df_name, input_df_date):
            # merge crypto data with input data
            df_name = pd.DataFrame()
            df_name = pd.concat([df_name, input_df], axis=1, ignore_index=True)
            df_name.columns = (['coin', 'day 0', 'Yesterday', 'day 1', 'day 2', 'day 3', 'day 4',
                                'day 5', 'day 6', 'day 7'])
            df_name = df_name[['coin', input_df_date]]

            df_name = pd.merge(left=df_name, right=hist_df, left_on=['coin', input_df_date],
                               right_on=['symbol', 'opentime'],
                               how='left', copy=False)

            df_name.columns = ([['coin', input_df_date, 'opentime', 'Open', 'High', 'Low', 'Close', 'symbol']])

            df_name = df_name[['Open', 'High', 'Low', 'Close']]

            return df_name

        try:
            yesterday_df = merge_data('yesterday_df', 'Yesterday')
            day0_df = merge_data('day0_df', 'day 0')
            day1_df = merge_data('day1_df', 'day 1')
            day2_df = merge_data('day2_df', 'day 2')
            day3_df = merge_data('day3_df', 'day 3')
            day4_df = merge_data('day4_df', 'day 4')
            day5_df = merge_data('day5_df', 'day 5')
            day6_df = merge_data('day6_df', 'day 6')
            day7_df = merge_data('day7_df', 'day 7')
        except:
            tk.messagebox.showerror("Information", "Invalid data merge")
            self.my_progress.stop()
            return None

        def final_frame():
            final_df = pd.concat([yesterday_df,
                                  input_df[['coin', 'date']],
                                  day0_df,
                                  day1_df,
                                  day2_df,
                                  day3_df,
                                  day4_df,
                                  day5_df,
                                  day6_df,
                                  day7_df], axis=1, ignore_index=True)

            # Renaming the columns
            final_df.columns = (['Open', 'High', 'Low', 'Close',
                                 'coin', 'date',
                                 'Open0', 'High0', 'Low0', 'Close0',
                                 'Open1', 'High1', 'Low1', 'Close1',
                                 'Open2', 'High2', 'Low2', 'Close2',
                                 'Open3', 'High3', 'Low3', 'Close3',
                                 'Open4', 'High4', 'Low4', 'Close4',
                                 'Open5', 'High5', 'Low5', 'Close5',
                                 'Open6', 'High6', 'Low6', 'Close6',
                                 'Open7', 'High7', 'Low7', 'Close7'])
            return final_df

        try:
            final_df = final_frame()
        except:
            tk.messagebox.showerror("Information", "Error occurred during final frame extraction!")
            self.my_progress.stop()
            return None

        def save_to_excel():
            with xw.App(visible=False) as app:
                wb = app.books.open('Crypto_Tracker.xlsx')

                wb.sheets('CryptoTracker').range('A3').value = final_df
                wb.sheets['CryptoTracker'].range('3:3').delete()
                wb.save(f"{save_folder_location}/CryptoTracker.xlsx")

        try:
            save_to_excel()
            self.my_progress.stop()
        except:
            tk.messagebox.showerror("Information", "Error occured during Excel conversion!")
            self.my_progress.stop()
            return None

        tk.messagebox.showinfo("Information", 'Updated & File Saved Successfully!')

CryptoTracker()
