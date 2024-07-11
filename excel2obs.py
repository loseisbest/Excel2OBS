import pandas as pd
import openpyxl  # 用于读取带宏的Excel文件
import websocket
import json
from tkinter import Tk, filedialog, Label, Entry, Button, Frame, Checkbutton, IntVar, OptionMenu, StringVar
import logging
import os
import threading
import time

# 设置日志
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# OBS WebSocket 地址和端口
obs_ws_url = "ws://localhost:4444"

class ExcelToOBS:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel2OBS 作者 B站:直播说 求一键三连 ")

        # 设置窗口图标
        icon_path = 'icon.ico'
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        else:
            logging.warning(f"Icon file not found: {icon_path}")

        self.file_path = None
        self.inputs = []
        self.previous_values = {}

        Label(root, text="Excel File:").grid(row=0, column=0)
        self.file_entry = Entry(root)
        self.file_entry.grid(row=0, column=1)
        Button(root, text="Browse", command=self.choose_file).grid(row=0, column=2)

        Label(root, text="Sheet Name:").grid(row=1, column=0)
        self.sheet_entry = Entry(root)
        self.sheet_entry.grid(row=1, column=1)

        self.inputs_frame = Frame(root)
        self.inputs_frame.grid(row=2, column=0, columnspan=4)

        self.add_input()

        Button(root, text="Add More", command=self.add_input).grid(row=3, column=0, columnspan=4)
        Button(root, text="Update Text", command=lambda: self.update_text(check_changes=False)).grid(row=4, column=0, columnspan=4)

        self.update_interval = 0.5  # 每0.5秒检测一次
        self.running = True
        self.start_update_thread()

    def choose_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        self.file_entry.delete(0, 'end')
        self.file_entry.insert(0, file_path)
        self.file_path = file_path
        logging.info(f'Selected file: {file_path}')

    def add_input(self):
        """添加新的行输入"""
        row_index = len(self.inputs)
        data_type_var = StringVar(self.inputs_frame)
        data_type_var.set("Text")  # 默认值
        row_entry = Entry(self.inputs_frame, width=5)  # 设置宽度为5
        column_entry = Entry(self.inputs_frame, width=5)  # 设置宽度为5
        name_entry = Entry(self.inputs_frame)
        value_label = Label(self.inputs_frame, text="N/A")
        check_var = IntVar()
        check_button = Checkbutton(self.inputs_frame, variable=check_var)
        data_type_menu = OptionMenu(self.inputs_frame, data_type_var, "Text", "Image")

        Label(self.inputs_frame, text=f"Input {row_index + 1}:").grid(row=row_index, column=0)
        data_type_menu.grid(row=row_index, column=1)
        name_entry.grid(row=row_index, column=2)
        row_entry.grid(row=row_index, column=3)
        column_entry.grid(row=row_index, column=4)
        value_label.grid(row=row_index, column=5)
        check_button.grid(row=row_index, column=6)

        row_entry.bind("<KeyRelease>", lambda event: self.update_value_label(row_entry, column_entry, value_label))
        column_entry.bind("<KeyRelease>", lambda event: self.update_value_label(row_entry, column_entry, value_label))

        self.inputs.append((data_type_var, row_entry, column_entry, name_entry, value_label, check_var))

    def update_value_label(self, row_entry, column_entry, value_label):
        """更新值标签"""
        row_str = row_entry.get().strip()
        column_str = column_entry.get().strip()

        if not self.file_path or not os.path.exists(self.file_path):
            logging.error("No valid Excel file selected.")
            return

        sheet_name = self.sheet_entry.get()
        if not sheet_name:
            logging.error("No sheet name provided.")
            return

        if not row_str.isdigit() or not column_str.isdigit():
            logging.error(f"Invalid row or column input: Row - {row_str}, Column - {column_str}. Row and column must be numbers.")
            return

        row = int(row_str) - 1
        column = int(column_str) - 1

        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
            if row < 0 or column < 0 or row >= len(df) or column >= len(df.columns):
                logging.error(f"Row or column out of range. Row: {row + 1}, Column: {column + 1}")
                return

            value = df.iloc[row, column]
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            logging.info(f'Read value from Excel: {value}')
            value_label.config(text=str(value))
        except Exception as e:
            logging.error(f'Error reading from Excel: {e}')

    def start_update_thread(self):
        """启动后台线程定期更新数据"""
        threading.Thread(target=self.periodic_update, daemon=True).start()

    def periodic_update(self):
        """定期更新数据"""
        while self.running:
            self.update_text(check_changes=True)
            time.sleep(self.update_interval)

    def update_text(self, check_changes=False):
        """从Excel读取数据并更新到OBS"""
        if not self.file_path:
            logging.error("No Excel file selected.")
            return

        sheet_name = self.sheet_entry.get()
        if not sheet_name:
            logging.error("No sheet name provided.")
            return

        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
            for i, (data_type_var, row_entry, column_entry, name_entry, value_label, check_var) in enumerate(self.inputs):
                row_str = row_entry.get().strip()
                column_str = column_entry.get().strip()
                source_name = name_entry.get().strip()

                if not row_str.isdigit() or not column_str.isdigit():
                    logging.error(f"Invalid row or column input: Row - {row_str}, Column - {column_str}. Row and column must be numbers.")
                    continue

                row = int(row_str) - 1
                column = int(column_str) - 1

                if row < 0 or column < 0 or row >= len(df) or column >= len(df.columns):
                    logging.error(f"Row or column out of range. Row: {row + 1}, Column: {column + 1}")
                    continue

                logging.debug(f'User Input - Row: {row + 1}, Column: {column + 1}')
                logging.debug(f'Calculated Index - Row: {row}, Column: {column}')

                try:
                    value = df.iloc[row, column]
                    if isinstance(value, float) and value.is_integer():
                        value = int(value)
                    logging.info(f'Read value from Excel: {value}')
                    value_label.config(text=str(value))

                    if source_name:
                        if check_changes:
                            if check_var.get():
                                previous_value = self.previous_values.get((row, column))
                                if previous_value != value:
                                    logging.info(f'Value changed from {previous_value} to {value}')
                                    self.send_update_to_obs(data_type_var.get(), value, source_name)
                                self.previous_values[(row, column)] = value
                        else:
                            self.send_update_to_obs(data_type_var.get(), value, source_name)
                            self.previous_values[(row, column)] = value
                except Exception as e:
                    logging.error(f'Error reading from Excel: {e}')
        except Exception as e:
            logging.error(f'Error loading Excel file: {e}')

    def send_update_to_obs(self, data_type, value, source_name):
        """根据数据类型将更新发送到OBS"""
        if data_type == "Image":
            logging.info(f"Updating image source '{source_name}' with file path: {self.clean_file_path(value)}")
            self.update_obs_image_source(self.clean_file_path(value), source_name)
        else:
            logging.info(f"Updating text source '{source_name}' with text: {str(value)}")
            self.update_obs_text_source(str(value), source_name)

    def clean_file_path(self, file_path):
        """清理文件路径中的不可见字符"""
        cleaned_path = file_path.strip()
        logging.debug(f'Original file path: {file_path}')
        # 移除不可见字符
        cleaned_path = ''.join(c for c in cleaned_path if c.isprintable())
        # 进一步处理路径中的特殊字符
        cleaned_path = cleaned_path.replace('\u202a', '').replace('\u202c', '')
        logging.debug(f'Cleaned file path: {cleaned_path}')
        return cleaned_path

    def update_obs_text_source(self, text, source_name):
        """发送文本数据到OBS的指定文本源，适用于OBS WebSocket 5.x"""
        try:
            ws = websocket.create_connection(obs_ws_url)

            identify_message = {
                "op": 1,
                "d": {
                    "rpcVersion": 1
                }
            }
            ws.send(json.dumps(identify_message))
            response = ws.recv()
            logging.info(f'Received identify response: {response}')

            update_message = {
                "op": 6,
                "d": {
                    "requestType": "SetInputSettings",
                    "requestData": {
                        "inputName": source_name,
                        "inputSettings": {
                            "text": text
                        }
                    },
                    "requestId": str(int(time.time()))
                }
            }
            ws.send(json.dumps(update_message))
            response = ws.recv()
            ws.close()

            response_data = json.loads(response)
            if response_data["d"]["requestStatus"]["result"]:
                logging.info(f'Successfully updated OBS text source: {response}')
            else:
                logging.error(f'Failed to update OBS text source: {response}')
        except Exception as e:
            logging.error(f'Failed to update OBS text source: {e}')

    def update_obs_image_source(self, file_path, source_name):
        """发送图片数据到OBS的指定图片源，适用于OBS WebSocket 5.x"""
        try:
            ws = websocket.create_connection(obs_ws_url)

            identify_message = {
                "op": 1,
                "d": {
                    "rpcVersion": 1
                }
            }
            ws.send(json.dumps(identify_message))
            response = ws.recv()
            logging.info(f'Received identify response: {response}')

            update_message = {
                "op": 6,
                "d": {
                    "requestType": "SetInputSettings",
                    "requestData": {
                        "inputName": source_name,
                        "inputSettings": {
                            "file": file_path
                        }
                    },
                    "requestId": str(int(time.time()))
                }
            }
            ws.send(json.dumps(update_message))
            response = ws.recv()
            ws.close()

            response_data = json.loads(response)
            if response_data["d"]["requestStatus"]["result"]:
                logging.info(f'Successfully updated OBS image source: {response}')
            else:
                logging.error(f'Failed to update OBS image source: {response}')
        except Exception as e:
            logging.error(f'Failed to update OBS image source: {e}')

    def stop(self):
        """停止后台线程并关闭窗口"""
        self.running = False
        self.root.destroy()

root = Tk()
app = ExcelToOBS(root)
root.protocol("WM_DELETE_WINDOW", app.stop)  # 窗口关闭时停止后台线程并销毁窗口
root.mainloop()
