import pandas as pd
import tkinter as tk
from tkinter import filedialog
from xlrd.biffh import XLRDError


def fileopen():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    print(file_path)

    try:
        df = pd.read_excel(file_path, header=None, usecols='A:S')

    except FileNotFoundError:
        print("找不到檔案，請重開一次")
        fileopen()

    except XLRDError:
        print("不支援的格式，請重開一次")
        fileopen()

    except Exception as e:
        print("其他錯誤，請洽工程人員")
        print("Error message:")
        print(e)

    print(df)

fileopen()
