from sendemail.sendemail import sendbygoogle
import tkinter as tk
from tkinter import filedialog

def getfilepath(file):
    file_path = filedialog.askopenfilename(filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
    file.set(file_path)

def main():
    # 建立一個視窗物件
    window = tk.Tk()

    # 設定視窗基本屬性
    window.title('寶晶能源薪資寄送程式')
    window.geometry('640x480')
    window.configure(background='white')

    # 新增一個視窗內的title
    header_label = tk.Label(window, text='寶晶能源薪資寄送程式', bg='white', font=100).pack()

    # 設定讀取薪水總表路徑的變數物件
    file_salary = tk.StringVar()

    salary_frame = tk.Frame(window)
    salary_frame.pack(side=tk.TOP)
    salary_label = tk.Label(salary_frame, text='請選取薪資總表:', bg='white')
    salary_label.pack(side=tk.LEFT)

    # 製作按鈕
    salary_button = tk.Button(salary_frame, text='Open', command=lambda: getfilepath(file_salary))
    salary_button.pack(side=tk.RIGHT)

    # 顯示路徑
    sal_low_label = tk.Label(window, text="None", textvariable=file_salary, bg='white')
    sal_low_label.pack()

    # 設定讀取人員名單路徑的變數物件
    file_contact = tk.StringVar()
    contact_frame = tk.Frame(window)
    contact_frame.pack(side=tk.TOP)
    contact_label = tk.Label(contact_frame, text='請選人員通訊錄:', bg='white')
    contact_label.pack(side=tk.LEFT)

    # 製作按鈕
    contact_button = tk.Button(contact_frame, text='Open', command=lambda: getfilepath(file_contact))
    contact_button.pack(side=tk.RIGHT)

    # 顯示路徑
    con_low_label = tk.Label(window, text="None", textvariable=file_contact, bg='white')
    con_low_label.pack()

    # 設定GMAIL的號密碼格子
    mail_setup_frame = tk.Frame(window)
    mail_setup_frame.pack(side=tk.TOP)

    mail_id_frame = tk.Frame(mail_setup_frame)
    mail_id_frame.pack(side=tk.TOP)
    mail_id_label = tk.Label(mail_id_frame, text="Mail 帳號：")
    mail_id_label.pack(side=tk.LEFT)
    mail_id = tk.Entry(mail_id_frame)
    mail_id.pack(side=tk.RIGHT)

    mail_pw_frame = tk.Frame(window)
    mail_pw_frame.pack(side=tk.TOP)
    mail_pw_label = tk.Label(mail_pw_frame, text="Mail 密碼：")
    mail_pw_label.pack(side=tk.LEFT)
    mail_pw = tk.Entry(mail_pw_frame)
    mail_pw.pack(side=tk.RIGHT)

    # 按下按鈕後開始執行程式
    send2_label = tk.Button(window, text='Send!',
                            command=lambda: sendbygoogle(file_salary.get(), file_contact.get(),
                                                         mail_id.get(), mail_pw.get())).pack()
    # 開始執行視窗
    window.mainloop()


if __name__ == '__main__':
    main()