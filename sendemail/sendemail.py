from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
import smtplib
from string import Template
from sendemail.salary import salary

def mailcont(sendfrom, dic_person):
    # 導入 gmail 帳號密碼
    # mail_acnt = mailacnt()

    # 建一個信件的物件
    content = MIMEMultipart()
    content["subject"] = dic_person['corpname'] + dic_person['salary_year'] + "年" + dic_person['salary_mon'] + "月薪資表"
    content["from"] = sendfrom
    content["to"] = dic_person['emp_email']

    # 讀取 HTML Email 樣板
    template = Template(Path("static/source/test.html").read_text(encoding="utf-8"))

    # 將資料(dict)套進 HTML 樣板內
    body = template.substitute(dic_person)

    # 包成 content 回傳
    content.attach(MIMEText(body, "html"))

    return content


def sendbygoogle(salary_path, employee_path, mail_id, mail_pw, status):

    # 使用 salary func 抓取全部人資料
    dic_all = salary(salary_path, employee_path)

    with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:
        try:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(mail_id, mail_pw)

            for person in dic_all:
                content = mailcont(mail_id, dic_all[person])
                smtp.send_message(content)
                print("{} send Complete!".format(person))
                status.set("{} send Complete!".format(person))
            print("All Complete!")
            status.set("All Complete!")
        except Exception as e:
            print("Error:", e)
            status.set(e)

