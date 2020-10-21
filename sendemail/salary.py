import pandas as pd

def salary():
    # 抓取薪資總表
    df = pd.read_excel("static/source/salary.xlsx", header=None, usecols='A:S')

    corpname = df.iloc[0, 0]        # 公司名稱
    salatitle = df.iloc[1, 0]
    salaryyr = salatitle[0:3]      # 抓取薪資的年份
    salarymon = salatitle[4:6]     # 抓取薪資的月份

    print(corpname, salatitle, salaryyr, salarymon)

    # 找出薪資表和勞健保表的起迄位置
    dfinsstart = df[df[0]=="※勞健保投保明細"].index[0]     #勞健保的表頭
    dfend = df[df[0] == "小計"].index       # 兩表表尾
    print(dfinsstart, dfend)

    df_salary = df.iloc[:dfend[0]]          # 薪資表
    df_ins = df.iloc[dfinsstart:dfend[1]]   # 勞健保表

    # 處理薪資表
    df_salary.iloc[2, 11:17] = df_salary.iloc[3, 11:17]     # 整理欄位名稱在同一欄
    df_salary = df_salary.rename(columns=df_salary.iloc[2]).drop([0, 1, 2, 3]).reset_index(drop=True)  # 設定欄位名稱、刪除不要的列、重設index

    # 處理勞健保表格
    df_ins = df_ins.reset_index(drop=True)

    # 整理欄位名稱
    healthins = df_ins.iloc[1, 2]
    laborins = df_ins.iloc[1, 5]
    df_ins.iloc[1, 2] = healthins + df_ins.iloc[2, 2]
    df_ins.iloc[1, 3] = healthins + df_ins.iloc[2, 3]
    df_ins.iloc[1, 4] = healthins + df_ins.iloc[2, 4]
    df_ins.iloc[1, 5] = laborins + df_ins.iloc[2, 5]
    df_ins.iloc[1, 6] = laborins + df_ins.iloc[2, 6]

    # 設定欄位名稱、刪除不要的列、重設index
    df_ins = df_ins[[i for i in range(8)]].rename(columns=df_ins.iloc[1]).drop([0, 1, 2]).reset_index(drop=True)

    # 讀取員工表格
    df_employee = pd.read_excel("static/source/employee.xlsx")

    # 將三個表格合併成一個大表
    df_tmp = pd.merge(df_employee, df_salary.drop(["員工"], axis=1), how='left', on='編號')
    df_all = pd.merge(df_tmp, df_ins.drop(['員工'], axis=1), how='left', on='編號')

    df_all.columns = ['編號', '員工', 'Email', '本薪', '假況扣款(應稅)', '加班費(免稅)',
           '其他給付(應稅)', '結清特休(免稅)', '伙食津貼', '小計', '薪資總額', '差異',
           '健保費', '勞保費', '補充保費', '勞退自提', '所得稅', '代扣小計', '調整',
           '實付薪資', '健保投保級距', '健保眷屬人數', '健保代扣款', '勞保投保級距', '勞保代扣款', '投保日期']

    df_all = df_all.fillna(0)

    # 資料轉成dict形式回傳
    dict_all = {}
    # for i in range(len(df_all)):
    for i in range(10):
        dict_person = {
            'corpname': corpname,
            'salary_year': salaryyr,
            'salary_mon': salarymon,
            'emp_name': df_all['員工'].iloc[i],
            'emp_email': df_all['Email'].iloc[i],
            'health_ins_amnt': '{:,d}'.format(df_all['健保投保級距'].iloc[i]),
            'labor_ins_amnt': '{:,d}'.format(df_all['勞保投保級距'].iloc[i]),
            'origin_slry': '{:,d}'.format(df_all['本薪'].iloc[i]),
            'dayoff_mslry': '{:,d}'.format(df_all['假況扣款(應稅)'].iloc[i]),
            'overdue_slry': '{:,d}'.format(df_all['加班費(免稅)'].iloc[i]),
            'other_slry': '{:,d}'.format(df_all['其他給付(應稅)'].iloc[i]),
            'food_add': '{:,d}'.format(df_all['伙食津貼'].iloc[i]),
            'tot_slry': '{:,d}'.format(df_all['小計'].iloc[i]),
            'health_ins': '{:,d}'.format(df_all['健保費'].iloc[i]),
            'labor_ins': '{:,d}'.format(df_all['勞保費'].iloc[i]),
            'add_ins': '{:,d}'.format(df_all['補充保費'].iloc[i]),
            'labor_pre': '{:,d}'.format(df_all['勞退自提'].iloc[i]),
            'slry_pre': '{:,d}'.format(df_all['所得稅'].iloc[i]),
            'extra': '{:,d}'.format(df_all['調整'].iloc[i]),
            'tot_pre': '{:,d}'.format(df_all['代扣小計'].iloc[i]),
            'act_salary': '{:,d}'.format(df_all['實付薪資'].iloc[i]),
        }
        dict_all.update({dict_person['emp_name']: dict_person})

    return dict_all
