import os
import pandas as pd
from datetime import datetime

def age_filtr_data(data_filtr: pd.DataFrame):
    list_columns = ["№", "Прізвище", "Ім'я", "По батькові", "Дата народження", "Вік"]
    data_column = []
    count_num = 0
    for _,date in data_filtr.iterrows():
        if (date["Дата народження"].month,date["Дата народження"].day)>(datetime.today().month,datetime.today().day):
            age = datetime.today().year - date["Дата народження"].year - 1
        elif (date["Дата народження"].month,date["Дата народження"].day)<=(datetime.today().month,datetime.today().day):
            age = datetime.today().year - date["Дата народження"].year
        count_num += 1
        data_column.append([count_num, date["Прізвище"], date["Ім'я"], date["По батькові"],date["Дата народження"], age])
    return pd.DataFrame(data_column,columns = list_columns)

def excel_file(list_sheet_name: list,csv_file: pd.DataFrame):
    try:
        file_path = os.path.join("D:/ProjectPython/LAB#4(Employees)", "File_Data.xlsx")
        with pd.ExcelWriter(file_path,engine="xlsxwriter") as excel_writer:
            csv_file.to_excel(excel_writer, index=False, sheet_name="all")
            for sheet_name in list_sheet_name:
                if sheet_name == "younger_18":
                    data_filtr_1 = csv_file[csv_file["Дата народження"] >= datetime.strptime("2007-01-01","%Y-%m-%d")]
                    data_column_1 = age_filtr_data(data_filtr_1)
                elif sheet_name == "18-45":
                    data_filtr_2 = csv_file[(csv_file["Дата народження"] >= datetime.strptime("1980-01-01","%Y-%m-%d"))
                    & (csv_file["Дата народження"] <= datetime.strptime("2006-12-31","%Y-%m-%d"))]
                    data_column_2 = age_filtr_data(data_filtr_2)
                elif sheet_name == "45-70":
                    data_filtr_3 = csv_file[(csv_file["Дата народження"] >= datetime.strptime("1955-01-01", "%Y-%m-%d"))
                    & (csv_file["Дата народження"] <= datetime.strptime("1979-12-31", "%Y-%m-%d"))]
                    data_column_3 = age_filtr_data(data_filtr_3)
                elif sheet_name == "older_70":
                    data_filtr_4 = csv_file[csv_file["Дата народження"] <= datetime.strptime("1954-12-31", "%Y-%m-%d")]
                    data_column_4 = age_filtr_data(data_filtr_4)
            data_column_1.to_excel(excel_writer, index=False, sheet_name="younger_18")
            data_column_2.to_excel(excel_writer, index=False, sheet_name="18-45")
            data_column_3.to_excel(excel_writer, index=False, sheet_name="45-70")
            data_column_4.to_excel(excel_writer, index=False, sheet_name="older_70")
    except Exception as error:
        print(f"Під час роботи з файловими даними xlsx виникла помилка: {error}")

try:
    if os.path.exists("Table_Data.csv"):
        csv_file = pd.read_csv("Table_Data.csv")
    else:
        print("Такого файлу не існує.")
    csv_file["Дата народження"] = pd.to_datetime(csv_file["Дата народження"], format="%Y-%m-%d")
    list_sheet_name = ["younger_18", "18-45", "45-70", "older_70"]
    excel_file(list_sheet_name, csv_file)
    print("Ок. Програма завершила свою роботу успішно.")
except Exception as error:
    print(f"При виконанні програми виникла помилка: {error}")