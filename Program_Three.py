import os
import pandas as pd # Пакет для роботи з файловими даними
import matplotlib.pyplot as plt # Пакет для роботи з діаграмами

def count_worker_gender_category(xlsx_file: pd.DataFrame, name_sheet: str):
    merged_df = pd.merge(xlsx_file[name_sheet], xlsx_file["all"], on=["Прізвище","Ім'я"], how="inner")
    count_categ_m = merged_df[merged_df["Стать"] == "Чоловік"].shape[0]
    count_categ_f = merged_df[merged_df["Стать"] == "Жінка"].shape[0]
    rate_categ_m = (count_categ_m/2000)*100
    rate_categ_f = (count_categ_f/2000)*100
    return count_categ_m,count_categ_f,rate_categ_m,rate_categ_f

try:
    print(" === Інформація про співробітників === \n\n")
    if os.path.exists("Table_Data.csv"):
        csv_file = pd.read_csv("Table_Data.csv")
        count_male = 0
        count_female = 0
        for _, data in csv_file.iterrows():
            if data["Стать"] == "Чоловік":
                count_male += 1
            elif data["Стать"] == "Жінка":
                count_female += 1
        rate_m = (count_male / 2000) * 100
        rate_f = (count_female / 2000) * 100
        print(f"Кількість співробітників чоловічої статі: {count_male}({int(rate_m)}%)")
        print(f"Кількість співробітників жіночої статі: {count_female}({int(rate_f)}%)\n")
    else:
        print("Такого файлу не існує.")
    if os.path.exists("File_Data.xlsx"):
        xlsx_file = pd.read_excel("File_Data.xlsx",sheet_name = None)
        count_younger_18 = xlsx_file["younger_18"]["№"].count()
        count_18_45 = xlsx_file["18-45"]["№"].count()
        count_45_70 = xlsx_file["45-70"]["№"].count()
        count_older_70 = xlsx_file["older_70"]["№"].count()
        rate_category_1 = (count_younger_18/2000)*100
        rate_category_2 = (count_18_45 / 2000) * 100
        rate_category_3 = (count_45_70 / 2000) * 100
        rate_category_4 = (count_older_70 / 2000) * 100
        print(f"Кількість співробітників молодше 18 років: {count_younger_18}({int(rate_category_1)}%)")
        print(f"Кількість співробітників від 18 до 45 років: {count_18_45}({int(rate_category_2)}%)")
        print(f"Кількість співробітників від 45 до 70 років: {count_45_70}({int(rate_category_3)}%)")
        print(f"Кількість співробітників старше 70 років: {count_older_70}({int(rate_category_4)}%)\n")

        count_younger_18_m, count_younger_18_f, rate_younger18_m, rate_younger18_f = count_worker_gender_category(xlsx_file,"younger_18")
        count_18_45_m, count_18_45_f, rate_1845_m, rate_1845_f = count_worker_gender_category(xlsx_file, "18-45")
        count_45_70_m, count_45_70_f, rate_4570_m, rate_4570_f = count_worker_gender_category(xlsx_file, "45-70")
        count_older_70_m, count_older_70_f, rate_older70_m, rate_older70_f = count_worker_gender_category(xlsx_file, "older_70")
        print(f"Кількість співробітників молодше 18 років чоловічої статі: {count_younger_18_m}({rate_younger18_m}%)")
        print(f"Кількість співробітників молодше 18 років жіночої статі: {count_younger_18_f}({rate_younger18_f}%)")
        print(f"Кількість співробітників від 18 до 45 років чоловічої статі: {count_18_45_m}({int(rate_1845_m)}%)")
        print(f"Кількість співробітників від 18 до 45 років жіночої статі: {count_18_45_f}({int(rate_1845_f)}%)")
        print(f"Кількість співробітників від 45 до 70 років чоловічої статі: {count_45_70_m}({int(rate_4570_m)}%)")
        print(f"Кількість співробітників від 45 до 70 років жіночої статі: {count_45_70_f}({int(rate_4570_f)}%)")
        print(f"Кількість співробітників старше 70 років чоловічої статі: {count_older_70_m}({int(rate_older70_m)}%)")
        print(f"Кількість співробітників старше 70 років жіночої статі: {count_older_70_f}({int(rate_older70_f)}%\n)")
        print("Будування діаграм для визначених даних...\n")

        plt.title("Кількість співробітників чоловічої та жіночої статі")
        value_1 = [count_male, count_female]
        labels_1 = ["Чоловіки", "Жінки"]
        plt.pie(value_1, labels=labels_1)
        plt.show()

        plt.title("Кількість співробітників вікових категорій")
        value_2 = [count_younger_18, count_18_45, count_45_70, count_older_70]
        labels_2 = ["Молодше 18 років", "Від 18 до 45 років", "Від 45 до 70 років", "Старше 70 років"]
        plt.pie(value_2, labels=labels_2)
        plt.show()

        plt.title("Кількість співробітників чоловічої та жіночої статі вікових категорій")
        value_3 = [count_younger_18_m, count_younger_18_f,count_18_45_m,count_18_45_f,count_45_70_m,count_45_70_f,
                   count_older_70_m,count_older_70_f]
        labels_3 = ["Чоловіки молодше 18 років", "Жінки молодше 18 років", "Чоловіки 18-45 років", "Жінки 18-45 років",
                    "Чоловіки 45-70 років", "Жінки 45-70 років", "Чоловіки старше 70 років", "Жінки старше 70 років"]
        plt.pie(value_3, labels=labels_3)
        plt.show()
    else:
        print("Такого файлу не існує.")
    print("Ок. Програму виконано успішно.")
except Exception as error:
    print(f"При виконанні програми виникла помилка: {error}")