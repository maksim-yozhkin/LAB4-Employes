from faker import Faker
from prettytable import PrettyTable
from datetime import datetime

fake_data = Faker(locale='uk_UA')
table = PrettyTable()
table.field_names = ["Прізвище","Ім'я","По батькові","Стать","Дата народження","Посада","Місто проживання",
                     "Адреса проживання","Телефон","Email"]
patronymic_m = {1:"Олександрович",2:"Васильович",3:"Богданович",4:"Кирилович",5:"Максимович",6:"Олегович",7:"Володимирович",
              8:"Леонідович",9:"Тарасович",10:"Ігорович",11:"Дмитрович",12:"Романович",13:"Петрович",14:"Едуардович",
              15:"Вітальович",16:"Ярославович",17:"Юрійович",18:"Святославович",19:"Назарович",20:"Миколайович"}
patronymic_f = {1: "Олександрівна", 2: "Василівна", 3: "Богданівна", 4: "Кирилівна", 5: "Максимівна", 6: "Олегівна",
              7: "Володимирівна", 8: "Леонідівна", 9: "Тарасівна", 10: "Ігорівна", 11: "Дмитрівна", 12: "Романівна",
              13: "Петрівна", 14: "Едуардівна",15: "Віталіївна", 16: "Ярославівна",17: "Юріївна", 18: "Святославівна",
              19: "Назарівна", 20: "Миколаївна"}
start_date = datetime.strptime('1938-01-01', '%Y-%m-%d').date()
end_date = datetime.strptime('2008-01-01', '%Y-%m-%d').date()

try:
    for patronymic in range(1, 21):
        gender = "Чоловік"
        table.add_row([fake_data.last_name(), fake_data.first_name_male(), patronymic_m[patronymic], gender,
                       fake_data.date_between(start_date=start_date, end_date=end_date), fake_data.job(),
                       fake_data.city(),
                       fake_data.address(), fake_data.phone_number(), fake_data.email()])
    for data in range(1180):
        gender = "Чоловік"
        table.add_row([fake_data.last_name(), fake_data.first_name_male(), "", gender,
                       fake_data.date_between(start_date=start_date, end_date=end_date), fake_data.job(),
                       fake_data.city(),
                       fake_data.address(), fake_data.phone_number(), fake_data.email()])
    for patronymic in range(1, 21):
        gender = "Жінка"
        table.add_row([fake_data.last_name(), fake_data.first_name_female(), patronymic_f[patronymic], gender,
                       fake_data.date_between(start_date=start_date, end_date=end_date), fake_data.job(),
                       fake_data.city(),
                       fake_data.address(), fake_data.phone_number(), fake_data.email()])
    for data in range(780):
        gender = "Жінка"
        table.add_row([fake_data.last_name(), fake_data.first_name_female(), "", gender,
                       fake_data.date_between(start_date=start_date, end_date=end_date), fake_data.job(),
                       fake_data.city(),
                       fake_data.address(), fake_data.phone_number(), fake_data.email()])
    with open("Table_Data.csv", mode="w", encoding="utf-8") as csv_file:
        csv_file.write(table.get_csv_string())
    print("Файл успішно створено.")
except Exception as error:
    print(f"При роботі з табличними даними виникла помилка: {error}")