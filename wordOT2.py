# Импортируем нужный объект из библиотеки
from docxtpl import DocxTemplate
import sqlite3
import telebot
token = '7526701581:AAEj5YAQJ8B-_VaKAEaGjdtw4ckL_aA1u-A'
bot = telebot.TeleBot(token)
# Форма СЗ на обучение по ОТ (ОТ обучение 2)
# Загрузка шаблона


doc = DocxTemplate("Форма СЗ на обучение по ОТ (ОТ обучение 2).docx")


context = {
    'name': '',
    'profession': '',
    'division': '',
    'data': '',
    'email': '',
    'name1': '',
    'profession1': '',
    'division1': '',
    'data1': '',
    'email1': '',
    'name2': '',
    'profession2': '',
    'division2': '',
    'data2': '',
    'email2': '',
    'name3': '',
    'profession3': '',
    'division3': '',
    'data3': '',
    'email3': '',
    'name4': '',
    'profession4': '',
    'division4': '',
    'data4': '',
    'email4': ''
}


con = sqlite3.connect('BdTrainingCenter.db')
cursor = con.cursor()


# bot.send_message(message.chat.id, 'Введите ФИО сотрудника по которому не обходимо составить отчет')


num = 0
sqlite_select = """SELECT employee.full_name, code_profession.Profession, employee.division, plane_OT.date, employee.email
FROM employee
JOIN plane_OT ON employee.ID = plane_OT.ID
JOIN code_profession ON employee.profession = code_profession.ID
WHERE employee.division = 1 AND plane_OT.date = 3"""
cursor.execute(sqlite_select)
# for row_emp in cursorEmployee.fetchall():
#     if row_emp[3] == 1:
#         for row_emp in cursorEmployee.fetchall():

# Данные для заполнения шаблона
for row in cursor.fetchall():
    print(row)
    if num == 0 :
        # 'name': '',
        #     'profession': '',
        #     'division': '',
        #     'data': '',
        #     'email': '',
        context['name'] = row[0]
        context['profession'] = row[1]
        print(context['profession'])
        context['division'] = row[2]
        context['data'] = row[3]
        context['email'] = row[4]
    elif num == 1 :
        # 'name': '',
        #     'profession': '',
        #     'division': '',
        #     'data': '',
        #     'email': '',
        context['name1'] = row[0]
        context['profession1'] = row[1]
        context['division1'] = row[2]
        context['data1'] = row[3]
        context['email1'] = row[4]
    elif num == 2 :
        # 'name': '',
        #     'profession': '',
        #     'division': '',
        #     'data': '',
        #     'email': '',
        context['name2'] = row[0]
        context['profession2'] = row[1]
        context['division2'] = row[2]
        context['data2'] = row[3]
        context['email2'] = row[4]
    elif num == 3 :
        # 'name': '',
        #     'profession': '',
        #     'division': '',
        #     'data': '',
        #     'email': '',
        context['name3'] = row[0]
        context['profession3'] = row[1]
        context['division3'] = row[2]
        context['data3'] = row[3]
        context['email3'] = row[4]
    else:
        # 'name': '',
        #     'profession': '',
        #     'division': '',
        #     'data': '',
        #     'email': '',
        context['name4'] = row[0]
        context['profession4'] = row[1]
        context['division4'] = row[2]
        context['data4'] = row[3]
        context['email4'] = row[4]
        num = 0
        # Заполнение шаблона данными
        doc.render(context)

        # Сохранение документа
        doc.save("ОТ2.docx")
    num = num+1
    #     'email': '',
    context['name'] = row[0]
    context['profession'] = row[1]
    print(context['profession'])
    context['division'] = row[2]
    context['data'] = row[3]
    context['email'] = row[4]

    context['name1'] = row[0]
    context['profession1'] = row[1]
    context['division1'] = row[2]
    context['data1'] = row[3]
    context['email1'] = row[4]

    context['name2'] = row[0]
    context['profession2'] = row[1]
    context['division2'] = row[2]
    context['data2'] = row[3]
    context['email2'] = row[4]

    context['name3'] = row[0]
    context['profession3'] = row[1]
    context['division3'] = row[2]
    context['data3'] = row[3]
    context['email3'] = row[4]
    context['name4'] = row[0]
    context['profession4'] = row[1]
    context['division4'] = row[2]
    context['data4'] = row[3]
    context['email4'] = row[4]


if num != 0:
    # Заполнение шаблона данными
    doc.render(context)

    # Сохранение документа
    # doc.save("ОТ2"+ f"{row[0]}"+ ".docx")
    doc.save("ОТ2.docx")

# df.to_excel('./training_report.xlsx', sheet_name=message.text, index=False)
# doc = open(r'D:\проекты питон\severalmaz_training_center_bot\training_report.xlsx',
#                            'rb')
# bot.send_document(message.from_user.id, doc)
with open('./ОТ2.docx', 'rb') as document:
    bot.send_document('5211807364', document)
document.close()














# if __name__ == "__main__":
#     bot.polling(none_stop=True)