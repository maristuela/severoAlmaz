# Импортируем нужный объект из библиотеки
from docxtpl import DocxTemplate
import sqlite3

import telebot



# def my_problem (id, moth, PB):
#     con = sqlite3.connect('BdTrainingCenter.db')
#     cursor = con.cursor()
#     # Запрос для добавление в план на обучение людей, которые прошли обучение, но у них истек срок 1ПБ
#     for row in id:
#         sqlite_select = f"""SELECT employee.ID
#         FROM employee
#         JOIN lerning_division ON {moth} = lerning_division.division
#         JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1
#         JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee
#         LEFT JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
#         WHERE work_schedule.name IS NULL AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1  AND training_report_PB.result = 1 AND training_report_PB.curse_num = 1 AND strftime('%Y', date('now')) - strftime('%Y', training_report_PB.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ПБ')
# --AND (strftime('%Y-%m-%d', work_schedule.start_date) >= lerning_division.start OR strftime('%Y-%m-%d', work_schedule.end_date) <= lerning_division.end_date) """
#
#     # SELECT employee.ID, lerning_division.start, lerning_division.end_date
#     #     FROM employee
#     #     JOIN lerning_division ON employee.division = lerning_division.division
#     #     JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1
#     #     JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee
#     #     JOIN work_schedule ON work_schedule.name = employee.full_name AND (strftime('%Y-%m-%d', work_schedule.start_date) > lerning_division.start OR strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date)
#     #     WHERE lerning_profession.type = 'ПБ' AND lerning_profession.number = 1 AND (strftime('%Y-%m-%d', work_schedule.start_date) >= lerning_division.start OR strftime('%Y-%m-%d', work_schedule.end_date) <= lerning_division.end_date)  AND training_report_PB.result = 1 AND training_report_PB.curse_num = 1 AND strftime('%Y', date('now')) - strftime('%Y', training_report_PB.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ПБ')
#     # SELECT employee.ID, lerning_division.start, lerning_division.end_date
#     #         FROM employee
#     #         JOIN lerning_division ON employee.division = lerning_division.division
#     #         JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1
#     #         JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee
#     #         LEFT JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
#     #         WHERE work_schedule.name IS NULL AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1  AND training_report_PB.result = 1 AND training_report_PB.curse_num = 1 AND strftime('%Y', date('now')) - strftime('%Y', training_report_PB.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ПБ')
#     cursor.execute(sqlite_select)
#
#     for row in cursor.fetchall():
#         con2 = sqlite3.connect('BdTrainingCenter.db')
#         cursorObj2 = con2.cursor()
#         albums = (row[0], number, ' ', row[1], row[2])
#         # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
#         # if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
#         #     albums = (row[0], 'Рабочие', 1, row[1], row[2])
#         # elif row[1] == 0:
#         #     albums = (row[0], 'Специалист', 1, row[1], row[2])
#         # elif row[1] == 2 or row[1] == 3 or row[1] == 5:
#         #     albums = (row[0], 'Руководители', 1, row[1], row[2])
#         # else : albums = (row[0], '', 1, row[1], row[2])
#         cursorObj2.execute("INSERT INTO plane_PB VALUES (?,?,?,?,?)", albums)
#         con2.commit()
#
#     # запрос на добавление на обучение людей, кторое ниразу не проходили обучение хотя должны были  1ПБ  людей у, которых либо неявка либо заваленно по результатам
#     sqlite_select = f"""SELECT employee.ID, lerning_division.start, lerning_division.end_date
#         FROM employee
#         JOIN lerning_division ON employee.division = lerning_division.division
#         JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = {number}
#         JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
#         LEFT JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee AND training_report_PB.curse_num = {number}
#         WHERE training_report_PB.ID_employee IS NULL OR training_report_PB.result = 0 OR training_report_PB.result = -1 AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
#
#         """
#     cursor.execute(sqlite_select)
#
#     for row in cursor.fetchall():
#         con2 = sqlite3.connect('BdTrainingCenter.db')
#         cursorObj2 = con2.cursor()
#         albums = (row[0], number, ' ', row[1], row[2])
#         # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
#         # if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
#         #     albums = (row[0], 'Рабочие', 1, row[1], row[2])
#         # elif row[1] == 0:
#         #     albums = (row[0], 'Специалист', 1, row[1], row[2])
#         # elif row[1] == 2 or row[1] == 3 or row[1] == 5:
#         #     albums = (row[0], 'Руководители', 1, row[1], row[2])
#         # else : albums = (row[0], '', 1, row[1], row[2])
#         cursorObj2.execute("INSERT INTO plane_PB VALUES (?,?,?, ?, ?)", albums)
#         con2.commit()

def table (number):

    con = sqlite3.connect('BdTrainingCenter.db')
    cursor = con.cursor()
    # Запрос для добавление в план на обучение людей, которые прошли обучение, но у них истек срок 1ПБ
    sqlite_select = f"""SELECT employee.ID, lerning_division.start, lerning_division.end_date
    FROM employee
    JOIN lerning_division ON employee.division = lerning_division.division
    JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = {number}
    JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee 
    JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) <= lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) >= lerning_division.end_date
    WHERE lerning_profession.type = 'ПБ' AND lerning_profession.number = {number} AND strftime('%Y-%m-%d', work_schedule.start_date) <= lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) >= lerning_division.end_date  AND training_report_PB.result = {number}  AND training_report_PB.curse_num = {number} AND strftime('%Y', date('now')) - strftime('%Y', training_report_PB.date) > (SELECT curse.year FROM curse WHERE curse.number = {number} AND curse.type = 'ПБ')
    """


    # SELECT employee.ID, lerning_division.start, lerning_division.end_date
    #     FROM employee
    #     JOIN lerning_division ON employee.division = lerning_division.division
    #     JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1
    #     JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee
    #     JOIN work_schedule ON work_schedule.name = employee.full_name AND (strftime('%Y-%m-%d', work_schedule.start_date) > lerning_division.start OR strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date)
    #     WHERE lerning_profession.type = 'ПБ' AND lerning_profession.number = 1 AND (strftime('%Y-%m-%d', work_schedule.start_date) >= lerning_division.start OR strftime('%Y-%m-%d', work_schedule.end_date) <= lerning_division.end_date)  AND training_report_PB.result = 1 AND training_report_PB.curse_num = 1 AND strftime('%Y', date('now')) - strftime('%Y', training_report_PB.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ПБ')
    # SELECT employee.ID, lerning_division.start, lerning_division.end_date
    #         FROM employee
    #         JOIN lerning_division ON employee.division = lerning_division.division
    #         JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1
    #         JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee
    #         LEFT JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
    #         WHERE work_schedule.name IS NULL AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1  AND training_report_PB.result = 1 AND training_report_PB.curse_num = 1 AND strftime('%Y', date('now')) - strftime('%Y', training_report_PB.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ПБ')
    cursor.execute(sqlite_select)

    for row in cursor.fetchall():
        con2 = sqlite3.connect('BdTrainingCenter.db')
        cursorObj2 = con2.cursor()
        albums = (row[0], number, ' ', row[1], row[2])
        # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
        # if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
        #     albums = (row[0], 'Рабочие', 1, row[1], row[2])
        # elif row[1] == 0:
        #     albums = (row[0], 'Специалист', 1, row[1], row[2])
        # elif row[1] == 2 or row[1] == 3 or row[1] == 5:
        #     albums = (row[0], 'Руководители', 1, row[1], row[2])
        # else : albums = (row[0], '', 1, row[1], row[2])
        cursorObj2.execute("INSERT INTO plane_PB VALUES (?,?,?,?,?)", albums)
        con2.commit()
    cursor = con.cursor()
    # Запрос для добавление в план на обучение людей, которые прошли обучение, но у них истек срок 1ПБ
    sqlite_select = f"""SELECT employee.ID, lerning_division.start, lerning_division.end_date
        FROM employee
        JOIN lerning_division ON employee.division = lerning_division.division
        JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = {number}
        JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) <= lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) >= lerning_division.end_date
        LEFT JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee AND training_report_PB.curse_num = {number}
        WHERE training_report_PB.ID_employee IS NULL OR training_report_PB.result = 0 OR training_report_PB.result = -1 AND strftime('%Y-%m-%d', work_schedule.start_date) <= lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) >= lerning_division.end_date 
"""

          # WHERE work_schedule.name IS NULL AND lerning_profession.type = 'ПБ' AND lerning_profession.number = 1  AND training_report_PB.result = 1 AND training_report_PB.curse_num = 1 AND strftime('%Y', date('now')) - strftime('%Y', training_report_PB.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ПБ')
    cursor.execute(sqlite_select)

    # for i in range(1, 2, 1):
    #     my_problem(cursor.fetchall(), i)



    # запрос на добавление на обучение людей, кторое ниразу не проходили обучение хотя должны были  1ПБ  людей у, которых либо неявка либо заваленно по результатам
    sqlite_select = f"""SELECT employee.ID, lerning_division.start, lerning_division.end_date
    FROM employee
    JOIN lerning_division ON employee.division = lerning_division.division
    JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ПБ' AND lerning_profession.number = {number}
    JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) <= lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) >= lerning_division.end_date
    LEFT JOIN training_report_PB ON employee.ID = training_report_PB.ID_employee AND training_report_PB.curse_num = {number}
    WHERE training_report_PB.ID_employee IS NULL OR training_report_PB.result = 0 OR training_report_PB.result = -1 AND strftime('%Y-%m-%d', work_schedule.start_date) <= lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) >= lerning_division.end_date

    """
    cursor.execute(sqlite_select)

    for row in cursor.fetchall():
        con2 = sqlite3.connect('BdTrainingCenter.db')
        cursorObj2 = con2.cursor()
        albums = (row[0], number, ' ', row[1], row[2])
        # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
        # if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
        #     albums = (row[0], 'Рабочие', 1, row[1], row[2])
        # elif row[1] == 0:
        #     albums = (row[0], 'Специалист', 1, row[1], row[2])
        # elif row[1] == 2 or row[1] == 3 or row[1] == 5:
        #     albums = (row[0], 'Руководители', 1, row[1], row[2])
        # else : albums = (row[0], '', 1, row[1], row[2])
        cursorObj2.execute("INSERT INTO plane_PB VALUES (?,?,?, ?, ?)", albums)
        con2.commit()

def update ():
    con1 = sqlite3.connect('BdTrainingCenter.db')
    cursor1 = con1.cursor()
    sqlite_select = """DELETE FROM plane_PB
    """



    # ПБ1
    cursor1.execute(sqlite_select)
    con1.commit()


    for i in range (1, 8, 1) :
        table(i)








update()

