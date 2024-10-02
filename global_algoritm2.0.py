# Импортируем нужный объект из библиотеки
from docxtpl import DocxTemplate
import sqlite3

import telebot

def update ():
    con1 = sqlite3.connect('BdTrainingCenter.db')
    cursor1 = con1.cursor()
    sqlite_select = """DELETE FROM plane_OT
    """
    cursor1.execute(sqlite_select)
    con1.commit()

    con = sqlite3.connect('BdTrainingCenter.db')
    cursor = con.cursor()

    num = 0
    sqlite_select = """SELECT employee.ID, employee.profession, employee.profession, lerning_division.month
    FROM employee
    JOIN lerning_division ON employee.division = lerning_division.division
    JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ОТ' AND lerning_profession.number = 1
    JOIN training_report_OT ON employee.ID = training_report_OT.ID_employee 
    JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
    WHERE lerning_profession.type = 'ОТ' AND lerning_profession.number = 1  AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date AND training_report_OT.curse_num = 1 AND strftime('%Y', date('now')) - strftime('%Y', training_report_OT.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ОТ')
    GROUP BY employee.ID
    """
    cursor.execute(sqlite_select)

    for row in cursor.fetchall():
        con2 = sqlite3.connect('BdTrainingCenter.db')
        cursorObj2 = con2.cursor()
        # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
        if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
            albums = (row[0], 'Рабочие', 1, row[1],'', row[2], row[3])
        elif row[1] == 0:
            albums = (row[0], 'Специалист', 1, row[1],'', row[2], row[3])
        elif row[1] == 2 or row[1] == 3 or row[1] == 5:
            albums = (row[0], 'Руководители', 1, row[1],'', row[2], row[3])
        else : albums = (row[0], '', 1, row[1], '', row[2], row[3])
        cursorObj2.execute("INSERT INTO plane_OT VALUES (?,?,?,?,?,?,?)", albums)
        con2.commit()

    sqlite_select = """SELECT employee.ID, employee.profession, lerning_division.start, lerning_division.end_date
    FROM employee
    JOIN lerning_division ON employee.division = lerning_division.division
    JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ОТ' AND lerning_profession.number = 1
    LEFT JOIN training_report_OT ON employee.ID = training_report_OT.ID_employee AND training_report_OT.curse_num = 1
    JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
    WHERE training_report_OT.ID_employee IS NULL AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date  
    GROUP BY employee.ID
    """
    cursor.execute(sqlite_select)

    for row in cursor.fetchall():
        con2 = sqlite3.connect('BdTrainingCenter.db')
        cursorObj2 = con2.cursor()
        # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
        if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
            albums = (row[0], 'Рабочие', 1, row[1],  '', row[2], row[3])
        elif row[1] == 0:
            albums = (row[0], 'Специалист', 1, row[1], '', row[2], row[3])
        elif row[1] == 2 or row[1] == 3 or row[1] == 5:
            albums = (row[0], 'Руководители', 1, row[1],  '', row[2], row[3])
        else : albums = (row[0], '', 1, row[1],  '', row[2], row[3])
        cursorObj2.execute("INSERT INTO plane_OT VALUES (?,?,?,?,?,?,?)", albums)
        con2.commit()



    # ОТ2

    con = sqlite3.connect('BdTrainingCenter.db')
    cursor = con.cursor()
    sqlite_select = """SELECT employee.ID,employee.profession, lerning_division.start, lerning_division.end_date
        FROM employee
        JOIN lerning_division ON employee.division = lerning_division.division
            JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ОТ' AND lerning_profession.number = 2
            JOIN training_report_OT ON employee.ID = training_report_OT.ID_employee
            JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
            WHERE lerning_profession.type = 'ОТ' AND lerning_profession.number = 2  AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date AND training_report_OT.curse_num = 2 AND strftime('%Y', date('now')) - strftime('%Y', training_report_OT.date) > (SELECT curse.year FROM curse WHERE curse.number = 2 AND curse.type = 'ОТ')
            GROUP BY employee.ID
            """
    cursor.execute(sqlite_select)

    for row in cursor.fetchall():
        con2 = sqlite3.connect('BdTrainingCenter.db')
        cursorObj2 = con2.cursor()
            # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
        if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
            albums = (row[0], 'Рабочие', 2, row[1],  '', row[2], row[3])
        elif row[1] == 0:
            albums = (row[0], 'Специалист', 2, row[1],  '', row[2], row[3])
        elif row[1] == 2 or row[1] == 3 or row[1] == 5:
            albums = (row[0], 'Руководители', 2, row[1],  '', row[2], row[3])
        else:
            albums = (row[0], '', 2, row[1],  '', row[2], row[3])
        cursorObj2.execute("INSERT INTO plane_OT VALUES (?,?,?,?,?, ?,?)", albums)
        con2.commit()

    sqlite_select = """SELECT employee.ID,employee.profession, lerning_division.start, lerning_division.end_date
            FROM employee
            JOIN lerning_division ON employee.division = lerning_division.division
            JOIN lerning_profession ON employee.profession = lerning_profession.ID AND lerning_profession.type = 'ОТ' AND lerning_profession.number = 2
            LEFT JOIN training_report_OT ON employee.ID = training_report_OT.ID_employee AND training_report_OT.curse_num = 2
            JOIN work_schedule ON work_schedule.name = employee.full_name AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date
            WHERE training_report_OT.ID_employee IS NULL AND lerning_profession.type = 'ОТ' AND lerning_profession.number = 2 AND strftime('%Y-%m-%d', work_schedule.start_date) < lerning_division.start AND strftime('%Y-%m-%d', work_schedule.end_date) > lerning_division.end_date 
            GROUP BY employee.ID
            """
    cursor.execute(sqlite_select)

    for row in cursor.fetchall():
        con2 = sqlite3.connect('BdTrainingCenter.db')
        cursorObj2 = con2.cursor()
            # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
        if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
            albums = (row[0], 'Рабочие', 2, row[1],  '', row[2], row[3])
        elif row[1] == 0:
            albums = (row[0], 'Специалист', 2, row[1],  '', row[2], row[3])
        elif row[1] == 2 or row[1] == 3 or row[1] == 5:
            albums = (row[0], 'Руководители', 2, row[1],  '', row[2], row[3])
        else:
            albums = (row[0], '', 2, row[1], row[2])
        cursorObj2.execute("INSERT INTO plane_OT VALUES (?,?,?,?,?, ?, ?)", albums)
        con2.commit()






update()




# SELECT employee.ID, employee.profession, lerning_division.month
# FROM employee
# JOIN lerning_division ON employee.division = lerning_division.division
# JOIN lerning_profession ON employee.profession = lerning_profession.ID
# LEFT JOIN training_report_OT ON employee.ID = training_report_OT.ID_employee AND training_report_OT.curse_num = 1
# WHERE training_report_OT.ID_employee IS NULL;
#
#
#
#
#
# /*
# SELECT employee.ID
# FROM employee
# LEFT JOIN training_report_OT ON employee.ID = training_report_OT.ID_employee AND training_report_OT.curse_num = 1
# WHERE training_report_OT.ID_employee IS NULL;*/




























# !!!!!!!
#
#
# con = sqlite3.connect('BdTrainingCenter.db')
# cursor = con.cursor()
#
# num = 0
# sqlite_select = """SELECT employee.ID, employee.profession, lerning_division.month
#     FROM employee
#     JOIN lerning_division ON employee.division = lerning_division.division
#     JOIN lerning_profession ON employee.profession = lerning_profession.ID
#     JOIN training_report_OT ON employee.ID = training_report_OT.ID_employee
#     WHERE lerning_profession.type = 'ОТ' AND lerning_profession.number = 2 AND strftime('%Y', date('now')) - strftime('%Y', training_report_OT.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ОТ')
#     GROUP BY ID_employee
#     """
# cursor.execute(sqlite_select)
#
# for row in cursor.fetchall():
#     con2 = sqlite3.connect('BdTrainingCenter.db')
#     cursorObj2 = con2.cursor()
#     # if row[3] == "Машинист мельниц" or row[3] == "Сепараторщик" or row[3] == "Машинист конвейера" or row[3] == "Машинист насосных установок" or row[3] == "Машинист установок по разрушению "
#     if row[1] == 4 or row[1] == 6 or row[1] == 7 or row[1] == 1 or row[1] == 8 or row[1] == 9:
#         albums = (row[0], 'Рабочие', 1, row[1], row[2])
#     elif row[1] == 0:
#         albums = (row[0], 'Специалист', 1, row[1], row[2])
#     elif row[1] == 2 or row[1] == 3 or row[1] == 5:
#         albums = (row[0], 'Руководители', 1, row[1], row[2])
#     else:
#         albums = (row[0], '', 1, row[1], row[2])
#     cursorObj2.execute("INSERT INTO plane_OT VALUES (?,?,?,?,?)", albums)
#     con2.commit()
#
#
#
#
#
#
#
#
#
#
#
#
#
#




















































    # JOIN lerning_profession ON employee.profession = lerning_profession.ID
    # WHERE employee.division = 1 AND plane_OT.date = 3

    # !!!!!!!!!!
    # SELECT employee.ID, employee.profession, lerning_division.month, MAX(training_report_OT.date) AS max_date, strftime('%Y', training_report_OT.date) AS year, date('now'), strftime('%Y', date('now')) - strftime('%Y', training_report_OT.date)
    # FROM employee
    # JOIN lerning_division ON employee.division = lerning_division.division
    # JOIN lerning_profession ON employee.profession = lerning_profession.ID
    # JOIN training_report_OT ON employee.ID = training_report_OT.ID_employee
    # WHERE lerning_profession.type = 'ОТ' AND lerning_profession.number = 1 AND strftime('%Y', date('now')) - strftime('%Y', training_report_OT.date) > (SELECT curse.year FROM curse WHERE curse.number = 1 AND curse.type = 'ОТ')
    # GROUP BY ID_employee
    #
    #
    #
    # --WHERE employee.division = 1 AND plane_OT.date = 3
    # --# SELECT MAX(date) AS max_date FROM training_report_OT
    #
    # --JOIN (/*
    # --    SELECT ID_employee, MAX(date) AS max_date
    # --    FROM training_report_OT
    # --    GROUP BY ID_employee
    # --) tr ON e.ID = tr.ID_employee*/

# SELECT MAX(date) AS max_date FROM training_report_OT