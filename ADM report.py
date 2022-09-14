import openpyxl
book = openpyxl.Workbook ()
sheet_new = book.active
DR = 0 # значение ДР
vsego_mes1 = 0 # сумарное поедание за четыре выхода потом делем на количество контейнеров получаем среднее зе месяц


book1 = openpyxl.open(r'E:\Работа Дез-Эльтор\БЕЛГОРОД\АДМ\ОТЧЕТЫ\Чек лист 2022\08\Чек-лист АДМ 05.08.2022.xlsx', read_only=True)
sheet1 = book1.worksheets [0]
book2= openpyxl.open(r'E:\Работа Дез-Эльтор\БЕЛГОРОД\АДМ\ОТЧЕТЫ\Чек лист 2022\08\Чек-лист АДМ 12.08.2022.xlsx', read_only=True)
sheet2 = book2.worksheets[0]
book3 = openpyxl.open(r'E:\Работа Дез-Эльтор\БЕЛГОРОД\АДМ\ОТЧЕТЫ\Чек лист 2022\08\Чек-лист АДМ 19.08.2022.xlsx', read_only=True)
sheet3 = book3.worksheets [0]
book4 = openpyxl.open(r'E:\Работа Дез-Эльтор\БЕЛГОРОД\АДМ\ОТЧЕТЫ\Чек лист 2022\08\Чек-лист АДМ 26.08.2022.xlsx', read_only=True)
sheet4 = book4.worksheets[0]


# book1 = openpyxl.open(input("1 выход "), read_only=True)
# sheet1 = book1.worksheets [0]
# book2= openpyxl.open(input("2 выход "), read_only=True)
# sheet2 = book2.worksheets[0]
# book3 = openpyxl.open(input("3 выход "), read_only=True)
# sheet3 = book3.worksheets [0]
# book4 = openpyxl.open(input("4 выход "), read_only=True)
# sheet4 = book4.worksheets[0]

for row in range (sheet1.min_row, sheet1.max_row):
    znachenie1 = str(sheet1[row][0].value)
    znachenie2 = str(sheet1[row][1].value)
    znachenie4 = str(sheet2[row][1].value)
    znachenie6 = str(sheet3[row][1].value)
    znachenie8 = str(sheet4[row][1].value)
    if znachenie2 == "ДР":
        DR += 1
    if znachenie4 == "ДР":
        DR +=1
    if znachenie6 == "ДР":
        DR += 1
    if znachenie8 == "ДР":
        DR += 1
    if znachenie2.isdigit() or znachenie4.isdigit() or znachenie6.isdigit() or znachenie8.isdigit():
        if  not znachenie2.isdigit():
            znachenie2 = 0
        if  not znachenie4.isdigit():
            znachenie4 = 0
        if  not znachenie6.isdigit():
            znachenie6 = 0
        if  not znachenie8.isdigit():
            znachenie8 = 0

        summ_mes = ((int(znachenie2)+int(znachenie4)+int (znachenie6)+int (znachenie8))/4) #средняя поедаемость в контейнере за месяц
        vsego_mes1 += summ_mes
        vsego_mes = round(vsego_mes1 / 245, 2) # средняя поедаемость за месяц по обьекту
        if summ_mes > 0:
            sheet_new.append([int(znachenie1), summ_mes])

sheet_new.append(["средняя поедаемость за месяц = ", vsego_mes, "%"] )
sheet_new.append(["всего др по I-II барьеру ", DR])

# подсчет ДР и мышей по третьему барьеру

def count_dr1 (sheet_namber):
    DRIII = 0
    mish = 0
    sheet1 = book1.worksheets[sheet_namber].values
    for i in sheet1:
           for a in i:
               if a == 'ДР':
                   DRIII+=1
               if a == "Мышь":
                   m = ((i[2]),i[6])
                   for s in m:
                       if s == None:
                           s = str("a") # неважно какая буква она ниже уберется главное не цифра
                       for t in s:
                           if  t.isdigit():
                               mish += int(t)
    return  DRIII,mish
book1_DR_mishi= count_dr1(2) +count_dr1(3)+count_dr1(1)+count_dr1(4)+count_dr1(5)
book1_DR = (sum(book1_DR_mishi[:: 2]))
book1_mishi = (sum(book1_DR_mishi[1::2]))

def count_dr2 (sheet_namber):
    DRIII = 0
    mish = 0
    sheet1 = book2.worksheets[sheet_namber].values
    for i in sheet1:
           for a in i:
               if a == 'ДР':
                   DRIII+=1
               if a == "Мышь":
                   m = ((i[2]),i[6])
                   for s in m:
                       if s == None:
                           s = str("a") # неважно какая буква она ниже уберется главное не цифра
                       for t in s:
                           if  t.isdigit():
                               mish += int(t)
    return  DRIII,mish
book2_DR_mishi= count_dr2(2) +count_dr2(3)+count_dr2(1)+count_dr2(4)+count_dr2(5)
book2_DR = (sum(book2_DR_mishi[:: 2]))
book2_mishi = (sum(book2_DR_mishi[1::2]))

def count_dr3 (sheet_namber):
    DRIII = 0
    mish = 0
    sheet1 = book3.worksheets[sheet_namber].values
    for i in sheet1:
           for a in i:
               if a == 'ДР':
                   DRIII+=1
               if a == "Мышь":
                   m = ((i[2]),i[6])
                   for s in m:
                       if s == None:
                           s = str("a") # неважно какая буква она ниже уберется главное не цифра
                       for t in s:
                           if  t.isdigit():
                               mish += int(t)
    return  DRIII,mish
book3_DR_mishi= count_dr3(2) +count_dr3(3)+count_dr3(1)+count_dr3(4)+count_dr3(5)
book3_DR = (sum(book3_DR_mishi[:: 2]))
book3_mishi = (sum(book3_DR_mishi[1::2]))
def count_dr4 (sheet_namber):
    DRIII = 0
    mish = 0
    sheet1 = book4.worksheets[sheet_namber].values
    for i in sheet1:
           for a in i:
               if a == 'ДР':
                   DRIII+=1
               if a == "Мышь":
                   m = ((i[2]),i[6])
                   for s in m:
                       if s == None:
                           s = str("a") # неважно какая буква она ниже уберется главное не цифра
                       for t in s:
                           if  t.isdigit():
                               mish += int(t)
    return  DRIII,mish
book4_DR_mishi= count_dr4(2) +count_dr4(3)+count_dr4(1)+count_dr4(4)+count_dr4(5)
book4_DR = (sum(book4_DR_mishi[:: 2]))
book4_mishi = (sum(book4_DR_mishi[1::2]))
sheet_new.append(["всего ДР по третьему барьеру", book1_DR+book2_DR+book3_DR+book4_DR])
sheet_new.append(["всего мышей по третьему барьеру", book1_mishi+book2_mishi+book3_mishi+book4_mishi])
book.save ("АДМ запись сюда.xlsx")
book.close ()
