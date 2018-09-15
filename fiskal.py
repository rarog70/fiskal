import os
import openpyxl

pth = os.getcwd()
os.chdir(pth)


def clear_str():
    if os.name == "posix":
        os.system("clear")
    elif os.name == "nt":
        os.system("cls")    


try:
    wb = openpyxl.load_workbook('template.xlsx')
except:
    print("Шаблон template.xlsx в каталоге " + pth + " не существует.\nСоздайте шаблон и повторите снова.")
sheet = wb.get_sheet_by_name('Лист1')
i, s = 10, 1
while s <= 20:
    clear_str()
    if s != 20:
        print("Ограничение - 20 строк на лист.\nДля выхода вместо номера набрать 'q'")
        sn = input("Фискальный накопитель № " + str(s) + "\n$: ")
    else:
        sn = input("Последний номер фискального накопителя\n$: ")
    if sn != "q":
        sheet['A' + str(i)].value = s
        sheet['B' + str(i)].value = "Фискальный накопитель"
        sn = sn.replace("Ж", ";")
        sn = sn.replace("ж", ";")
        ser = sn.split(";")
        sheet['C' + str(i)].value = ser[0]
        sheet['D' + str(i)].value = "1"
        i += 1
        s += 1
    else:
        break
wb.save("АПП_ФН.xlsx")
print("Файл АПП_ФН.xlsx сформирован")
if input("Посмотреть полученый АПП? (y/n)") == "y":
    if os.name == "posix":
        os.system("libreoffice --calc АПП_ФН.xlsx")
    elif os.name == "nt":
        os.system("АПП_ФН.xlsx")
