import os
import openpyxl
from colorama import init, Fore
import re

pth = os.getcwd()
os.chdir(pth)


def clear_str():
    if os.name == "posix":
        os.system("clear")
    elif os.name == "nt":
        os.system("cls")


try:
    wb = openpyxl.load_workbook('template.xlsx')
except FileNotFoundError:
    print(f"Шаблон template.xlsx в каталоге {pth} не существует.\nСоздайте шаблон и повторите снова.")
sheet = wb.get_sheet_by_name('Лист1')
i, s = 10, 1
while s <= 20:
    clear_str()
    if s != 20:
        print(Fore.YELLOW + "Ограничение - 20 строк на лист.\nДля выхода введите 'q'")
        serial = input(Fore.GREEN + f"Фискальный накопитель № {str(s)}\n{Fore.BLUE}$: ")
    else:
        serial = input(Fore.RED + "Последний номер фискального накопителя\n$: ")
    if serial != "q":
        sheet['A' + str(i)].value = s
        sheet['B' + str(i)].value = "Фискальный накопитель"
        serial = re.sub(r"[жЖ]", ";", serial)
        ser = serial.split(";")
        sheet['C' + str(i)].value = ser[0]
        sheet['D' + str(i)].value = "1"
        i += 1
        s += 1
    else:
        break
wb.save("АПП_ФН.xlsx")
print(Fore.GREEN + "Файл АПП_ФН.xlsx сформирован")
if input("Посмотреть полученый АПП? (y/n)") == "y":
    if os.name == "posix":
        os.system("libreoffice --calc АПП_ФН.xlsx")
    elif os.name == "nt":
        os.system("АПП_ФН.xlsx")
