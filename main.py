import openpyxl
import os
import sys
import tkinter as tk
from win32com import client

templates = [
    {
        'name': 'test'
    }
]

ADDRESSEE = [
    {
        'name': 'Данъци (приходи за централния бюджет)',
        'short_name': 'Данъци',
        'iban': 'BG88BNBG96618000195001',
        'osnovanie_row1': 'Данъци и други приходи',
        'osnovanie_row2': ''
    },
    {
        'name': 'Държавно обществено осигуряване (ДОО)',
        'short_name': 'ДОО',
        'iban': 'BG97BNBG96618000112001',
        'osnovanie_row1': 'Осигурителни вноски - ДОО',
        'osnovanie_row2': ''
    },
    {
        'name': 'Здравно осигуряване (ЗОО)',
        'short_name': 'ЗОО',
        'iban': 'BG16BNBG96618000112101',
        'osnovanie_row1': 'Осигурителни вноски - НЗОК',
        'osnovanie_row2': ''
    },
    {
        'name': 'Допълнително задължително пенсионно осигуряване (ДЗПО)',
        'short_name': 'ДЗПО',
        'iban': 'BG65BNBG96618000111801',
        'osnovanie_row1': 'Осигурителни вноски - ДЗПО',
        'osnovanie_row2': ''
    }
]
ALPHABET = 'a b c d e f g h i j k l m n o p q r s t u v w x y z aa ab ac ad ae af ag ah ai aj ak'.split(' ')

window = tk.Tk(className='Accountant')
window.geometry('800x600')

def fill_excel(*args):
    if payment_type.get() == '':
        error_label = tk.Label(window, text = 'Избери метод на плащане')
    addresse = [adr for adr in ADDRESSEE if adr['short_name'] == payment_type.get()][0]
    EXCEL_FILE = [f for f in os.listdir('.') if f == 'obrazec.xlsx'][0]

    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active

    # IBAN
    for i in range(len(addresse['iban'])):
        sheet[f'{ALPHABET[i+1]}21'] = addresse['iban'][i]
        
    # OSNOVANIE 1
    for i in range(len(addresse['osnovanie_row1'])):
        sheet[f'{ALPHABET[i+1]}27'] = addresse['osnovanie_row1'][i]

    # Firm
    for i in range(len(firm.text_var.get())):
        sheet[f'{ALPHABET[i+1]}37'] = firm.text_var.get()[i]

    # Firm EIK
    for i in range(len(firm_eik.text_var.get())):
        sheet[f'{ALPHABET[i+1]}43'] = firm_eik.text_var.get()[i]

    # Nareditel
    for i in range(len(nareditel.text_var.get())):
        sheet[f'{ALPHABET[i+1]}46'] = nareditel.text_var.get()[i]

    # Nareditel IBAN
    for i in range(len(nareditel_iban.text_var.get())):
        sheet[f'{ALPHABET[i+1]}49'] = nareditel_iban.text_var.get()[i]

    # SUM
    index = 0
    for i in range(len(payment_sum.text_var.get()) - 1, -1, -1):
        find_index = ALPHABET.index('aj') - index
        sheet[f'{ALPHABET[find_index]}24'] = payment_sum.text_var.get()[i]
        index += 1

    wb.save(f'{firm.text_var.get()}.xlsx')

payment_label = tk.Label(window, text = 'Метод на плащане ')
payment_label.grid(row = 0, column=0)

payment_type = tk.StringVar()
payment_type.set('--Choose option--')

type_of_payment_drop = tk.OptionMenu(window, payment_type, *[adr['short_name'] for adr in ADDRESSEE])
type_of_payment_drop.grid(row = 0, column=1)

templates_options = tk.StringVar()
templates_options.set('Шаблони')

templates_options_menu = tk.OptionMenu(window, templates_options, *[t['name'] for t in templates])
templates_options_menu.grid(column=9, row=0)


class InputBox:
    def __init__(self, label_text, col, row):
        self.text_var = tk.StringVar()
        self.text_label = tk.Label(window, text = label_text)
        self.text_entry = tk.Entry(window, textvariable = self.text_var)
        self.text_label.grid(column = col, row = row)
        self.text_entry.grid(column = col+1, row = row)

firm = InputBox('Задължено лице', 0, 1)
firm_eik = InputBox('ЕИК/код по БУЛСТАТ', 0, 2)
nareditel = InputBox('Наредител', 0, 3)
nareditel_iban = InputBox('IBAN на наредителя', 0, 4)
payment_sum = InputBox('Сума', 0, 5)

download_button = tk.Button(window, text = 'Свали', command=fill_excel)
download_button.grid(column=1, row=6)

window.mainloop()


# 4 vida iban na poluchatel vzavisimost kakvo plashtat

