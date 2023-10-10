import openpyxl
import os
import tkinter as tk
import json
import datetime

templates = []

temapltes_file = open('./templates.json', 'r')
templates = json.loads(temapltes_file.read())
templates = sorted(templates, key = lambda t: t['name'].lower())

class ErrorMessage:

    def __init__(self):
        self.label = tk.Label(window, text='')

    def set_msg(self, text, element):
        self.label = tk.Label(window, text=text, fg='#ff1100')
        row = element.grid_info()['row']
        column = element.grid_info()['column'] + 1
        self.label.grid(row=row, column=column)

    def hide(self):
        self.label.grid_remove()

def is_float(string):
    try:
        float(string)
        return True
    except ValueError:
        return False

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
COLUMNS = 'a b c d e f g h i j k l m n o p q r s t u v w x y z aa ab ac ad ae af ag ah ai aj ak'.split(' ')

window = tk.Tk(className='Платежно Нареждане Бюджет')
window.geometry('800x600')

EXCEL_FILE = [f for f in os.listdir('.') if f == 'obrazec.xlsx'][0]

wb = openpyxl.load_workbook(EXCEL_FILE)
sheet = wb.active

error_message = ErrorMessage()

def fill_cells(text, row_num):
    for i in range(len(text)):
        sheet[f'{COLUMNS[i+1]}{row_num}'] = text[i]

def fill_excel(*args):
    error_message.hide()
    if payment_type.get() == '--Choose option--':
        error_message.set_msg('Избери метод на плащане', type_of_payment_drop)
        return
    if payment_sum.check_number() is False:
        return
    addresse = [adr for adr in ADDRESSEE if adr['short_name'] == payment_type.get()][0]

    # IBAN
    fill_cells(addresse['iban'], 21)
        
    # OSNOVANIE 1
    fill_cells(addresse['osnovanie_row1'], 27)

    # Firm
    fill_cells(firm.var.get(), 37)

    # Firm EIK
    fill_cells(firm_eik.var.get(), 43)

    # Nareditel
    fill_cells(nareditel.var.get(), 46)

    # Nareditel IBAN
    fill_cells(nareditel_iban.var.get(), 49)

    # SUM
    index = 0
    coins = 0
    sum_string = ''
    if '.' not in payment_sum.var.get():
        payment_sum.var.set(f'{payment_sum.var.get()}.00')
    
    sum_string = payment_sum.var.get().replace('.', '')
    for i in range(len(sum_string) - 1, -1, -1):
        find_index = COLUMNS.index('aj') - index - coins
        sheet[f'{COLUMNS[find_index]}24'] = sum_string[i]
        index += 1

    d = datetime.datetime.now()
    file_name = f'{firm.var.get()} {payment_type.get()} {d.strftime("%d")}-{d.strftime("%m")}-{d.strftime("%Y")} {d.strftime("%H")}-{d.strftime("%M")}-{d.strftime("%S")}.xlsx'
    wb.save(f'./FILES/{file_name}')

payment_label = tk.Label(window, text = 'Метод на плащане ')
payment_label.grid(row = 0, column=0)

payment_type = tk.StringVar()
payment_type.set('--Choose option--')

type_of_payment_drop = tk.OptionMenu(window, payment_type, *[adr['short_name'] for adr in ADDRESSEE])
type_of_payment_drop.grid(row = 0, column=1)


class InputBox:
    def __init__(self, label_text, col, row):
        self.var = tk.StringVar()
        self.label = tk.Label(window, text = label_text)
        self.entry = tk.Entry(window, textvariable = self.var)
        self.label.grid(column = col, row = row)
        self.entry.grid(column = col+1, row = row)

    def delete_value(self):
        self.entry.delete(0, len(str(self.var.get())))

    def set_value(self, value):
        self.entry.insert(0, value)

    def check_number(self):
        if is_float(self.var.get()):
            return True
        else:
            error_message.set_msg('Невалидна сума', self.entry)
            return False

firm = InputBox('Задължено лице', 0, 1)
firm_eik = InputBox('ЕИК/код по БУЛСТАТ', 0, 2)
nareditel = InputBox('Наредител', 0, 3)
nareditel_iban = InputBox('IBAN на наредителя', 0, 4)
payment_sum = InputBox('Сума', 0, 5)

download_button = tk.Button(window, text = 'Свали', command=fill_excel)
download_button.grid(column=1, row=6)

def add_template():
    options = {
        'name': save_template.var.get(),
        'firm_name': firm.var.get(),
        'firm_eik': firm_eik.var.get(),
        'nareditel': nareditel.var.get(),
        'nareditel_iban': nareditel_iban.var.get(),
        'payment_type': payment_type.get()
    }
    templates.append(options)
    new_sorted_templates = sorted(templates, key = lambda t: t['name'].lower())

    with open('./bin/templates.json', 'w') as f:
        json.dump(new_sorted_templates, f, indent = 4)
    templates_options_menu = tk.OptionMenu(window, templates_options, *[t['name'] for t in new_sorted_templates])
    templates_options_menu.grid(column=9, row=0)

save_template = InputBox('Запази шаблон:', 0, 8)
save_template_btn = tk.Button(window, text = 'Запази', command=add_template)
save_template_btn.grid(column=2, row=8)

templates_options = tk.StringVar()
templates_options.set('Шаблони')

def set_template(*args):

    template = [t for t in templates if t['name'] == templates_options.get()][0]

    firm.delete_value()
    firm_eik.delete_value()
    nareditel.delete_value()
    nareditel_iban.delete_value()
    payment_sum.delete_value()

    firm.set_value(template['firm_name'])
    firm_eik.set_value(template['firm_eik'])
    nareditel.set_value(template['nareditel'])
    nareditel_iban.set_value(template['nareditel_iban'])
    payment_type.set(template['payment_type'])

if len(templates) > 0: 
    templates_options_menu = tk.OptionMenu(window, templates_options,  *[t['name'] for t in templates])
    templates_options_menu.grid(column=9, row=0)

templates_options.trace('w', set_template)

window.mainloop()
