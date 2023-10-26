import openpyxl
import os
import tkinter as tk
import json
import datetime

# Define constants
EXCEL_FILE = 'sample.xlsx'
TEMPLATE_FILE = './templates.json'
COLUMNS = 'a b c d e f g h i j k l m n o p q r s t u v w x y z aa ab ac ad ae af ag ah ai aj ak'.split(' ')
FILE_PATH = './templates.json'

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

# Load templates from the JSON file
def load_templates():
    try:
        with open(FILE_PATH, 'r') as TEMPLATE_FILE:
            return json.load(TEMPLATE_FILE)
    except:
        return [{'name': '--Няма създадени шаблони--'}]

# Save templates to the JSON file
def save_templates(templates):
    with open(FILE_PATH, 'w') as template_file:
        json.dump(templates, template_file, indent=4)

class PaymentBudgetApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Платежно Нареждане Бюджет')
        self.templates = load_templates()
        self.create_gui()

    def create_gui(self):
        
        self.payment_type_label = tk.Label(self.root, text = 'Метод на плащане')
        self.payment_type_label.grid(row = 0, column = 0)
        self.payment_type_var = tk.StringVar()
        self.payment_type_var.set('--Изберете метод--')
        self.payment_type_option = tk.OptionMenu(self.root, self.payment_type_var, *[adr['short_name'] for adr in ADDRESSEE])
        self.payment_type_option.grid(row = 0, column = 1)

        self.templates_var = tk.StringVar()
        self.templates_var.set('--Няма създадени шаблони--')
        self.templates_options = tk.OptionMenu(self.root, self.templates_var, *[t['name'] for t in self.templates], command = self.set_template)
        self.templates_options.grid(row = 0, column = 2)

        def create_label_entry(title, row):
            label = tk.Label(self.root, text = title)
            label.grid(row = row, column = 0)
            var = tk.StringVar()
            entry = tk.Entry(self.root, textvariable = var)
            entry.grid(row = row, column = 1)
            return var
        self.company = create_label_entry('Задължено лице',1)
        self.comapy_eik = create_label_entry('ЕИК/код по БУЛСТАТ', 2)
        self.nareditel = create_label_entry('наредител', 3)
        self.nareditel_iban = create_label_entry('IBAN на наредителя', 4)
        self.payment_sum = create_label_entry('Сума', 5)
        
        self.download_button = tk.Button(self.root, text = 'Свали', command=self.fill_excel)
        self.download_button.grid(row = 6, column = 1)

        self.save_template = create_label_entry('Запази шаблон', 7)
        self.save_temaplte_btn = tk.Button(self.root, text = 'Запази', command=self.add_template)
        self.save_temaplte_btn.grid(row = 7, column = 2, sticky='W')

        self.delete_template = tk.Button(self.root, text = 'Изтрий шаблон', command=self.delete_template, fg='red')
        self.delete_template.grid(row = 8, column = 1)
        self.error_var = tk.StringVar()
        self.error_label = tk.Label(self.root, text = self.error_var.get(), fg='red')

    def update_optionmenu(self):
        self.templates_options['menu'].delete(0, 'end')
        for opt in self.templates:
            self.templates_options['menu'].add_command(label=opt['name'], command=tk._setit(self.templates_var, opt['name']))

    def set_error(self, text, row, column):
        self.error_var.set(text)
        self.error_label.grid(row = row, column = column)

    def remove_error(self):
        self.error_var.set('')
        self.error_label.grid_remove()

    def fill_cells(self, text, row_num, sheet):
        for i, cell_value in enumerate(text):
            sheet[f'{COLUMNS[i+1]}{row_num}'] = cell_value

    def fill_excel(self, *args):
        wb = openpyxl.load_workbook(EXCEL_FILE) # Loading and set the sample.xlsx
        sheet = wb.active
        addresse = [adr for adr in ADDRESSEE if adr['short_name'] == self.payment_type_var.get()][0]
        values_iterator = [
            {
                "text": addresse['iban'], 
                "row_num": 21
            },
            {
                "text": addresse['osnovanie_row1'], 
                "row_num": 27
            },
            {
                "text": self.company.get(), 
                "row_num": 37
            },
            {
                "text": self.comapy_eik.get(), 
                "row_num": 43
            },
            {
                "text": self.nareditel.get(), 
                "row_num": 46
            },
            {
                "text": self.nareditel_iban.get(), 
                "row_num": 49
            }
        ]

        for val in values_iterator:
            self.fill_cells(val['text'], val['row_num'], sheet)

        whole_sum = ''
        if '.' not in self.payment_sum.get():
            whole_sum = f'{self.payment_sum.get()}00'
        else: whole_sum = self.payment_sum.get().replace('.', '')
        index = 0
        for i in range(len(whole_sum) - 1, -1, -1):
            find_index = COLUMNS.index('aj') - index
            sheet[f'{COLUMNS[find_index]}24'] = whole_sum[i]
            index += 1

        date = datetime.datetime.now()
        wb.save(f'./files/{self.company.get()} {self.payment_type_var.get()} {date.strftime("%d")}-{date.strftime("%m")}-{date.strftime("%Y")}-{date.strftime("%H")}-{date.strftime("%M")}-{date.strftime("%S")}.xlsx')
        # Setting IBAN

    def add_template(self):
        options = {
            'name': self.save_template.get(),
            'firm_name': self.company.get(),
            'firm_eik': self.comapy_eik.get(),
            'nareditel': self.nareditel.get(),
            'nareditel_iban': self.nareditel_iban.get(),
            'payment_type': self.payment_type_var.get()
        }
        self.templates.append(options)
        self.templates = sorted(self.templates, key = lambda t: t['name'].lower())
        save_templates(self.templates)
        self.update_optionmenu()
        self.templates_var.set(self.save_template.get())

    def set_template(self, *args):
        template = [t for t in self.templates if t['name'] == self.templates_var.get()][0]

        self.company.set(template['firm_name'])
        self.comapy_eik.set(template['firm_eik'])
        self.nareditel.set(template['nareditel'])
        self.nareditel_iban.set(template['nareditel_iban'])
        self.payment_type_var.set(template['payment_type'])

    def delete_template(self):
        template_name = [t for t in self.templates if t['name'] == self.templates_var.get()][0]
        if template_name['name'] == '--Няма създадени шаблони--': return
        self.templates.remove(template_name)
        save_templates(self.templates)
        self.update_optionmenu()
        self.templates_var.set('--Няма създадени шаблони--')


if __name__ == "__main__":
    root = tk.Tk()
    app = PaymentBudgetApp(root)
    root.geometry('800x600')
    root.mainloop()
