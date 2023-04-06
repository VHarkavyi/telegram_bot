import time
import telebot, datetime, csv, types
import csv
import openpyxl
import os
from telebot.types import ReplyKeyboardMarkup, KeyboardButton


bot = telebot.TeleBot('6133056354:AAH7DGv0JOw1Tny3_DB7mcnbY4VHmBCPFNY')

# Создание кнопок
keyboard = ReplyKeyboardMarkup(resize_keyboard=True)
add_spend_button = KeyboardButton('Добавить трату')
balance_button = KeyboardButton('Баланс')
keyboard.add(add_spend_button)
keyboard.add(balance_button)

user_global_state = {'step': 'start'}
print(user_global_state)

@bot.message_handler(commands=['start'])
def start_handler(message):
    bot.send_message(message.chat.id, 'Выберите действие: ', reply_markup=keyboard)

################################################################################################

@bot.message_handler(func=lambda message: message.text == 'Баланс')
def balance_handler(message):
    user_state = user_global_state.get(message.chat.id)
    if user_state and user_state['step'] == 'checking_availability':
        bot.send_message(message.chat.id, "Другая альпаська все еще не вышла из файла, я напишу тебе!")
        return
    user_global_state[message.chat.id] = {'step': 'balance'}
    print(user_global_state)
    def send_balance(message):
            # Открываем файл Excel
            workbook = openpyxl.load_workbook('Budget.xlsx', data_only=True)

            # Получаем нужный лист
            worksheet = workbook['Apr 23']

            # Включаем автоматический расчет формул
            workbook.calculate_dimension = 'auto'

            # Сохраняем изменения в файле
            workbook.save('Budget.xlsx')

            # Получаем значение ячейки
            value_mblack = round(worksheet['I24'].value, 2)
            value_mwhite = round(worksheet['J24'].value, 2)
            value_cash = round(worksheet['K24'].value, 2)
            value_ukrsib = round(worksheet['M24'].value, 2)
            value_privat = round(worksheet['N24'].value, 2)

            bot.send_message(message.chat.id,
                             'MonoBlack = ' + str(value_mblack) + ' UAH\n'
                             + 'MonoWhite = ' + str(value_mwhite) + ' UAH\n'
                             + 'Cash = ' + str(value_cash) + ' UAH\n'
                             + 'Ukrsib = ' + str(value_ukrsib) + ' UAH\n'
                             + 'Privat = ' + str(value_privat) + ' UAH\n'
                             , reply_markup=keyboard)
            user_global_state[message.chat.id] = {'step': 'start'}
            print(user_global_state)
    try:
        send_balance(message)
    except PermissionError:
        bot.send_message(message.chat.id, "Файл уже открыть какой-то альпаськой, я заблокиловань :(\nЯ напишу, как только файл будет доступен.")
        user_global_state[message.chat.id] = {'step': 'checking_availability'}
        print(user_global_state)
        time.sleep(30)
        check_availability(message)

#################################################################################################

@bot.message_handler(func=lambda message: message.text == 'Добавить трату')
def add_spends_handler(message):
    user_state = user_global_state.get(message.chat.id)
    if user_state and user_state['step'] == 'checking_availability':
        bot.send_message(message.chat.id, "Другая альпаська все еще не вышла из файла, я напишу тебе!")
        return
    user_global_state[message.chat.id] = {'step': 'enter_amount'}
    print(user_global_state)
    bot.send_message(message.chat.id, 'Введите сумму:')

@bot.message_handler(func=lambda message: message.text.isdigit(), content_types=['text'])
def process_spend(message):
    user_state = user_global_state.get(message.chat.id)
    if not user_state or user_state['step'] != 'enter_amount':
        return
    try:
        spend = float(message.text.replace(',', '.'))
        bot.send_message(message.chat.id, f"Добавьте комментарий: ")
        user_state['spend'] = spend
        user_state['step'] = 'enter_comment'
        print(user_global_state)
    except ValueError:
        bot.send_message(message.chat.id, "Ошибка! Введите число.")

@bot.message_handler(func=lambda message: True, content_types=['text'])
def process_comment(message):
    user_state = user_global_state.get(message.chat.id)
    if not user_state or user_state['step'] != 'enter_comment':
        return
    comment = message.text
    now = datetime.datetime.now()
    spend_entry = [user_state['spend'], now.strftime("%m/%d"), comment]
    write_data_to_file(spend_entry, message)
    user_global_state[message.chat.id] = {'step': 'start'}

###########################################################################################

def write_data_to_file(spend_entry, message):
    with open('output.csv', mode='w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file, delimiter=';', quoting=csv.QUOTE_MINIMAL)
        writer.writerow(spend_entry)
        file.close()

        def write_data_to_excel():
            # Открываем файл CSV и читаем его содержимое
            with open('output.csv', 'r', encoding='utf-8') as csv_file:
                csv_reader = csv.reader(csv_file)
                csv_data = list(csv_reader)
                print(csv_data)

            # Открываем книгу Excel
            workbook = openpyxl.load_workbook('Budget.xlsx')

            # Получаем нужный лист
            worksheet = workbook['Sheet4']

            # Находим последнюю заполненную строку в таблице
            last_row = 1
            while worksheet.cell(row=last_row, column=1).value is not None:
                last_row += 1

            # Записываем данные в Excel-файл, начиная со следующей свободной ячейки в столбце A
            for row in csv_data:
                row_data = row[0].split(';')
                for col_index, cell_value in enumerate(row_data):
                    worksheet.cell(row=last_row, column=col_index + 1, value=cell_value)
                last_row += 1

            # Сохраняем изменения в книге Excel
            workbook.save('Budget.xlsx')

        try:
            write_data_to_excel()
            bot.send_message(message.chat.id, "Спасибо, ваша трата успешно записана! Выберите следующее действие.", reply_markup=keyboard)
            os.remove('output.csv')
            user_global_state[message.chat.id] = {'step': 'start'}
            print(user_global_state)
        except PermissionError:
            bot.send_message(message.chat.id, "Файл уже открыть какой-то альпаськой, я заблокиловань :(\nЯ напишу, как только файл будет доступен.")
            user_global_state[message.chat.id] = {'step': 'checking_availability'}
            check_availability(message)

def check_availability(message):
    user_state = user_global_state[message.chat.id]
    if not user_state or user_state['step'] != 'checking_availability':
        return
    try:
        workbook = openpyxl.load_workbook('Budget.xlsx')
        workbook.save('Budget.xlsx')
        bot.send_message(message.chat.id, "Файл снова доступен!")
        user_global_state[message.chat.id] = {'step': 'start'}
        print(user_global_state)
    except PermissionError:
        time.sleep(30)
        check_availability(message)

def is_number(n):
    try:
        float(n)
    except ValueError:
        return False
    return True
###########################################################################################


# Запуск бота
bot.polling(none_stop=True)