import config
import telebot
import Generate_xls_file
from telebot import types
import xlrd, xlwt
from xlrd import open_workbook
from xlutils.copy import copy

bot = telebot.TeleBot('token')

@bot.message_handler(content_types=["text"])

def start(m):
    #bot.send_message(m.chat.id, "Привет, введи своё имя")
    bot.send_message(m.chat.id, "[День,Модель,Время,Сумма,Сотрудник]\nПример ввода - 01.06.2018,Ninebot,10,200,Иван")
    bot.register_next_step_handler(m,Current)
    
        

def Current(m):
    current_list = str(m.text) #Recieve msg
    current_list = current_list.split(',')
    #print(current_list,len(current_list))
    rb = open_workbook(current_list[0] + '.xls')
    wb = copy(rb)
    read = rb.sheet_by_index(0)
    sheet = wb.get_sheet(0) # Работа в текущем Exel-листе
    for i in range(len(current_list)):
        sheet.write(read.nrows, i, current_list[i])
       # print(current_list[i])
    wb.save(current_list[0] + '.xls')
    bot.send_message(m.chat.id, "Добавлено!, молодец")
    doc = open(current_list[0] + '.xls', 'rb')
    bot.send_document(124641536, doc)


if __name__ == '__main__':
    bot.polling(none_stop=True)
