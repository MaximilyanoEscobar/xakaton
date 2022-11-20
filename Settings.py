from aiogram import types
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

BOT_TOKEN = '5735681613:AAHxRfOOKeW5XxMwdG3mQmSOyBVxnLHqp9M'
kb_mm = ['Заполнить заявку', 'Справка']
MainMenu = types.ReplyKeyboardMarkup(resize_keyboard=True).add(*kb_mm)
kb_am = ['Список вопросов', 'Добавить вопрос', 'Изменить вопрос', 'Удалить вопрос', 'Главное меню']
AdminMenu = types.ReplyKeyboardMarkup(resize_keyboard=True).add(*kb_am)
GoToQuestions = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="Да", callback_data="Yes"), InlineKeyboardButton(text="Нет", callback_data="No")
        ]
    ])
kb_sa = ['Отменить действие']
AdminStopKeyboard = types.ReplyKeyboardMarkup(resize_keyboard=True).add(*kb_sa)
