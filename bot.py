import os
from datetime import datetime, date
from threading import Thread
import time


from schedule import every, repeat, run_pending
from telebot import types
import telebot
import gspread


def google_acc():
    creds_path = os.environ.get('GSHEETS_CREDS_PATH')
    gc = gspread.service_account(filename=creds_path)
    return gc


group_id = -1001429649964
bot = telebot.TeleBot(os.environ.get('TELEGRAM_CREDS'))
gc = google_acc()


created_polls = {}


@repeat(every().monday.at('08:55'), 1)
@repeat(every().monday.at('10:50'), 2)
@repeat(every().monday.at('12:35'), 3)
@repeat(every().tuesday.at('08:55'), 1)
@repeat(every().tuesday.at('10:50'), 2)
@repeat(every().tuesday.at('12:35'), 3)
@repeat(every().wednesday.at('08:55'), 1)
@repeat(every().wednesday.at('11:55'), 2) # TODO: change to 10:50
@repeat(every().wednesday.at('12:35'), 3)
def create_poll(lecture_number):
    question = {
        1: "Перша пара, відмічаємося",
        2: "Друга пара, відмічаємося",
        3: "Третя пара, відмічаємося"
    }[lecture_number]

    response = bot.send_poll(
        chat_id=group_id, 
        question=question, 
        options=["Був", "Не був"], 
        is_anonymous=False, 
        type='quiz', 
        correct_option_id=0
    )

    created_polls[response.poll.id] = (date.today(), lecture_number)

    print(f'Created poll with id {response.poll.id}')


@bot.poll_answer_handler()
def handle_poll_response(poll_answer):
    poll_id = poll_answer.poll_id
    person_id = poll_answer.user.id

    if poll_id not in created_polls:
        print(f'Skipping response to poll, beacause this poll is not registered. Poll answer details: {poll_answer}')
        return

    lecture_date, lecture_number = created_polls[poll_id]
    response = poll_answer.option_ids[0]

    students = {
        397549160: "Андрій Клебанов",
        181083338: "Влад Загревський",
        582658338: "Юра Андріїв",
        693506992: "Настя Трачук",
        367235871: "Максим Бут"
    }
    if person_id not in students:
        print(f'Skipping response for lecture #{lecture_number} on {lecture_date}, beacause respondent is not a group student. Poll answer details: {poll_answer}')
        return
    student_name = students[person_id]

    if response == 0:
        mark_student_presence(student_name, lecture_date, lecture_number)


def mark_student_presence(student_name, lecture_date, lecture_number):
    # Open the spreadsheet
    sh = gc.open("Відвідуваність астрономів, 1 маг 22/23")
    months = {
        'January': 'Січень',
        'February': 'Лютий',
        'March': 'Березень',
        'April': 'Квітень',
        'May': 'Травень',
        'June': 'Червень',
        'July': 'Липень',
        'August': 'Серпень',
        'September': 'Вересень',
        'October': 'Жовтень',
        'November': 'Листопад',
        'December': 'Грудень'
    }
    month = months[lecture_date.strftime("%B")]
    sheet = sh.worksheet(month)

    # Find the student row
    student_row = sheet.find(student_name)
    student_row_index = student_row.row

    # Find the lecture column
    date_column = sheet.find(lecture_date.strftime("%d.%m.%Y"))
    date_col_index = date_column.col
    lecture_col_index = date_col_index + lecture_number - 1

    # Mark the cell as true to indicate presence
    sheet.update_cell(student_row_index, lecture_col_index, 'TRUE')


def do_schedule():
    while True:
        run_pending()
        time.sleep(1)


if __name__ == '__main__':
    thread = Thread(target=do_schedule)
    thread.start()

    print('Starting polling...')
    bot.infinity_polling()
