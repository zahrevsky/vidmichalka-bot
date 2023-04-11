import os
from datetime import datetime, date
from threading import Thread
import time
from logging import getLogger, StreamHandler

from schedule import run_pending
import telebot
import gspread
import yaml


with open("config.yaml") as inp:
    config = yaml.safe_load(inp)
    CLIENTS = config['clients']
    ADMIN_ID = config['admin_id']
    OPTIONS = config['options']

bot = telebot.TeleBot(os.environ.get('TELEGRAM_CREDS'))
gc = gspread.service_account(filename=os.environ.get('GSHEETS_CREDS_PATH'))
logger = getLogger(__name__)
logger.setLevel('INFO')
logger.addHandler(StreamHandler())
created_polls = {}


class Client:
    def __init__(self, title, spreadsheet_title, chat_id, students, head, group_address):
        self.title = title
        self.spreadsheet_title = spreadsheet_title
        self.chat_id = chat_id
        self.students = students
        self.head = head
        self.group_address = group_address
    
    def _lecture_cell(self, day, number):
        logger.info(f"Finding lecture cell for day {day} and number {number}")

        sheet = self.attendance_sheet(day)
        date_column = sheet.find(day.strftime("%d.%m.%Y"))

        # Lecture cell is `number` columns to the right and one row down from date
        lecture_cell = sheet.cell(date_column.row + 1, date_column.col + number - 1)

        logger.info(
            f"Lecture cell is row {lecture_cell.row}, "
            f"column {column_idx_to_letter(lecture_cell.col)}"
        )

        return lecture_cell

    def mark_student_presence(self, student_name, lecture_date, lecture_number, mark):
        logger.info(f"Marking student {student_name} as {mark}")

        sheet = self.attendance_sheet(lecture_date)

        # Find the student row
        row_idx = sheet.find(student_name).row
        # Find the lecture column
        col_idx = self._lecture_cell(lecture_date, lecture_number).col

        # Mark the cell
        sheet.update_cell(row_idx, col_idx, mark)

        logger.info(
            f"Marked student {student_name} as {mark} in sheet {sheet.title} "
            f"at row {row_idx}, column {column_idx_to_letter(col_idx)}"
        )
    
    def attendance_sheet(self, lecture_date):
        sh = gc.open(self.spreadsheet_title)

        months_sheet_titles = {
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
        month = lecture_date.strftime("%B")
        sheet_title = months_sheet_titles[month]

        sheet = sh.worksheet(sheet_title)
        return sheet
    
    def lecture_title(self, day, number):
        lecture_title = self._lecture_cell(day, number).value

        if lecture_title is None:
            logger.error(
                f"Could not find lecture title for day "
                f"{day} ({day.strftime('%A')}) and number {number}"
            )
        
        return lecture_title
    

def column_idx_to_letter(col_number):
    col_idx = col_number - 1
    if col_idx < 26:
        return chr(ord('A') + col_idx)
    else:
        return column_idx_to_letter(col_idx // 26) + chr(ord('A') + col_idx % 26)


# @repeat(every().monday.at('09:20'), 1)
# @repeat(every().monday.at('11:20'), 2)
# @repeat(every().monday.at('13:05'), 3)
# @repeat(every().tuesday.at('09:25'), 1)
# @repeat(every().tuesday.at('11:20'), 2)
# @repeat(every().tuesday.at('13:05'), 3)
# @repeat(every().wednesday.at('09:25'), 1)
# @repeat(every().wednesday.at('11:20'), 2)
# @repeat(every().wednesday.at('13:05'), 3)
def create_poll_for_client(client, lecture_number, day=None):
    logger.info(f"Creating a poll...")

    if day is None:
        day = date.today()

    lecture = client.lecture_title(day, lecture_number).lower()
    question = \
        f"{client.group_address}, " \
        f"відмічаємося на {lecture_number} парі: " \
        f"{lecture}" \

    response = bot.send_poll(
        chat_id=client.chat_id, 
        question=question,
        options=[option['text'] for option in OPTIONS], 
        is_anonymous=False, 
        type='quiz', 
        correct_option_id=0
    )

    created_polls[response.poll.id] = (client, day, lecture_number)

    print(f'Created poll with id {response.poll.id}')


@bot.poll_answer_handler()
def handle_poll_response(poll_answer):
    logger.info(
        f"Received a response: {poll_answer}"
    )

    poll_id = poll_answer.poll_id
    person_id = poll_answer.user.id

    if poll_id not in created_polls:
        print(
            f'Skipping response to poll, because this poll is not registered. '
            f'Poll answer details: {poll_answer}'
        )
        return

    client, lecture_date, lecture_number = created_polls[poll_id]
    response = poll_answer.option_ids[0]

    students = {
        student['tg_id']: student['name']
        for student in client.students
    }
    if person_id not in students:
        print(
            f'Skipping response for lecture #{lecture_number} '
            f'on {lecture_date}, because respondent is not a group student. '
            f'Poll answer details: {poll_answer}'
        )
        return
    student_name = students[person_id]

    client.mark_student_presence(
        student_name, lecture_date, lecture_number, OPTIONS[response]['emoji']
    )


@bot.message_handler(commands=['create_poll'])
def handle_poll_creation(message):
    if message.from_user.id != ADMIN_ID:
        logger.info(
            f"Received a /create_poll from unauthorized user: {message}"
        )
        return
    logger.info(f"Received a /create_poll: {message}")

    clientname, lecture_number, year, month, day = message.text.split()[1:]
    client = Client(**[c for c in CLIENTS if c['title'] == clientname][0])
    create_poll_for_client(
        client, int(lecture_number), day=date(int(year), int(month), int(day))
    )


@bot.message_handler(func=lambda m: True)
def handle_message(message):
    logger.info(f"Received a message: {message}")


def do_schedule():
    logger.info('Starting schedule job...')
    while True:
        run_pending()
        print(f'Tick at {datetime.now()}')
        time.sleep(1)


if __name__ == '__main__':
    logger.info('Starting bot...')
    # thread = Thread(target=do_schedule)
    # thread.start()

    # test_client = Client(**CLIENTS[0]) #TODO: get client by title from list in clients
    # create_poll_for_client(test_client, 1, day=date(2023, 4, 4))

    logger.info('Starting polling...')
    bot.infinity_polling()
