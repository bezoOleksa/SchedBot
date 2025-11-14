# In the process of rewriting. *Non-working version*

import os
import sys
import requests
import json
import openpyxl
import time
import random
# import logging


ADMIN = os.getenv('TELEGRAM_BOT_ADMIN')
TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
API_URL = 'https://api.telegram.org/bot'+TOKEN


startMessage = 'Вітаю у боті для розкладу занять! 👋\n\n' \
             + 'Я буду надсилати вам автоматичні повідомлення на початку та в кінці уроку\n' \
             + 'Щоб вимкнути/увімкнути сповіщення, скористайтесь командами /mute та /unmute\n\n' \
             + 'Для отримання розкладу на сьогодні, надішліть /today\n' \
             + 'Для розкладу на завтра, надішліть /tomorrow\n\n' \
             + 'Для початку, давайте налаштуємо ваш розклад.'

helpMessage = startMessage  # needs to be done!

askRole = 'Будь ласка, оберіть вашу роль:'
keyboardRole = {'keyboard': [[{'text': 'Учень'}, {'text': 'Вчитель'}]], 
                'resize_keyboard': True, 'one_time_keyboard': True}
answerRole = ('учень', 'вчитель')

askGrade = 'Чудово! Тепер оберіть ваш клас:'
keyboardGrade = {'keyboard': [[{'text': '9'}, {'text': '10'}, {'text': '11'}]], 
                 'resize_keyboard': True, 'one_time_keyboard': True}

askGroup = 'Будь ласка, оберіть свою групу:'
groups = {'9': ['М-21', 'ІФ-22', 'ОІ-23', 'КМ-24', 'ПА-25'],
          '10': ['М-31', 'ІФ-32', 'ІЮ-33', 'КМ-34', 'ОІФ-35'],
          '11': ['М-41', 'ІФ-42', 'ПМ-43', 'ІН-ІФ-44']}

askHalf = 'І останнє, оберіть вашу підгрупу:'
keyboardHalf = {'keyboard': [[{'text': '1'}, {'text': '2'}]], 
                'resize_keyboard': True, 'one_time_keyboard': True}

finalMessage = 'Дякую! Налаштування завершено. Тепер ви будете отримувати повідомлення з вашим персональним розкладом.'
teacherNote = 'Наразі функціонал для вчителів ще не розроблено, але ви все одно можете отримувати повідомлення про початок і кінець уроку. Дякуємо за розуміння!'
unrecognizedMessage = 'Вибачте, я не зрозумів вашу відповідь. Будь ласка, скористайтесь кнопками або командами. Для допомоги надішліть /help.'
scheduleSetupError = 'Будь ласка, спершу завершіть налаштування за допомогою команди /start, щоб отримати персональний розклад.'
weekendMessage = 'Сьогодні вихідний! Відпочивайте. 🥳'

muteAnswer = '✅ Автоматичні сповіщення вимкнено. Щоб увімкнути їх знову, скористайтесь командою /unmute.'
unmuteAnswer = '✅ Автоматичні сповіщення увімкнено! Щоб вимкнути їх, скористайтесь командою /mute.'

lessonStartMessage = '🔔 Початок уроку: '
breakMessages = [
    'Час для короткого відпочинку!',
    'Відновлюй сили, попереду нові знання.',
    'Зроби перерву, ти на це заслуговуєш.',
    'Кілька хвилин для себе.',
    'Переключись на щось приємне.',
    'Час для чаю або кави!',
    'Розслабся, скоро продовжимо.',
    'Невеличка пауза для великих звершень.'
]
endOfDayMessages = [
    'Це був останній урок на сьогодні! Вітаємо, ви впорались! 🎉',
    'Навчальний день завершено! Час відпочивати. ✨',
    'Уроки скінчились! Ви молодці! 👍',
    'Ще один день позаду! Гарного вечора!',
    'Ви чудово попрацювали! Тепер час для відпочинку.',
    'На сьогодні все! Набирайтесь сил на завтра.'
]


def loadFiles():
    global lastUpdate, Timetable
    global Users
    try:
        with open('config.json', 'r', encoding='ascii') as file:
            data = json.load(file)
        lastUpdate = data['lastUpdate']
        Timetable = data['Timetable']
        print('Successfully loaded config.json')

        with open('users.json', 'r', encoding='utf-8') as file:
            Users = json.load(file)
        print('Successfully loaded users.json')

    except Exception as e:
        print('An error occurred while loading files (config.json, users.json):', e)
        sys.exit()


def makeSchedule():
    schedule = {}
    lessonsPerDay = 8
    daysPerWeek = 6
    startingRow = 13

    try:
        excelFile = openpyxl.load_workbook('./schedule.xlsx')
        scheduleExcel = excelFile.active
        isMerged = lambda cell: isinstance(cell, openpyxl.cell.cell.MergedCell)
        groupNames = groups['9'] + groups['10'] + groups['11']  # grades 9, 10 and 11 possible only

        for group in range(2, 2 + len(groupNames) * 2):  # startingColumn = 2
            name = groupNames[group // 2 - 1]
            if name not in schedule:
                schedule[name] = [[], []]  # subgroups

            for day in range(daysPerWeek):
                schedule[name][group % 2].append([])
                prevCell = None

                for lesson in range(lessonsPerDay):
                    rowInExcel = startingRow + day*(lessonsPerDay+1) + lesson
                    lessonCell = scheduleExcel.cell(row=rowInExcel, column=group)
                    if isMerged(lessonCell):
                        lessonCell = scheduleExcel.cell(row=rowInExcel, column=group-1)

                    if lessonCell.value or not prevCell:
                        schedule[name][group % 2][day].append(lessonCell.value)
                        prevCell = lessonCell.value

        with open('schedule.json', 'w', encoding='utf-8') as scheduleFile:
            json.dump(schedule, scheduleFile, indent=2, ensure_ascii=False)
        return schedule

    except Exception as e:
        print('Exception occured while making schedule:', e)
        sys.exit()

# Schedule = makeSchedule()

def sendMessage(chatID, text, keyboard={}):
    params = {'chat_id': chatID, 'text': text}
    if keyboard:
        params['reply_markup'] = json.dumps(keyboard)
    try:
        send = requests.post(API_URL + '/sendMessage', params=params, timeout=10)
        send.raise_for_status()
    except requests.exceptions.RequestException as e:
        print('Error ocurred sending message. Error:', e, end='; ')
        time.sleep(0.4 + random.random() / 2)
        try:
            send = requests.post(API_URL + '/sendMessage', params=params, timeout=10)
            send.raise_for_status()
            print('Sent on second attempt')
        except requests.exceptions.RequestException:
            print('No success on second attempt')


def uploadSchedule(document):
    global Schedule
    if document.get('file_name') != 'schedule.xlsx':
        sendMessage(ADMIN, 'Please upload file named schedule.xlsx to update schedule')
        return

    try:
        os.rename('schedule.xlsx', 'schedule_backup.xlsx')
        fileID = document.get('file_id')
        getFile = requests.get(API_URL + '/getFile', params={'file_id': fileID})
        getFile.raise_for_status()
        filePath = getFile.json()['result']['file_path']

        newFile = requests.get(f'https://api.telegram.org/file/bot{TOKEN}/{filePath}')
        with open('schedule.xlsx', 'wb') as file:
            for chunk in newFile.iter_content(chunk_size=8192):
                if chunk:
                    file.write(chunk)

        Schedule = makeSchedule()
        sendMessage(ADMIN, 'Successfully updated schedule.xlsx file')
        os.remove('schedule_backup.xlsx')

    except Exception as e:
        os.remove('schedule.xlsx')
        os.rename('schedule_backup.xlsx','schedule.xlsx')
        print('Error uploading schedule:', e)
        sendMessage(ADMIN, f'Error uploading schedule: {e}')


def notify():
    pass

def reactToMessage(update):
    if 'message' not in update:
        return

    chatID = update['message']['chat']['id']

    if 'document' in update['message'] and chatID == ADMIN:
        uploadSchedule(update['message']['document'])
        return

    if 'text' not in update['message']:
        return
    text = update['message']['text']

    if text == '/start' or chatID not in Users:
        Users[chatID] = {'sendAuto': True, 'stage': 0}
        sendMessage(chatID, startMessage)
        sendMessage(chatID, askRole, keyboard=keyboardRole)
        Users[chatID]['stage'] = 1

    elif text == '/help':
        sendMessage(chatID, helpMessage)

    elif text == '/mute':
        Users[chatID]['sendAuto'] = False
        sendMessage(chatID, muteAnswer)

    elif text == '/unmute':
        Users[chatID]['sendAuto'] = True
        sendMessage(chatID, unmuteAnswer)

    elif text == '/sched' or text == '/today':
        pass

    elif text == '/tomorrow':
        pass

    else:
        stage = Users.get(chatID, {}).get('stage')

        if stage == 1 and text.lower() in answerRole:
            if text.lower() == answerRole[0]:
                Users[chatID]['role'] = 'student'
                sendMessage(chatID, askGrade, keyboard=keyboardGrade)
                Users[chatID]['stage'] = 2
            else:
                Users[chatID]['role'] = 'teacher'
                sendMessage(chatID, teacherNote)
                sendMessage(chatID, askRole, keyboard=keyboardRole)
                Users[chatID]['stage'] = 1

        elif stage == 2 and text in groups:
            Users[chatID]['grade'] = text
            keyboardGroup = {'keyboard': [[{'text': group} for group in groups[text]]],
                             'resize_keyboard': True, 'one_time_keyboard': True}
            sendMessage(chatID, askGroup, keyboard=keyboardGroup)
            Users[chatID]['stage'] = 3

        elif stage == 3 and text in groups.get(Users[chatID].get('grade'), []):
            Users[chatID]['group'] = text
            sendMessage(chatID, askHalf, keyboard=keyboardHalf)
            Users[chatID]['stage'] = 4

        elif stage == 4 and text in ('1', '2'):
            Users[chatID]['half'] = int(text) - 1
            sendMessage(chatID, finalMessage)
            Users[chatID]['stage'] = 5

        else:
            sendMessage(chatID, unrecognizedMessage)
