from tkinter import *
from tkinter.ttk import *
from pyponv2 import get_all_divisions, get_all, get_lesson
import json
import requests


sheet = {}
cells = list()
current_lesson = {}
lessons = list()
lesson = {}

def get_lessons():
    file_name = file_entry.get()
    sheet_name = sheet_entry.get()

    global cells, sheet
    cells, sheet = get_all(file_name, sheet_name)
    cell_quantity_text['text'] = 'Знайдено {} пар'.format(len(cells))
    print_lesson( int(lesson_num_entry.get()) if lesson_num_entry.get() else 0)

    magick_button['state'] = 'normal'
    divisions_button['state'] = 'normal'
    find_button['state'] = 'normal'
    prev_button['state'] = 'normal'
    next_button['state'] = 'normal'
    send_button['state'] = 'normal'

def get_divisions():
    access_token = access_token_entry.get()
    print(access_token)
    if len(access_token) == 0:
        return

    divisions = get_all_divisions(sheet)
    print(divisions)

    for division in divisions:
        x = requests.post('https://timetable.univera.app/division', json={"data": division}, headers={'Authorization': 'Bearer '+access_token})
        print(x.text)

def send_to_server():
    access_token = access_token_entry.get()
    if len(access_token) == 0:
        print('no access token')
        return

    lessons = []
    for cell in cells:
        lessons.append(get_lesson(cell, sheet))

    x = requests.post('https://timetable.univera.app/hui', json={"data": lessons}, headers={'Authorization': 'Bearer '+access_token})
    print(x.text)

def generate():
    global cells, lessons, sheet
    for cell in cells:
        lessons.append(get_lesson(cell, sheet))

    data = json.dumps(lessons)

    f = open(sheet_entry.get(), 'w')
    f.write(data)
    f.close()

def print_lesson(lesson_index):
    cell = cells[lesson_index]
    lesson_num_entry.delete(0, END)
    lesson_num_entry.insert(0, lesson_index)
    lesson_name_label['text'] = 'ЦЕ знайшлося в клітинці {}'.format(str(cell).split('.')[1].replace('>', ''))

    global lesson
    global lesson_type
    lesson = get_lesson(cell, sheet)
    print(lesson)

    lesson_subject['text'] = lesson['subject']
    lesson_link['text'] = lesson['link'] if 'link' in lesson else 'Немає посилання'
    lesson_type['text'] = lesson['type']
    lesson_divisions['text'] = lesson['divisions']
    lesson_lecturers['text'] = lesson['lecturers']
    lesson_repeat['text'] = lesson['repeat']


def onKeyPress(event):
    print('You pressed %s\n' % (event.char, ))

root = Tk()
root.title("This pypon is yours")
root.geometry('1000x600')

label = Label(root, text ='Добро пожаловать в ад! Вставь аксес токен кста').grid(row=0, column=1)
access_token_entry = Entry(root)
access_token_entry.grid(row=0, column=2)

# file params ui
Label(root, text="file path").grid(row=1, column=0)
file_entry = Entry(root, width=50)
file_entry.insert(0,'ari.xlsx')
file_entry.grid(row=1, column=1)

Label(root, text='sheet name').grid(row=2, column=0)
sheet_entry = Entry(root, width=50)
sheet_entry.insert(0,'Лист1')
sheet_entry.grid(row=2, column=1)

Button(root, text='Перезавантажити', command=get_lessons).grid(row=1, column=2)

cell_quantity_text = Label(root)
cell_quantity_text.grid(row=3, column=0)

# lesson ui
#

# Lesson number entry
lesson_num_entry = Entry(root)
lesson_num_entry.grid(row=5, column=0)

# Prev button
prev_button = Button(root, text='<', state='disabled',command=lambda:print_lesson(int(lesson_num_entry.get()) - 1))
prev_button.grid(row=5, column=1)
root.bind('<Left>', lambda event:print_lesson(int(lesson_num_entry.get()) - 1))
# Find button
find_button = Button(root, text='Знайти', state='disabled',command=lambda:print_lesson(int(lesson_num_entry.get())))
find_button.grid(row=5, column=2)
# Next button
next_button = Button(root, text='>', state='disabled',command=lambda:print_lesson(int(lesson_num_entry.get()) + 1))
next_button.grid(row=5, column=3)
root.bind('<Right>', lambda event:print_lesson(int(lesson_num_entry.get()) + 1))

lesson_name_label = Label(root, text='тут шось буде')
lesson_name_label.grid(row=4, column=2)

Label(root, text='Subject').grid(row=6, column=0)
lesson_subject = Label(root)
lesson_subject.grid(row=6, column=1)

# Lesson link label
Label(root, text='Link').grid(row=7, column=0)
lesson_link = Label(root, wraplength=300)
lesson_link.grid(row=7, column=1)
lesson_type = 'lecture'

# Lesson type label
Label(root, text='Type').grid(row=8, column=0)
lesson_type = Label(root)
lesson_type.grid(row=8, column=1)

# Lesson repeat label
Label(root, text='Repeat').grid(row=9, column=0)
lesson_repeat = Label(root)
lesson_repeat.grid(row=9, column=1)

# Lesson divisions label
Label(root, text='Divisions').grid(row=10, column=0)
lesson_divisions = Label(root)
lesson_divisions.grid(row=10, column=1)

# Lesson lecturers label
Label(root, text='Lecturers').grid(row=11, column=0)
lesson_lecturers = Label(root)
lesson_lecturers.grid(row=11, column=1)

magick_button = Button(root, text='generate JSON', state='disabled', command=generate)
magick_button.grid(row=2, column=2)

divisions_button = Button(root, text='Завантажити групи на сервер', width=30, state='disabled', command=get_divisions)
divisions_button.grid(row=3, column=2)

send_button = Button(root, text='Відправити на сервер розклад', state='disabled', command=send_to_server)
send_button.grid(row=2, column=3)

root.mainloop()
