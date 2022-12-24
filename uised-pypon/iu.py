from tkinter import *
from tkinter.ttk import *
from pypon import *

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
    print_lesson(0)

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

    lesson_subject.delete(0, END)
    lesson_subject.insert(0, lesson['subject'])

    lesson_link.delete(0, END)
    lesson_link.insert(0, lesson['link'])
    lesson_type=lesson['type']

root = Tk()
root.title("This pypon is yours")
root.geometry('600x400')

label = Label(root, text ='Добро пожаловать в ад! Вставь аксес токен кста').grid(row=0, column=1)
access_token_entry = Entry(root)
access_token_entry.grid(row=0, column=2)

# file params ui
Label(root, text="file path").grid(row=1, column=0)
file_entry = Entry(root)
file_entry.insert(0,'routine-ics-2course.xlsx')
file_entry.grid(row=1, column=1)

Label(root, text='sheet name').grid(row=2, column=0)
sheet_entry = Entry(root)
sheet_entry.insert(0,'Лист1')
sheet_entry.grid(row=2, column=1)

Button(root, text='хуярь', command=get_lessons).grid(row=1, column=2)

cell_quantity_text = Label(root)
cell_quantity_text.grid(row=3, column=0)

# lesson ui
lesson_num_entry = Entry(root)
lesson_num_entry.grid(row=4, column=0)
Button(root, text='Знайти', command=lambda:print_lesson(int(lesson_num_entry.get()))).grid(row=4, column=1)

lesson_name_label = Label(root, text='тут шось буде')
lesson_name_label.grid(row=4, column=2)

Label(root, text='Subject').grid(row=5, column=0)
lesson_subject = Entry(root)
lesson_subject.grid(row=5, column=1)
Label(root, text='Link').grid(row=6, column=0)
lesson_link = Entry(root)
lesson_link.grid(row=6, column=1)
lesson_type = 'lecture'

lesson_type_lecture = Radiobutton(root, text='lecture', variable=lesson_type, value='lecture') 
lesson_type_lecture.grid(row=7, column=0)

lesson_type_practice = Radiobutton(root, text='practice', variable=lesson_type, value='practice') 
lesson_type_practice.grid(row=8, column=0)

Button(root, text='magic', command=generate).grid(row=10, column=2)

root.mainloop()
