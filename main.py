from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd
import pathlib
from docx_adoc import *
import os
from tkinter.ttk import Combobox

def callback():
    name = fd.askopenfilename()
    ePath.config(state = 'normal')
    ePath.delete('1', END)
    ePath.insert('1', name)
    ePath.config(state='readonly')


def convert_docx_to_adoc():    
    doc = Document(ePath.get())
    convert(doc)

root = Tk()
root.title('Конверт docx в adoc')
root.geometry('800x600')
root.resizable(width=False, height=False)
root['bg'] = 'black'

Button(root, text = 'Выберите docx файл', font='Arial 25 bold',
        fg='white', bg='black', command = callback).pack(pady=10)

lbPath = Label(root, text='Путь к файлу:', fg='white', bg='black',font='Arial 25 bold')
lbPath.pack()
combo = Combobox(root) 
combo['values'] = (1, 2, 3, 4, 5, "Текст") 
combo.current(0)  # установите вариант по умолчанию  
combo.grid(column=0, row=0)  
ePath = Entry(root, width=50, state = 'readonly')
ePath.pack(pady=10)

btnConvert = Button(root, text='Конвертировать', fg='white', bg='black',font='Arial 25 bold', command = convert_docx_to_adoc).pack(pady=10)

root.mainloop()