from tkinter import Tk, filedialog, Label, Button, Entry
from openpyxl import Workbook
from pathlib import Path
from threading import Thread
from warnings import filterwarnings

window = Tk()
window.title('File_List_Creator')
window.minsize(width=300, height=145)
window.maxsize(width=300, height=145)

def threading_btn2():
    thread_btn2 = Thread(target=func_btn2)
    thread_btn2.start()

def func_btn2():
    global folder
    folder = filedialog.askdirectory()
    ent2.insert(0, folder)

def threading_btn3():
    thread_btn2 = Thread(target=ask_for_filename)
    thread_btn2.start()

def ask_for_filename():
    labl4.config(text='Listing all the files available in folder and subfolder.')
    filterwarnings("ignore", category=DeprecationWarning)
    excel_file = Workbook()
    work_sheet = excel_file.active
    work_sheet.title = 'List of files'
    work_sheet.append(['File-Name', 'Path'])
    for column in work_sheet.columns:
        for cell in column:
            alignment_obj = cell.alignment.copy(horizontal='center', vertical='center')
            cell.alignment = alignment_obj
    current_path = folder
    paths = Path(current_path)
    for path in paths.rglob("*"):
        string_path = str(f"{path}")
        file_name = string_path.split('\\')[-1]
        work_sheet.append([file_name, string_path])
    excel_file.save(f"List.xlsx")
    labl4.config(text='List.xlsx generated successfully in current working directory',wraplength=300, justify="left")


labl1 = Label(window, text='Please select the root folder to create a filelist', font=(None, 10, 'bold')).place(x=0,
                                                                                                                y=4)
labl2 = Label(window, text='Path').place(x=0, y=30)
ent2 = Entry(window, bd=4, width=37)
ent2.place(x=35, y=30)
btn2 = Button(window, text='...', bg='green', command=threading_btn2)
btn2.place(x=273, y=30)

btn3 = Button(window, text='Submit', bg='green', command=threading_btn3)
btn3.place(x=130, y=60)

labl3 = Label(window, text='After submission wait for success message.').place(x=0, y=90)
labl4 = Label(window)
labl4.place(x=0, y=110)
window.mainloop()
