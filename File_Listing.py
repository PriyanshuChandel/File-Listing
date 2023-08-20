from tkinter import Tk, filedialog, Label, Button, Entry, Toplevel
from tkinter.ttk import Progressbar, Style
from openpyxl import Workbook
from pathlib import Path
from threading import Thread
from warnings import filterwarnings
from os.path import join, dirname, isfile, exists
from datetime import datetime

fileHandler = open(f"logs_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt", 'a')

iconFile = join(dirname(__file__), 'listing.ico')
aboutIcon = join(dirname(__file__), 'info.ico')

sourcePath = ''


def threadingFileDialog(pathEntry, fileDialogBtn, submitBtn, messageText, progressBar, progressStyle):
    thread_btn2 = Thread(target=FileDialog, args=(pathEntry, fileDialogBtn, submitBtn, messageText, progressBar,
                                                  progressStyle))
    thread_btn2.start()


def FileDialog(pathEntry, fileDialogBtn, submitBtn, messageText, progressBar, progressStyle):
    global sourcePath
    progressBar.config(value=0)
    progressStyle.configure("Custom.Horizontal.TProgressbar", text='0 %')
    sourcePath = filedialog.askdirectory()
    pathEntry.config(state='normal')
    pathEntry.delete(0, 'end')
    pathEntry.insert(0, sourcePath)
    if len(pathEntry.get()) > 0:
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{sourcePath}] selected as source path\n')
        submitBtn.config(state='normal', bg='green')
        fileDialogBtn.config(state='disabled', bg='light grey')
    else:
        messageText.config(text='Please select source path')
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{sourcePath}] No source path selected\n')
    pathEntry.config(state='disabled')


def threadingFileListing(messageText, submitBtn, fileDialogBtn, progressBar, window, progressStyle):
    thread_btn2 = Thread(target=FileListing, args=(messageText, submitBtn, fileDialogBtn, progressBar, window,
                                                   progressStyle))
    thread_btn2.start()


def FileListing(messageText, submitBtn, fileDialogBtn, progressBar, window, progressStyle):
    submitBtn.config(state='disabled', bg='light grey')
    messageText.config(text='Listing...')
    filterwarnings("ignore", category=DeprecationWarning)
    excelFile = Workbook()
    workSheet = excelFile.active
    workSheet.title = 'List of files'
    fileHandler.write(f'{datetime.now().replace(microsecond=0)}[List of files] worksheet created\n')
    workSheet.append(['File-Name', 'Directory'])
    for column in workSheet.columns:
        for cell in column:
            alignmentObj = cell.alignment.copy(horizontal='center', vertical='center')
            cell.alignment = alignmentObj
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Listing...\n')
    directories = list(Path(sourcePath).rglob("*"))
    totalFiles = len(directories)
    listingSuccess = 0
    fileListed = 0
    directoriesSubdirectories = 0
    if len(directories) > 0:
        for path in directories:
            stringPath = str(f"{path}")
            if isfile(stringPath):
                fileName = stringPath.split('\\')[-1]
                fileListed = fileListed + 1
            else:
                fileName = ''
                directoriesSubdirectories = directoriesSubdirectories + 1
            workSheet.append([fileName, stringPath])
            listingSuccess = listingSuccess + 1
            updateProgress(listingSuccess, progressBar, totalFiles, window, progressStyle)
        excelFileName = "File_List.xlsx"
        excelFIleCounter = 1
        while exists(excelFileName):
            excelFileName = f"File_List_{excelFIleCounter}.xlsx"
            excelFIleCounter = excelFIleCounter + 1
        excelFile.save(excelFileName)
        fileHandler.write(f'{datetime.now().replace(microsecond=0)}Listing Done for\nGrand_Total: [{listingSuccess}]\n'
                          f'Directories/Subdirectories : [{listingSuccess - fileListed}]\n'
                          f'Total_Files : [{fileListed}]\n')
        fileHandler.write(
            f"{datetime.now().replace(microsecond=0)}[{excelFileName}] saved in current directory.\n")
        messageText.config(text=f'Done, {excelFileName} created')
    else:
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{sourcePath}] is empty\n')
        messageText.config(text='Error! check logs')
    fileDialogBtn.config(state='normal', bg='green')


def updateProgress(listingSuccess, progressBar, totalFiles, window, progressStyle):
    resultVal = (listingSuccess / totalFiles) * 100
    progressBar['value'] = resultVal
    progressStyle.configure("Custom.Horizontal.TProgressbar", text='{:g} %'.format(resultVal))
    window.update()


def mainGUI():
    window = Tk()
    window.config(bg='light grey')
    window.title('File Listing v1.0')
    window.geometry('357x235')
    window.resizable(False, False)
    window.iconbitmap(iconFile)
    mainLabel = Label(window, text='File List', font=('Arial', 20, 'bold'), fg='blue', bg='light grey')
    mainLabel.place(x=175, y=20, anchor='center')
    sourcePathLabel = Label(window, text='Path:', font=('Arial', 12, 'bold italic'), bg='light grey').place(x=2, y=60)
    pathEntry = Entry(window, bd=4, width=40, bg='white', state='disabled')
    pathEntry.place(x=60, y=60)
    fileDialogBtn = Button(window, text='...', bg='green', fg='white', font=('Arial', 12),
                           command=lambda: threadingFileDialog(pathEntry, fileDialogBtn, submitBtn, messageLabel,
                                                               progress, progressStyle))
    fileDialogBtn.place(x=320, y=58)
    submitBtn = Button(window, text='Submit', bg='light grey', fg='white', font=('Arial', 12, 'bold'),
                       command=lambda: threadingFileListing(messageLabel, submitBtn, fileDialogBtn, progress, window,
                                                            progressStyle),
                       state='disabled')
    submitBtn.place(x=150, y=98)
    progress = Progressbar(window, length=340, mode="determinate", style="Custom.Horizontal.TProgressbar")
    progress.place(x=5, y=150)
    progressStyle = Style()
    progressStyle.configure("Custom.Horizontal.TProgressbar", thickness=20, troughcolor='gray88', background='light green',
                            troughrelief='flat', relief='flat', text='0 %')
    progressStyle.layout('Custom.Horizontal.TProgressbar', [('Horizontal.Progressbar.trough',
                                                             {'children': [('Horizontal.Progressbar.pbar',
                                                                            {'side': 'left', 'sticky': 'ns'})],
                                                              'sticky': 'nswe'}),
                                                            ('Horizontal.Progressbar.label', {'sticky': ''})])
    messageLabel = Label(window, font=('Arial', 11, 'bold'), bg='light grey')
    messageLabel.place(x=5, y=185)
    aboutBtn = Button(window, text='?', bg='brown', command=lambda: aboutWindow(window))
    aboutBtn.place(x=338, y=207)
    window.mainloop()


def aboutWindow(mainWin):
    aboutWin = Toplevel(mainWin)
    aboutWin.grab_set()
    aboutWin.geometry('285x90')
    aboutWin.resizable(False, False)
    aboutWin.title('About')
    aboutWin.iconbitmap(aboutIcon)
    aboutWinLabel = Label(aboutWin, text=f'Version - 1.0\nDeveloped by Priyanshu\nFor any improvement please reach on '
                                         f'below email\nEmail : chandelpriyanshu8@outlook.com\nMobile : '
                                         f'+91-8285775109 '
                                         f'', font=('Helvetica', 9)).place(x=1, y=6)


mainGUI()
