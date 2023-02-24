# List Files and Folders Recursively

This Python program takes a directory path as input and lists all the files and folders present inside it or its subfolders recursively.

### Requirements
- Python 3.x
- `tkinter`: Python's standard GUI package. It provides the `Tk` interface to create GUIs. Used to create the file selection dialog box for selecting the input file and output folder.
- `openpyxl`: A Python library used for working with Excel spreadsheets. Used to create and write data to the output Excel file.
- `pathlib`: A Python library used for working with file system paths. Used to create and validate file paths.
- `threading`: A Python library used to perform parallel execution of the program. Used to keep the GUI responsive while the program is executing.
- `warnings`: A Python library used to handle warning messages. Used to filter out warning messages related to Excel file format.

To install these dependencies, you can use pip, the package installer for Python:
> `pip install tkinter openpyxl pathlib`

Note that `threading` and `warnings` are built-in libraries and do not need to be installed separately.

### Installation and Usage
1. Clone the repository or download the zip file and extract it to your desired location.
2. Open the terminal/command prompt and navigate to the directory where the program is saved.
3. To run the program, enter the following command:
  > `python File_Listing.py`
4. Select the desired directory in GUI by clicking on `...`
5. Click on `Submit` button.
6. The program will start executing and will list all the files and folders present inside the specified directory and its subfolders in file `List/xlsx`.

### Contributions
Contributions to this repo are welcome. If you find a bug or have a suggestion for improvement, please open an issue on the repository. If you would like to make changes to the code, feel free to submit a pull request.

### Acknowledgments
This program was created as a part of a programming challenge. Special thanks to the challenge organizers for the inspiration.
