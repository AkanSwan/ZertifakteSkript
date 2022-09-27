import os
import sys
import win32com.client
import tkinter
from time import sleep

# Press Umschalt+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def get_names():
    try:
        with open(os.path.abspath(os.path.curdir + '/data/names.txt'), encoding='utf-8') as file:
            print('Reading Names')
            return file.read().splitlines()

    except FileNotFoundError as e:
        print(f"""No file with names found. Please create names.txt at {os.path.abspath(os.path.curdir)}/data/""")
        sleep(15)
        sys.exit()


def main(names=[]):
    word = win32com.client.Dispatch('Word.Application')
    path = os.path.abspath(os.curdir)
    doc = word.Documents.Open(path + '/data/template.docx')
    try:
        os.mkdir(f'{path}/pdfs')
    except Exception as e:
        print(e)
    search_name = 'Herr Max Mustermann'
    for name in names:
        word.Selection.Find.Execute(search_name, False, False, False, False, False,
                                    True, 1, False, name, 2)
        try:
            os.remove(f'{path}/pdfs/Zertifikat_{name}.pdf')
        except Exception as e:
            print(e)
        doc.SaveAs2(f'{path}/pdfs/Zertifikat_{name}.pdf', FileFormat=17)
        print(f'Created for {name}')
        search_name = name
    print('Done')
    doc.Close(0)
    word.Quit()
    sleep(15)
    sys.exit()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main(get_names())

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
