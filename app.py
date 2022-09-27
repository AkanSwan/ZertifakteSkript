import os
import sys
import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename



def get_names():
    with open(askopenfilename(initialdir=os.curdir, filetypes=[("text", "*.txt")]), encoding='utf-8') as file:
        print('Reading Names')
        return file.read().splitlines()


def apply_template(names):
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(askopenfilename(initialdir=os.curdir, filetypes=[('Word', ".docx")]))
    path = os.path.abspath(os.curdir)
    try:
        os.mkdir(f'{path}/pdfs')
    except Exception as e:
        print(e)
    search_name = 'Herr Max Mustermann'
    for name in names:
        print(name)
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
    sys.exit()


window = Tk()
window.after_idle(apply_template(get_names()))
