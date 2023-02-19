from threading import Thread
from tkinter import filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap import Frame
from ttkbootstrap import Menu
from ttkbootstrap import Label
from ttkbootstrap.dialogs.dialogs import MessageDialog
from ttkbootstrap import Entry
from visedit import StringEdit
import yaml


class App(Frame):

    def __init__(self, master=None, event=None):

        super().__init__(master)
        self.master.geometry("1200x800")

        self.master.protocol("WM_DELETE_WINDOW", self.click_exit)

        self.filepath = ''
        self.data = yaml.safe_dump(
            {'line1_text': '\n', 'line1_title': '',
            'line2_text': '\n', 'line2_title': '',
            'line3_text': '\n', 'line3_title': ''}
        )

        stre1 = ''
        stre2 = ''
        stre3 = ''
        strst1 = ''
        strst2 = ''
        strst3 = ''

        # メニューバーの作成
        menubar = Menu()

        menu_file = Menu(menubar, tearoff=False)
        menu_file.add_command(label='ファイルを開く', command=self.menu_file_open_click, accelerator="Ctrl+O")
        menu_file.add_command(label='上書き保存', command=self.menu_file_over_write_save_click, accelerator='Ctrl+S')
        menu_file.add_command(label='名前をつけて保存', command=self.menu_file_save_as_click, accelerator='Ctrl+Shift+S')
        menu_file.add_separator()
        menu_file.add_command(label='終了', command=self.click_exit)
        menu_file.bind_all('<Control-o>', self.menu_file_open_click)
        menu_file.bind_all('<Control-s>', self.menu_file_over_write_save_click)
        menu_file.bind_all('<Control-S>', self.menu_file_save_as_click)
        menubar.add_cascade(label='ファイル', menu=menu_file)

        self.master.config(menu = menubar)

        f1 = Frame(self.master, padding=5)
        self.st1 = ScrolledText(f1, padding=5, autohide=True, wrap=CHAR, font=('', 10))
        self.st2 = ScrolledText(f1, padding=5, autohide=True, wrap=CHAR, font=('', 10))
        self.st3 = ScrolledText(f1, padding=5, autohide=True, wrap=CHAR, font=('', 10))
        self.st1.insert(0., strst1)
        self.st2.insert(0., strst2)
        self.st3.insert(0., strst3)

        f2 = Frame(self.master, height=50, padding=5)
        f2_1 = Frame(f2, height=50, padding=2)
        f2_2 = Frame(f2, height=50, padding=2)
        f2_3 = Frame(f2, height=50, padding=2)
        self.e1 = Entry(f2_1)
        self.e2 = Entry(f2_2)
        self.e3 = Entry(f2_3)
        self.e1.insert(END, stre1)
        self.e2.insert(END, stre2)
        self.e3.insert(END, stre3)

        f3 = Frame(self.master, height=20, padding=10)
        f3.pack(fill=X)

        f4 = Frame(self.master, height=40, padding=5)

        f2.pack(fill=X)
        f2_1.place(relx=0.0, rely=0.0, relwidth=0.1, relheight=1.0)
        f2_2.place(relx=0.1, rely=0.0, relwidth=0.6, relheight=1.0)
        f2_3.place(relx=0.7, rely=0.0, relwidth=0.3, relheight=1.0)
        self.e1.pack(fill=BOTH, expand=True, padx=3)
        self.e2.pack(fill=BOTH, expand=True, padx=3)
        self.e3.pack(fill=BOTH, expand=True, padx=3)
        # self.e1.place(relx=0.0, rely=0.0, relwidth=0.1, relheight=1.0)
        # self.e2.place(relx=0.1, rely=0.0, relwidth=0.6, relheight=1.0)
        # self.e3.place(relx=0.7, rely=0.0, relwidth=0.3, relheight=1.0)

        f1.pack(fill=BOTH, expand=YES)
        self.st1.place(relx=0.0, rely=0.0, relwidth=0.1, relheight=1.0)
        self.st2.place(relx=0.1, rely=0.0, relwidth=0.6, relheight=1.0)
        self.st3.place(relx=0.7, rely=0.0, relwidth=0.3, relheight=1.0)

        self.e1.focus_set()

    def menu_file_open_click(self, event=None):
        self.filepath = filedialog.askopenfilename(
            title = '編集ファイルを選択',
            initialdir = './'
        )
        print(self.filepath)

        if self.filepath:
            with open(self.filepath, mode='rt', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                self.data = data
                self.e1.delete('0',END)
                self.e2.delete('0',END)
                self.e3.delete('0',END)
                self.e1.insert(END,data['line1_title'])
                self.e2.insert(END,data['line2_title'])
                self.e3.insert(END,data['line3_title'])
                self.st1.delete('1.0',END)
                self.st2.delete('1.0',END)
                self.st3.delete('1.0',END)
                self.st1.insert(END,data['line1_text'])
                self.st2.insert(END,data['line2_text'])
                self.st3.insert(END,data['line3_text'])
            self.master.title(f'ThreeCrows - {self.filepath}')

    def menu_file_save_as_click(self, event=None):
        self.filepath = filedialog.asksaveasfilename(
            title='名前をつけて保存',
            initialdir='./',
            initialfile='noname',
            filetypes=[('ThreeCrowsプロジェクトファイル', '.tcs'), ('YAMLファイル', '.yaml')],
            defaultextension='tcs'
        )
        print(self.filepath)
        self.save_file(self.filepath)

    def menu_file_over_write_save_click(self, event=None):
        if self.filepath:
            self.save_file(self.filepath)
        else:
            self.menu_file_save_as_click()

    def save_file(self, path):
        with open(path, mode='wt', encoding='utf-8') as f:
            stre1, stre2, stre3 = self.e1.get(), self.e2.get(), self.e3.get()
            strst1, strst2, strst3 = self.st1.get('1.0', END), self.st2.get('1.0', END), self.st3.get('1.0', END)
            d = {'line1_text': strst1, 'line1_title': stre1, 'line2_text': strst2, 'line2_title': stre2, 'line3_text': strst3, 'line3_title': stre3}
            self.data = yaml.safe_dump(d)
            print(self.data)
            yaml.safe_dump(d, f, encoding='utf-8', allow_unicode=True)

    def click_exit(self):

        stre1, stre2, stre3 = self.e1.get(), self.e2.get(), self.e3.get()
        strst1, strst2, strst3 = self.st1.get('1.0', END), self.st2.get('1.0', END), self.st3.get('1.0', END)
        d = {'line1_text': strst1, 'line1_title': stre1, 'line2_text': strst2, 'line2_title': stre2, 'line3_text': strst3, 'line3_title': stre3}
        y = yaml.safe_dump(d)

        if self.data == y:
            self.master.destroy()
        elif self.data != y:
            mg = MessageDialog('更新内容が保存されていません。アプリを終了しますか。', title='ThreeCrows-終了確認', buttons=['キャンセル', '保存せず終了:danger', '保存して終了:success'])
            mg.show()
            if mg.result == '保存して終了':
                self.menu_file_over_write_save_click()
                self.master.destroy()
            elif mg.result == '保存せず終了':
                self.master.destroy()
            else:
                pass


if __name__ == "__main__":
    root = ttk.Window(title='ThreeCrows', themename='darkly', minsize=(500,300))
    app = App(master=root)
    app.mainloop()


