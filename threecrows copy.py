from threading import Thread
from tkinter import filedialog
from tkinter import *
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap import Frame
from ttkbootstrap import LabelFrame
from ttkbootstrap import Menu
from ttkbootstrap import Label
from ttkbootstrap.dialogs.dialogs import MessageDialog
from ttkbootstrap import Entry
from visedit import StringEdit
import yaml
from icecream import ic


class Main(Frame):

    def __init__(self, master=None, event=None):

        super().__init__(master)
        self.master.geometry("1200x800")

        self.master.protocol("WM_DELETE_WINDOW", lambda:self.file_close(shouldExit=True))

        self.settings = {}
        with open(r'E:\Documents\GitHub\ThreeCrows\settings.yaml', mode='rt', encoding='utf-8') as f:
            self.settings = yaml.safe_load(f)

        self.filepath = ''
        self.data = {'line1_text': '', 'line1_title': '', 'line2_text': '', 'line2_title': '', 'line3_text': '', 'line3_title': ''}

        stre1 = ''
        stre2 = ''
        stre3 = ''
        strst1 = ''
        strst2 = ''
        strst3 = ''

        # メニューバーの作成
        menubar = Menu()

        # メニュバー - ファイル
        self.menu_file = Menu(menubar, tearoff=False)

        self.menu_file.add_command(label='ファイルを開く', command=self.file_open, accelerator="Ctrl+O")
        self.menu_file.add_command(label='上書き保存', command=self.file_over_write_save, accelerator='Ctrl+S')
        self.menu_file.add_command(label='名前をつけて保存', command=self.file_save_as, accelerator='Ctrl+Shift+S')

        self.menu_file_recently = Menu(self.menu_file, tearoff=False)


        try:
            self.recently_files: list = self.settings['recently_files']
            self.menu_file_recently.add_command(label=self.recently_files[0], command=lambda:self.file_open(file=self.recently_files[0]), accelerator='Ctrl+R')
            self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file=self.recently_files[1]))
            self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file=self.recently_files[2]))
            self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file=self.recently_files[3]))
            self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file=self.recently_files[4]))
        except IndexError as e:
            pass

        self.open_last_file_id = self.menu_file_recently.bind_all('<Control-r>', self.open_last_file)
        self.menu_file.add_cascade(label='最近使用したファイル', menu=self.menu_file_recently)

        self.menu_file.add_separator()
        self.menu_file.add_command(label='終了', command=lambda:self.file_close(shouldExit=True))
        self.menu_file.bind_all('<Control-o>', self.file_open)
        self.menu_file.bind_all('<Control-s>', self.file_over_write_save)
        self.menu_file.bind_all('<Control-S>', self.file_save_as)
        menubar.add_cascade(label='ファイル', menu=self.menu_file)

        self.master.config(menu = menubar)

        # 各パーツを製作
        f1 = Frame(self.master, padding=5)

        f1_1 = Frame(f1, padding=5)
        f1_2 = Frame(f1, padding=5)
        f1_3 = Frame(f1, padding=5)

        self.e1 = Entry(f1_1)
        self.e2 = Entry(f1_2)
        self.e3 = Entry(f1_3)
        self.e1.insert(END, stre1)
        self.e2.insert(END, stre2)
        self.e3.insert(END, stre3)

        self.st1 = ScrolledText(f1_1, padding=3, autohide=True, wrap=CHAR, font=('', 10))
        self.st2 = ScrolledText(f1_2, padding=3, autohide=True, wrap=CHAR, font=('', 10))
        self.st3 = ScrolledText(f1_3, padding=3, autohide=True, wrap=CHAR, font=('', 10))
        self.st1.insert(0., strst1)
        self.st2.insert(0., strst2)
        self.st3.insert(0., strst3)

        # statusbar要素を作成
        self.sb_elements = []
        ## 文字数カウンター
        def st1_letter_count():
            title = self.e1.get() if self.e1.get() else '左'
            text = '{0}: {1}文字'.format(title, self.letter_count(obj=self.st1)) + ' '*30
            return ('label', text)
        def st2_letter_count():
            title = self.e2.get() if self.e2.get() else '中央'
            text = '{0}: {1}文字'.format(title, self.letter_count(obj=self.st2)) + ' '*30
            return ('label', text)
        def st3_letter_count():
            title = self.e3.get() if self.e3.get() else '右'
            text = '{0}: {1}文字'.format(title, self.letter_count(obj=self.st3)) + ' '*30
            return ('label', text)
        ## 行数カウンター
        def st1_line_count():
            title = self.e1.get() if self.e1.get() else '左'
            text = '{0}: {1}行'.format(title, self.line_count(obj=self.st1)) + ' '*30
            return ('label', text)
        def st2_line_count():
            title = self.e2.get() if self.e2.get() else '中央'
            text = '{0}: {1}行'.format(title, self.line_count(obj=self.st2)) + ' '*30
            return ('label', text)
        def st3_line_count():
            title = self.e3.get() if self.e3.get() else '右'
            text = '{0}: {1}行'.format(title, self.line_count(obj=self.st3)) + ' '*30
            return ('label', text)
        ## ショートカットキー
        self.hotkeys1 = ('label', '[Ctrl+O]: 開く  [Ctrl+S]: 上書き保存  [Ctrl+Shift+S]: 名前をつけて保存  [Ctrl+R]: 最後に使ったファイルを開く（起動直後のみ）')
        self.hotkeys2 = ('label', '[Ctrl+^]: 検索・置換  [Ctrl+Z]: 取り消し  [Ctrl+Shift+Z]: 取り消しを戻す')
        self.hotkeys3 = ('label', '[Ctrl+D][Ctrl+←]: 左に移る  [Ctrl+F][Ctrl+→]: 右に移る')
        ## 顔文字
        self.kaomoji = ''
        ## ツールボタン
        self.toolbutton_open = ('button', '開く', self.file_open)
        self.toolbutton_over_write_save = ('button', '上書き保存', self.file_over_write_save)
        self.toolbutton_save_as = ('button', '名前をつけて保存', self.file_save_as)

        # ステータスバー作成メソッド
        def make_toolbar_element(elementtype=None, text=None, commmand=None):
            if elementtype == 'label':
                return Label(text=text)
            if elementtype == 'button':
                return Button(text=text, command=commmand)

        # 設定ファイルからステータスバーの設定を読み込むメソッド
        def statusbar_element_setting_load(name: str=str):
            if name == 'st1_letter_count': return(st1_letter_count())
            elif name == 'st2_letter_count': return(st2_letter_count())
            elif name == 'st3_letter_count': return(st3_letter_count())
            elif name == 'st1_line_count': return(st1_line_count())
            elif name == 'st2_line_count': return(st2_line_count())
            elif name == 'st3_line_count': return(st3_line_count())
            elif name == 'hotkeys1': return(self.hotkeys1)
            elif name == 'hotkeys2': return(self.hotkeys2)
            elif name == 'hotkeys3': return(self.hotkeys3)
            elif name == 'toolbutton_open': return(self.toolbutton_open)
            elif name == 'toolbutton_over_write_save': return(self.toolbutton_over_write_save)
            elif name == 'toolbutton_save_as': return(self.toolbutton_save_as)

        # ステータスバー設定メソッド
        def statusbar_element_setting(event=None, num=None):

            # 設定ファイルからステータスバーについての設定を読み込む
            try:
                l = self.settings['statusbar_element_settings'][num]
            except KeyError:
                l = self.settings['statusbar_element_settings'][num] = None
            finally:
                l2 = []

            try:
                for name in l: l2.append(statusbar_element_setting_load(name))
            except TypeError:
                pass

            # ステータスバーを作る
            print(l2)
            if num in range(6) and l2[0]:
                self.statusbar_element_dict[num] = dict()
                for i,e in zip(range(len(l2)), l2):
                    self.statusbar_element_dict[num][i] = make_toolbar_element(*e)
                for e in self.statusbar_element_dict[num].values():
                    eval('self.statusbar'+str(num)+'.add(e)')

        # ステータスバー更新メソッド
        def statusbar_element_reload(event=None):
            # self.statusbar_element_dict[0][0]['text'] = '更新'
            for e in self.statusbar_element_dict.keys():
                settings = self.settings['statusbar_element_settings'][e]
                for f in self.statusbar_element_dict[e].keys():
                    new_element = statusbar_element_setting_load(settings[f])[1]
                    if type(self.statusbar_element_dict[e][f]) == Label:
                        self.statusbar_element_dict[e][f]['text'] = new_element
                        # print(eval(f'self.statusbar{e}').paneconfig(self.statusbar_element_dict[e][f]))
                    elif type(self.statusbar_element_dict[e][f]) == Button: print('Button')

        self.statusbar_element_dict = dict()

        # ステータスバーを生成
        self.statusbar0 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar1 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar2 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar3 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar4 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar5 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        for i in range(6):
            statusbar_element_setting(num=i)


        # キーバインドを設定
        def test_focus_get(e):
            widget = self.master.focus_get()
            print(widget, 'has focus')
        def focus_to_right(e):
            e.widget.tk_focusNext().tk_focusNext().focus_set()
        def focus_to_left(e):
            e.widget.tk_focusPrev().tk_focusPrev().focus_set()
        self.master.bind('<KeyPress>', statusbar_element_reload)
        self.master.bind('<Control-Right>', focus_to_right)
        self.master.bind('<Control-f>', focus_to_right)
        self.master.bind('<Control-Left>', focus_to_left)
        self.master.bind('<Control-d>', focus_to_left)
        # self.master.bind('<Control-f>', test_focus_get)
        self.master.bind('<Control-^>', test_focus_get)

        # 各パーツを設置
        self.statusbar0.pack(fill=X, side=BOTTOM) if len(self.statusbar0.panes()) else print('statusbar0 was not packed')
        self.statusbar1.pack(fill=X, side=BOTTOM) if len(self.statusbar1.panes()) else print('statusbar1 was not packed')
        self.statusbar2.pack(fill=X, side=BOTTOM) if len(self.statusbar2.panes()) else print('statusbar2 was not packed')
        self.statusbar3.pack(fill=X, side=BOTTOM) if len(self.statusbar3.panes()) else print('statusbar3 was not packed')
        self.statusbar4.pack(fill=X) if len(self.statusbar4.panes()) else print('statusbar4 was not packed')
        self.statusbar5.pack(fill=X) if len(self.statusbar5.panes()) else print('statusbar5 was not packed')
        f1.pack(fill=BOTH, expand=YES)

        f1_1.place(relx=0.0, rely=0.0, relwidth=0.1, relheight=1.0)
        f1_2.place(relx=0.1, rely=0.0, relwidth=0.6, relheight=1.0)
        f1_3.place(relx=0.7, rely=0.0, relwidth=0.3, relheight=1.0)

        self.e1.pack(fill=X, expand=False, padx=3, pady=10)
        self.e2.pack(fill=X, expand=False, padx=3, pady=10)
        self.e3.pack(fill=X, expand=False, padx=3, pady=10)

        self.st1.pack(fill=BOTH, expand=YES)
        self.st2.pack(fill=BOTH, expand=YES)
        self.st3.pack(fill=BOTH, expand=YES)

        self.e1.focus_set()

    def file_open(self, event=None, file: str|None =None):
        '''
        '''
        self.file_close()
        # 開くファイルを指定されている場合、ファイル選択ダイアログをスキップする
        if file:
            self.filepath = file
        else:
            self.filepath = filedialog.askopenfilename(
                title = '編集ファイルを選択',
                initialdir = './',
                initialfile='file.tcs',
                filetypes=[('ThreeCrowsプロジェクトファイル', '.tcs'), ('YAMLファイル', '.yaml')],
                defaultextension='tcs')

        print(self.filepath)

        if self.filepath:
            try:
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
                    print(data)
                self.master.title(f'ThreeCrows - {self.filepath}')

                # 最近使用したファイルのリストを修正し、settings.yamlに反映
                self.recently_files.insert(0, self.filepath)
                # 重複を削除
                self.recently_files = list(dict.fromkeys(self.recently_files))
                # 5つ以内になるよう削除
                del self.recently_files[4:-1]
                self.settings['recently_files'] = self.recently_files
                with open('.\\settings.yaml', mode='wt', encoding='utf-8') as f:
                    yaml.dump(self.settings, f, allow_unicode=True)
                # 更新
                for i in self.recently_files:
                    print(i)
                    self.menu_file_recently.delete(0)
                try:
                    self.menu_file_recently.add_command(label=self.recently_files[0] + ' (現在のファイル)'#, command=lambda:self.file_open(file=self.recently_files[0])
                                                        )
                    self.menu_file_recently.unbind('<Control-r>', self.open_last_file_id)
                    self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file=self.recently_files[1]))
                    self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file=self.recently_files[2]))
                    self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file=self.recently_files[3]))
                    self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file=self.recently_files[4]))
                except IndexError:
                    pass

            except FileNotFoundError: # ファイルが存在しなかった場合
                md = MessageDialog(title='TreeCrows - エラー', alert=True, buttons=['OK'], message='ファイルが見つかりません')
                md.show()
                # 最近使用したファイルに見つからなかったファイルが入っている場合、削除しsettins.yamlに反映する
                if self.filepath in self.recently_files:
                    self.recently_files.remove(self.filepath)
                    self.settings['recently_files'] = self.recently_files
                with open('.\\settings.yaml', mode='wt', encoding='utf-8') as f:
                    yaml.dump(self.settings, f, allow_unicode=True)

            except KeyError: # ファイルが読み込めなかった場合
                md = MessageDialog(title='TreeCrows - エラー', alert=True, buttons=['OK'], message='ファイルが読み込めませんでした。\nファイルが破損している、またはThreeCrowsプロジェクトファイルではない可能性があります')
                md.show()

    def file_save_as(self, event=None):
        self.filepath = filedialog.asksaveasfilename(
            title='名前をつけて保存',
            initialdir='./',
            initialfile='noname',
            filetypes=[('ThreeCrowsプロジェクトファイル', '.tcs'), ('YAMLファイル', '.yaml')],
            defaultextension='tcs'
        )
        print(self.filepath)
        self.save_file(self.filepath)

    def file_over_write_save(self, event=None):
        if self.filepath:
            self.save_file(self.filepath)
        else:
            self.file_save_as()

    def save_file(self, path):
        try:
            with open(path, mode='wt', encoding='utf-8') as f:
                stre1, stre2, stre3 = self.e1.get(), self.e2.get(), self.e3.get()
                strst1, strst2, strst3 = self.st1.get('1.0', END).removesuffix('\n'), self.st2.get('1.0', END).removesuffix('\n'), self.st3.get('1.0', END).removesuffix('\n')
                d = {'line1_text': strst1, 'line1_title': stre1, 'line2_text': strst2, 'line2_title': stre2, 'line3_text': strst3, 'line3_title': stre3}
                self.data = d
                print(self.data)
                yaml.safe_dump(d, f, encoding='utf-8', allow_unicode=True)
        except FileNotFoundError as e:
            print(e)

    def file_close(self, shouldExit=False):

        stre1, stre2, stre3 = self.e1.get(), self.e2.get(), self.e3.get()
        strst1, strst2, strst3 = self.st1.get('1.0', END), self.st2.get('1.0', END), self.st3.get('1.0', END)
        d = {'line1_text': strst1, 'line1_title': stre1, 'line2_text': strst2, 'line2_title': stre2, 'line3_text': strst3, 'line3_title': stre3}
        for k in ('line1_text', 'line2_text', 'line3_text'):
            d[k] = d[k][:-1]
        print(self.data)
        print(d)

        if shouldExit:
            message = '更新内容が保存されていません。アプリを終了しますか。'
            title = 'ThreeCrows - 終了確認'
            buttons=['キャンセル', '保存せず終了:danger', '保存して終了:success']
        else:
            message = '更新内容が保存されていません。ファイルを閉じ、ファイルを変更しますか'
            title = 'ThreeCrows - 確認'
            buttons=['キャンセル', '保存せず変更:danger', '保存して変更:success']

        if self.data == d:
            self.master.destroy() if shouldExit else 0
        elif self.data != d:
            mg = MessageDialog(message=message, title=title, buttons=buttons)
            mg.show()
            if mg.result in ('保存せず終了', '保存せず変更'):
                if shouldExit:
                    self.master.destroy()
                else: print(buttons[1])
            elif mg.result in ('保存して終了', '保存して変更'):
                self.file_over_write_save()
                if shouldExit:
                    self.master.destroy()
                else: print(buttons[2])
            else:
                pass

        shouldExit = False

    def open_last_file(self, event=None):
        self.file_open(file=self.recently_files[0])

    def letter_count(self, obj:ScrolledText|str=''):
        if type(obj) == ttk.scrolled.ScrolledText: return len(obj.get("1.0", "end-1c").replace('\n', ''))
        elif type(obj) == str: return len(obj)

    def line_count(self, obj:ScrolledText|str=''):
        if type(obj) == ttk.scrolled.ScrolledText:
            obj = obj.get("1.0", "end-1c")
            return obj.count('\n') + 1 if obj else 0


class Editor(Main):

    def __init__(self, master=None, event=None):
        super().__init__(master, event)
        self.edit_histrory = []

    def recode_edit_history(self, event=None):
        print(app.e1.get())
        current_text = [app.e1.get(), app.st1.get('1.0',END),
                        app.e2.get(), app.st2.get('1.0',END),
                        app.e3.get(), app.st3.get('1.0',END)]
        print(self)
        print(self.edit_histrory)
        print(type(self.edit_histrory))
        self.edit_hisitory.append(current_text)
        print(self.edit_history)


if __name__ == "__main__":
    root = ttk.Window(title='ThreeCrows', themename='darkly')
    app = Main(master=root)
    app.mainloop()