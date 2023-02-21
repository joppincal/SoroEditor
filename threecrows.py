from collections import deque
import datetime
import re
from tkinter import filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap import Entry
from ttkbootstrap import Frame
from ttkbootstrap import Menu
from ttkbootstrap import Label
from ttkbootstrap import Button
from ttkbootstrap import PanedWindow
from ttkbootstrap import Style
from ttkbootstrap.dialogs.dialogs import MessageDialog
import yaml

class Main(Frame):

    def __init__(self, master=None, event=None, ):

        super().__init__(master)
        self.master.geometry('1200x800')

        self.master.protocol('WM_DELETE_WINDOW', lambda:self.file_close(shouldExit=True))

        version = '0.1.1'
        # 設定ファイルを読み込み
        self.settingFile_Error_md = None
        self.settings = {'columns': {'number': 3, 'percentage': [10, 60, 30]}, 'font': {'name': 'nomal', 'size': 0}, 'recently_files': [], 'statusbar_element_settings': {0: ['hotkeys3'], 1: ['hotkeys2'], 2: ['hotkeys1'], 3: ['st2_letter_count', 'st2_line_count', 'st3_letter_count', 'st3_line_count'], 4: [None], 5: [None]}, 'themename': 'darkly', 'version': version}
        try:
            with open(r'.\settings.yaml', mode='rt', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                if data.keys() != self.settings.keys(): raise KeyError
        ## 内容に不備がある場合新規作成し、古い設定ファイルを別名で保存
        except (KeyError, AttributeError) as e:
            print(e)
            with open('./settings.yaml', mode='wt', encoding='utf-8') as f:
                yaml.safe_dump(self.settings, f, allow_unicode=True)
            t = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
            error_filename = f'settings(error)_{t}.yaml'
            settingFile_Error_message = '設定ファイルに不備がありました\n'\
                                        '設定ファイルをデフォルトに戻します。\n'\
                                        '不備のある設定ファイルを'\
                                        f'「{error_filename}」として保存します'
            with open('./'+f'{error_filename}', mode='wt', encoding='utf-8') as f:
                yaml.safe_dump(data, f, allow_unicode=True)
            self.settingFile_Error_md = MessageDialog(message=settingFile_Error_message, buttons=['OK'])
        ## 存在しない場合新規作成
        except FileNotFoundError as e:
            print(e)
            with open(r'.\settings.yaml', mode='wt', encoding='utf-8') as f:
                yaml.safe_dump(self.settings, f, allow_unicode=True)
            settingFile_Error_message = '設定ファイルが存在しなかったため作成しました'
            self.settingFile_Error_md = MessageDialog(message=settingFile_Error_message, buttons=['OK'])
        else: self.settings = data

        # 設定ファイルから各設定を読み込み
        ## テーマを設定
        self.windowstyle = Style()
        self.windowstyle.theme_use(self.settings['themename'])
        ## フォント設定
        self.font = (self.settings['font']['name'], self.settings['font']['size'])
        ## 列数
        self.number_of_columns = (self.settings['columns']['number'])
        ## 列比率
        self.column_percentage = [x*0.01 for x in self.settings['columns']['percentage']]
        ## self.number_of_columns と self.column_percentage の内、数の少ない方に合わせる
        if self.number_of_columns < len(self.column_percentage):
            i = self.number_of_columns - len(self.column_percentage)
            del self.column_percentage[i:]
        elif self.number_of_columns > len(self.column_percentage):
            self.number_of_columns = len(self.column_percentage)

        self.filepath = ''
        self.data = {}
        for i in range(self.number_of_columns):
            self.data[i] = {'text': '', 'title': ''}
        self.edit_history = deque([self.data])
        self.undo_history = deque([])

        # メニューバーの作成
        menubar = Menu()

        ## メニュバー - ファイル
        self.menu_file = Menu(menubar, tearoff=False)

        self.menu_file.add_command(label='ファイルを開く', command=self.file_open, accelerator='Ctrl+O')
        self.menu_file.add_command(label='上書き保存', command=self.file_over_write_save, accelerator='Ctrl+S')
        self.menu_file.add_command(label='名前をつけて保存', command=self.file_save_as, accelerator='Ctrl+Shift+S')

        ###メニューバー - ファイル - 最近使用したファイル
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
        menubar.add_cascade(label='ファイル', menu=self.menu_file)

        ## メニューバー - 編集
        self.menu_edit = Menu(menubar, tearoff=False)

        self.menu_edit.add_command(label='切り取り', command=self.cut, accelerator='Ctrl+X')
        self.menu_edit.add_command(label='コピー', command=self.copy, accelerator='Ctrl+C')
        self.menu_edit.add_command(label='貼り付け', command=self.paste, accelerator='Ctrl+V')
        self.menu_edit.add_command(label='すべて選択', command=self.select_all, accelerator='Ctrl+A')
        self.menu_edit.add_command(label='1行選択', command=self.select_line)
        self.menu_edit.add_command(label='取り消し', command=self.undo, accelerator='Ctrl+Z')
        self.menu_edit.add_command(label='取り消しを戻す', command=self.repeat, accelerator='Ctrl+Shift+Z')
        menubar.add_cascade(label='編集', menu=self.menu_edit)

        self.master.config(menu = menubar)

        # 各パーツを製作
        f1 = Frame(self.master, padding=5)

        self.columns = []
        self.entrys = []
        self.scrolledtexts = []
        self.text_in_entrys = []
        self.text_in_scrolledtexts = []

        # テキストエディタ部分の要素の準備
        for i in range(self.number_of_columns):
            self.columns.append(Frame(f1, padding=5))
            self.entrys.append(Entry(self.columns[-1], font=self.font))
            self.scrolledtexts.append(ScrolledText(self.columns[-1], padding=3, autohide=True, wrap=CHAR, font=self.font))
            self.text_in_entrys.append('')
            self.text_in_scrolledtexts.append('')

        for i in range(self.number_of_columns):
            self.entrys[i].insert(END, self.text_in_entrys[i])
            self.scrolledtexts[i].insert(0., self.text_in_scrolledtexts[i])

        # statusbar要素を作成
        self.sb_elements = []
        def letter_count(i=0):
            if i >= self.number_of_columns: return
            title = self.entrys[i].get() if self.entrys[i].get() else f'列{i+1}'
            text = f'{title}: {self.letter_count(obj=self.scrolledtexts[i])}文字'
            return ('label', text)
        def line_count(i=0):
            if i >= self.number_of_columns: return
            title = self.entrys[i].get() if self.entrys[i].get() else f'列{i+1}'
            text = f'{title}: {self.line_count(obj=self.scrolledtexts[i])}行'
            return ('label', text)
        ## ショートカットキー
        self.hotkeys1 = ('label', '[Ctrl+O]: 開く  [Ctrl+S]: 上書き保存  [Ctrl+Shift+S]: 名前をつけて保存  [Ctrl+R]: 最後に使ったファイルを開く（起動直後のみ）')
        self.hotkeys2 = ('label', '[Ctrl+^]: 検索・置換  [Ctrl+Z]: 取り消し  [Ctrl+Shift+Z]: 取り消しを戻す')
        self.hotkeys3 = ('label', '[Ctrl+D][Ctrl+←]: 左に移る  [Ctrl+F][Ctrl+→]: 右に移る')
        ## 顔文字
        self.kaomoji = ''
        ## ツールボタン
        self.toolbutton_open = ('button', 'ファイルを開く[Ctrl+O]', self.file_open)
        self.toolbutton_over_write_save = ('button', '上書き保存[Ctrl+S]', self.file_over_write_save)
        self.toolbutton_save_as = ('button', '名前をつけて保存[Ctrl+Shift+S]', self.file_save_as)
        ## 初期ステータスバーメッセージ
        self.statusbar_message = ('label', 'ステータスバーの表示項目はメニューバーの 設定-ステータスバー から変更できます')

        # ステータスバー作成メソッド
        def make_toolbar_element(elementtype=None, text=None, commmand=None):
            if elementtype == 'label':
                return Label(text=text)
            if elementtype == 'button':
                return Button(text=text, command=commmand, bootstyle='secondary', takefocus=False)

        # 設定ファイルからステータスバーの設定を読み込むメソッド
        def statusbar_element_setting_load(name: str=str):
            if   re.compile(r'letter_count_\d+').search(name): return(letter_count(int(re.search(r'\d+', name).group()) - 1))
            elif re.compile(r'line_count_\d+').search(name): return(line_count(int(re.search(r'\d+', name).group()) - 1))
            elif name == 'hotkeys1': return(self.hotkeys1)
            elif name == 'hotkeys2': return(self.hotkeys2)
            elif name == 'hotkeys3': return(self.hotkeys3)
            elif name == 'toolbutton_open': return(self.toolbutton_open)
            elif name == 'toolbutton_over_write_save': return(self.toolbutton_over_write_save)
            elif name == 'toolbutton_save_as': return(self.toolbutton_save_as)
            elif name == 'statusbar_message': return(self.statusbar_message)

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
            except TypeError as e:
                pass

            # ステータスバーを作る
            if num in range(6) and l2:
                self.statusbar_element_dict[num] = dict()
                for i,e in zip(range(len(l2)), l2):
                    if not e: continue
                    self.statusbar_element_dict[num][i] = make_toolbar_element(*e)
                for e in self.statusbar_element_dict[num].values():
                    i = 1 if num < 4 else 0
                    eval('self.statusbar'+str(num)+f'.add(e, weight={str(i)})')
                if num > 3:
                    eval('self.statusbar'+str(num)+'.add(Frame())')

        # ステータスバー更新メソッド
        def statusbar_element_reload(event=None):
            for e in self.statusbar_element_dict.keys():
                settings = self.settings['statusbar_element_settings'][e]
                for f in self.statusbar_element_dict[e].keys():
                    new_element = statusbar_element_setting_load(settings[f])[1]
                    if type(self.statusbar_element_dict[e][f]) == Label:
                        self.statusbar_element_dict[e][f]['text'] = new_element
                    elif type(self.statusbar_element_dict[e][f]) == Button:
                        pass

        self.statusbar_element_dict = dict()

        # ステータスバーを生成
        self.statusbar0 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar1 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar2 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar3 = PanedWindow(self.master, height=30, orient=HORIZONTAL)
        self.statusbar4 = PanedWindow(self.master, height=50, orient=HORIZONTAL)
        self.statusbar5 = PanedWindow(self.master, height=50, orient=HORIZONTAL)
        for i in range(5):
            statusbar_element_setting(num=i)


        # キーバインドを設定
        def test_focus_get(e):
            widget = self.master.focus_get()
            print(widget, 'has focus')
        def focus_to_right(e):
            e.widget.tk_focusNext().tk_focusNext().focus_set()
        def focus_to_left(e):
            e.widget.tk_focusPrev().tk_focusPrev().focus_set()
        def focus_to_bottom(e):
            e.widget.tk_focusNext().focus_set()
        self.master.bind('<KeyPress>', statusbar_element_reload)
        # self.master.bind('<KeyPress>', lambda e:print(e), '+')
        # self.master.bind('<KeyRelease>', lambda e:print(e), '+')
        self.master.bind('<KeyRelease>', self.recode_edit_history)
        self.master.bind('<Control-z>', self.undo)
        self.master.bind('<Control-Z>', self.repeat)
        self.menu_file.bind_all('<Control-o>', self.file_open)
        self.menu_file.bind_all('<Control-s>', self.file_over_write_save)
        self.menu_file.bind_all('<Control-S>', self.file_save_as)
        self.master.bind('<Control-Right>', focus_to_right)
        self.master.bind('<Control-f>', focus_to_right)
        self.master.bind('<Control-Left>', focus_to_left)
        self.master.bind('<Control-d>', focus_to_left)
        for enter in self.entrys:
            enter.bind('<Down>', focus_to_bottom)
            enter.bind('<Return>', focus_to_bottom)
        self.master.bind('<Control-^>', test_focus_get)

        # 各パーツを設置
        self.statusbar0.pack(fill=X, side=BOTTOM, padx=5, pady=3) if len(self.statusbar0.panes()) else print('statusbar0 was not packed')
        self.statusbar1.pack(fill=X, side=BOTTOM, padx=5, pady=3) if len(self.statusbar1.panes()) else print('statusbar1 was not packed')
        self.statusbar2.pack(fill=X, side=BOTTOM, padx=5, pady=3) if len(self.statusbar2.panes()) else print('statusbar2 was not packed')
        self.statusbar3.pack(fill=X, side=BOTTOM, padx=5, pady=3) if len(self.statusbar3.panes()) else print('statusbar3 was not packed')
        self.statusbar4.pack(fill=X) if len(self.statusbar4.panes()) else print('statusbar4 was not packed')
        self.statusbar5.pack(fill=X) if len(self.statusbar5.panes()) else print('statusbar5 was not packed')

        f1.pack(fill=BOTH, expand=YES)

        for i in range(self.number_of_columns):
            # 列の枠を作成
            j, relx = i, 0.0
            while j: relx, j = relx + self.column_percentage[j-1], j - 1
            self.columns[i].place(relx=relx, rely=0.0, relwidth=self.column_percentage[i], relheight=1.0)
            # Entryを作成
            self.entrys[i].pack(fill=X, expand=False, padx=3, pady=10)
            # ScrolledTextを作成
            self.scrolledtexts[i].pack(fill=BOTH, expand=YES)

        # 初期フォーカスをe1にセット
        self.entrys[0].focus_set()


    # 以下、ファイルを開始、保存、終了に関するメソッド
    def file_open(self, event=None, file: str|None=None):
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
                    self.edit_history.clear()
                    for i in range(self.number_of_columns):
                        self.entrys[i].delete('0', END)
                        self.entrys[i].insert(END, data[i]['title'])
                        self.scrolledtexts[i].delete('1.0', END)
                        self.scrolledtexts[i].insert(0., data[i]['text'])
                    self.edit_history.appendleft(data)
                    print(data)
                self.master.title(f'ThreeCrows - {self.filepath}')

                # 最近使用したファイルのリストを修正し、settings.yamlに反映
                self.recently_files.insert(0, self.filepath)
                # 重複を削除
                self.recently_files = list(dict.fromkeys(self.recently_files))
                # 5つ以内になるよう削除
                del self.recently_files[4:-1]
                self.settings['recently_files'] = self.recently_files
                with open('./settings.yaml', mode='wt', encoding='utf-8') as f:
                    yaml.dump(self.settings, f, allow_unicode=True)
                # 更新
                for i in self.recently_files:
                    self.menu_file_recently.delete(0)
                try:
                    self.menu_file_recently.add_command(label=self.recently_files[0] + ' (現在のファイル)')
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
        if self.filepath:
            print(self.filepath)
            self.save_file(self.filepath)
            self.master.title(f'ThreeCrows - {self.filepath}')

    def file_over_write_save(self, event=None):
        if self.filepath:
            self.save_file(self.filepath)
        else:
            self.file_save_as()

    def save_file(self, path):
        try:
            with open(path, mode='wt', encoding='utf-8') as f:
                current_data = self.get_current_data()
                self.data = current_data
                yaml.safe_dump(current_data, f, encoding='utf-8', allow_unicode=True)
        except FileNotFoundError as e:
            print(e)

    def file_close(self, shouldExit=False):

        current_data = self.get_current_data()

        if shouldExit:
            message = '更新内容が保存されていません。アプリを終了しますか。'
            title = 'ThreeCrows - 終了確認'
            buttons=['キャンセル', '保存せず終了:danger', '保存して終了:success']
        else:
            message = '更新内容が保存されていません。ファイルを閉じ、ファイルを変更しますか'
            title = 'ThreeCrows - 確認'
            buttons=['キャンセル', '保存せず変更:danger', '保存して変更:success']

        if self.data == current_data:
            self.master.destroy() if shouldExit else 0
        elif self.data != current_data:
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
        if not self.filepath:
            self.file_open(file=self.recently_files[0])

    # 以下、エディターの編集に関するメソッド
    def letter_count(self, obj:ScrolledText|str=''):
        if type(obj) == ttk.scrolled.ScrolledText: return len(obj.get('1.0', 'end-1c').replace('\n', ''))
        elif type(obj) == str: return len(obj)

    def line_count(self, obj:ScrolledText|str=''):
        if type(obj) == ttk.scrolled.ScrolledText:
            obj = obj.get('1.0', 'end-1c')
            return obj.count('\n') + 1 if obj else 0

    def get_current_data(self):
        current_data = {}
        for i in range(self.number_of_columns):
            current_data[i] = {}
            current_data[i]['text'] = self.scrolledtexts[i].get('1.0',END).removesuffix('\n')
            current_data[i]['title'] = self.entrys[i].get()
        return current_data

    def recode_edit_history(self, event=None):
        current_text = self.get_current_data()
        if current_text != self.edit_history[0]:
            print('recode')
            self.edit_history.appendleft(current_text)
            self.undo_history.clear()

    def undo(self, event=None):
        if len(self.edit_history) <= 1:
            return
        current_text = self.get_current_data()
        for i in range(len(self.number_of_columns)):
            if current_text[i]['text'] != self.edit_history[1][i]['text']:
                print('undo')
                self.scrolledtexts[i].delete('1.0', END)
                self.scrolledtexts[i].insert(END, self.edit_history[1][i]['text'])
            if current_text[i]['title'] != self.edit_history[1][i]['title']:
                print('undo')
                self.entrys[i].delete('0', END)
                self.entrys[i].insert(END, self.edit_history[1][i]['title'])
        self.undo_history.appendleft(self.edit_history.popleft())

    def repeat(self, event=None):
        if len(self.undo_history) == 0:
            return
        current_text = self.get_current_data()
        for i in range(len(self.number_of_columns)):
            if current_text[i]['text'] != self.undo_history[0][i]['text']:
                print('repeat')
                self.scrolledtexts[i].delete('1.0', END)
                self.scrolledtexts[i].insert(END, self.undo_history[0][i]['text'])
            if current_text[i]['title'] != self.undo_history[0][i]['title']:
                print('repeat')
                self.entrys[i].delete('0', END)
                self.entrys[i].insert(END, self.undo_history[0][i]['title'])
        self.edit_history.appendleft(self.undo_history.popleft())

    def cut(self, e=None):
        if self.master.focus_get().winfo_class() == 'Text':
            if self.master.focus_get().tag_ranges(SEL):
                t = self.master.focus_get().get(SEL_FIRST, SEL_LAST)
                self.master.focus_get().delete(SEL_FIRST, SEL_LAST)
        elif self.master.focus_get().winfo_class() == 'TEntry':
            t = self.master.focus_get().get()
            self.master.focus_get().delete(0, END)
        self.clipboard_append(t)
        self.recode_edit_history()
    def copy(self, e=None):
        if self.master.focus_get().winfo_class() == 'Text':
            if self.master.focus_get().tag_ranges(SEL):
                t = self.master.focus_get().get(SEL_FIRST, SEL_LAST)
        elif self.master.focus_get().winfo_class() == 'TEntry':
            t = self.master.focus_get().get()
        self.clipboard_append(t)
    def paste(self, e=None):
        t = self.clipboard_get()
        if self.master.focus_get().winfo_class() == 'Text':
            if self.master.focus_get().tag_ranges(SEL):
                self.master.focus_get().delete(SEL_FIRST, SEL_LAST)
        self.master.focus_get().insert(INSERT, t)
        self.recode_edit_history()
    def select_all(self, e=None):
        if self.master.focus_get().winfo_class() == 'Text':
            self.master.focus_get().tag_add(SEL, '1.0', END+'-1c')
        elif self.master.focus_get().winfo_class() == 'TEntry':
            pass
    def select_line(self, e=None):
        if self.master.focus_get().winfo_class() == 'Text':
            self.master.focus_get().tag_add(SEL, INSERT+' linestart', INSERT+' lineend')
        elif self.master.focus_get().winfo_class() == 'TEntry':
            pass


if __name__ == '__main__':
    root = ttk.Window(title='ThreeCrows', minsize=(1200, 800))
    app = Main(master=root)
    if app.settingFile_Error_md: app.settingFile_Error_md.show()
    app.mainloop()