from collections import deque
import datetime
from os import chdir, path
import re
import sys
import tkinter
from tkinter import BooleanVar, StringVar, filedialog, font
import webbrowser
from ttkbootstrap import Button, Checkbutton, Entry, Frame, Label, Labelframe,\
    Menu, Notebook, OptionMenu, PanedWindow, Scrollbar, Separator, Spinbox, Style,\
    Text, Toplevel, Treeview
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs.dialogs import MessageDialog
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.themes.standard import *
import yaml

__version__ = '0.2.3'
__projversion__ = '0.2.0'
with open(path.join(path.dirname(__file__), 'ThirdPartyNotices.txt'), 'rt', encoding='utf-8') as f:
    __thirdpartynotices__ = f.read()
__icon__ = ''

class Main(Frame):

    def __init__(self, master=None):

        super().__init__(master)
        self.master.geometry('1600x1000')

        self.master.protocol('WM_DELETE_WINDOW', lambda:self.file_close(True))

        # 設定ファイルを読み込み
        chdir(path.dirname(sys.argv[0]))
        self.settingFile_Error_md = None
        self.settings ={'between_lines': 10,
                        'columns': {'number': 3, 'percentage': [10, 60, 30]},
                        'display_line_number': False,
                        'font': {'family': 'nomal', 'size': 20},
                        'recently_files': [],
                        'selection_line_highlight': True,
                        'statusbar_element_settings':
                            {0: ['statusbar_message'],
                            1: ['hotkeys2', 'hotkeys3'],
                            2: ['hotkeys1'],
                            3: ['letter_count_2', 'line_count_2', 'letter_count_3', 'line_count_3'],
                            4: ['toolbutton_open', 'toolbutton_over_write_save', 'toolbutton_save_as'],
                            5: [None]},
                        'themename': 'darkly',
                        'version': __version__,
                        'wrap': NONE}
        try:
            with open('./settings.yaml', mode='rt', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                if set(data.keys()) & set(self.settings.keys()) != set(data.keys()):
                    raise KeyError
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
            with open('./settings.yaml', mode='wt', encoding='utf-8') as f:
                yaml.safe_dump(self.settings, f, allow_unicode=True)
            settingFile_Error_message = '設定ファイルが存在しなかったため作成しました'
            self.settingFile_Error_md = MessageDialog(message=settingFile_Error_message, buttons=['OK'])
        else: self.settings = data

        # 設定ファイルから各設定を読み込み
        ## テーマを設定
        self.windowstyle = Style()
        self.windowstyle.theme_use(self.settings['themename'])
        ## フォント設定
        self.font = font.Font(family=self.settings['font']['family'], size=self.settings['font']['size'])
        ## 列数
        self.number_of_columns = self.settings['columns']['number']
        ## 列比率
        self.column_percentage = [x*0.01 for x in self.settings['columns']['percentage']]
        for _ in range(10-len(self.column_percentage)):
            self.column_percentage.append(0)
        ## 折り返し
        self.wrap = self.settings['wrap']
        ## 行間
        self.between_lines = self.settings['between_lines']
        ## 行番号
        self.display_line_number = self.settings['display_line_number']
        ## 選択行の強調
        self.selection_line_highlight = self.settings['selection_line_highlight']

        self.filepath = ''
        self.data0 = {}
        for i in range(10):
            self.data0[i] = {'text': '', 'title': ''}
        self.data0['columns'] = {'number': self.number_of_columns, 'percentage': [int(x*100) for x in self.column_percentage]}
        self.data0['version'] = __projversion__
        self.data = self.data0
        self.edit_history = deque([self.data])
        self.undo_history = deque([])
        self.recently_files: list = self.settings['recently_files']
        try:
            self.initialdir = path.dirname(self.recently_files[0])
        except (IndexError, TypeError):
            self.initialdir = ''

        # メニューバーの作成
        menubar = Menu()

        ## メニュバー - ファイル
        self.menu_file = Menu(menubar, tearoff=False)

        self.menu_file.add_command(label='ファイルを開く(O)', command=self.file_open, accelerator='Ctrl+O', underline=8)
        self.menu_file.add_command(label='上書き保存(S)', command=self.file_over_write_save, accelerator='Ctrl+S', underline=6)
        self.menu_file.add_command(label='名前をつけて保存(A)', command=self.file_save_as, accelerator='Ctrl+Shift+S', underline=9)
        self.menu_file.add_command(label='プロジェクト設定(F)', command=ProjectFileSettingWindow, underline=9)

        ###メニューバー - ファイル - 最近使用したファイル
        self.menu_file_recently = Menu(self.menu_file, tearoff=False)

        try:
            self.menu_file_recently.add_command(label=self.recently_files[0], command=lambda:self.file_open(file=self.recently_files[0]), accelerator='Ctrl+R')
            self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file=self.recently_files[1]))
            self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file=self.recently_files[2]))
            self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file=self.recently_files[3]))
            self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file=self.recently_files[4]))
        except IndexError as e:
            pass

        self.menu_file.add_cascade(label='最近使用したファイル(R)', menu=self.menu_file_recently, underline=11)

        self.menu_file.add_separator()
        self.menu_file.add_command(label='終了(Q)', command=lambda:self.file_close(shouldExit=True), underline=3)
        menubar.add_cascade(label='ファイル(F)', menu=self.menu_file, underline=5)

        ## メニューバー - 編集
        self.menu_edit = Menu(menubar, tearoff=False)

        self.menu_edit.add_command(label='切り取り(T)', command=self.cut, accelerator='Ctrl+X', underline=5)
        self.menu_edit.add_command(label='コピー(C)', command=self.copy, accelerator='Ctrl+C', underline=4)
        self.menu_edit.add_command(label='貼り付け(P)', command=self.paste, accelerator='Ctrl+V', underline=5)
        self.menu_edit.add_command(label='すべて選択(A)', command=self.select_all, accelerator='Ctrl+A', underline=6)
        self.menu_edit.add_command(label='1行選択(L)', command=self.select_line, underline=5)
        self.menu_edit.add_command(label='取り消し(U)', command=self.undo, accelerator='Ctrl+Z', underline=5)
        self.menu_edit.add_command(label='取り消しを戻す(R)', command=self.repeat, accelerator='Ctrl+Shift+Z', underline=8)
        menubar.add_cascade(label='編集(E)', menu=self.menu_edit, underline=3)

        ## メニューバー - 設定
        self.menu_setting = Menu(menubar, tearoff=True)
        self.menu_setting.add_command(label='設定(S)', command=SettingWindow, accelerator='Ctrl+Shift+P', underline=3)
        menubar.add_cascade(label='設定(S)', menu=self.menu_setting, underline=3)

        ## メニューバー - ヘルプ
        self.menu_help = Menu(menubar, tearoff=True)
        self.menu_help.add_command(label='ヘルプを表示(H)', command=HelpWindow, accelerator='F1', underline=7)
        self.menu_help.add_command(label='CastellaEditorについて(A)', command=AboutWindow, underline=16)
        self.menu_help.add_command(label='ライセンス情報(L)', command=ThirdPartyNoticesWindow, underline=8)
        menubar.add_cascade(label='ヘルプ(H)', menu=self.menu_help, underline=4)

        # メニューバーの設置
        self.master.config(menu = menubar)

        # メニューバー関連のキーバインドを設定
        self.open_last_file_id = self.menu_file_recently.bind_all('<Control-r>', self.open_last_file)

        # 各パーツを製作
        self.f1 = Frame(self.master, padding=5)
        self.f2 = Frame(self.master)
        self.vbar = Scrollbar(self.f2, command=self.vbar_command, style=ROUND, takefocus=False)

        self.columns = []
        self.entrys = []
        self.maintexts = []

        self.make_text_editor()

        # 初期フォーカスを列1のタイトルにセット
        self.entrys[0].focus_set()

        # ステータスバー要素を作成
        self.sb_elements = []
        ## 文字数カウンター
        def letter_count(i=0):
            if i >= self.number_of_columns: return
            title = self.entrys[i].get() if self.entrys[i].get() else f'列{i+1}'
            text = f'{title}: {self.letter_count(obj=self.maintexts[i])}文字'
            return ('label', text)
        ## 行数カウンター
        def line_count(i=0):
            if i >= self.number_of_columns: return
            title = self.entrys[i].get() if self.entrys[i].get() else f'列{i+1}'
            text = f'{title}: {self.line_count(obj=self.maintexts[i])}行'
            return ('label', text)
        ## 行数カウンター（デバッグ）
        def line_count_debug(i=0):
            if i >= self.number_of_columns: return
            title = self.entrys[i].get() if self.entrys[i].get() else f'列{i+1}'
            text = f'{title}: {self.line_count(obj=self.maintexts[i], debug=True)}行'
            return ('label', text)
        ## カーソルの現在位置
        def current_place():
            widget = self.master.focus_get()
            try:
                index = widget.index(INSERT)
            except AttributeError:
                return ('label', ' '*40)
            text = ''
            for i in range(len(self.entrys)):
                if widget.winfo_id() == self.entrys[i].winfo_id():
                    title = self.entrys[i].get()
                    title = title if title else f'列{i+1}'
                    text = f'{title} - タイトル: {int(index)+1}字'
            for i in range(len(self.maintexts)):
                if widget.winfo_id() == self.maintexts[i].winfo_id():
                    title = self.entrys[i].get()
                    title = title if title else f'列{i+1}'
                    l = index.split('.')
                    text = f'{title} - 本文: {l[0]}行 {int(l[1])+1}字'
            return ('label', text)
        ## ショートカットキー
        self.hotkeys1 = ('label', '[Ctrl+O]: 開く  [Ctrl+S]: 上書き保存  [Ctrl+Shift+S]: 名前をつけて保存  [Ctrl+R]: 最後に使ったファイルを開く（起動直後のみ）')
        self.hotkeys2 = ('label', '[Ctrl+^]: 検索・置換  [Ctrl+Z]: 取り消し  [Ctrl+Shift+Z]: 取り消しを戻す')
        self.hotkeys3 = ('label', '[Ctrl+Q][Ctrl+,]: 左に移る  [Ctrl+W][Ctrl+.]: 右に移る')
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
                return Button(text=text, command=commmand, takefocus=False)

        # 設定ファイルからステータスバーの設定を読み込むメソッド
        def statusbar_element_setting_load(name: str=str):
            if   re.compile(r'letter_count_\d+').search(name): return(letter_count(int(re.search(r'\d+', name).group()) - 1))
            elif re.compile(r'line_count_\d+').search(name): return(line_count(int(re.search(r'\d+', name).group()) - 1))
            elif re.compile(r'line_count_debug_\d+').search(name): return(line_count_debug(int(re.search(r'\d+', name).group()) - 1))
            elif name == 'current_place': return current_place()
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
                    self.statusbars[num].add(e, weight=i)
                if num > 3:
                    self.statusbars[num].add(Frame())

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
        self.statusbars = list()
        for i in range(4):
            self.statusbars.append(PanedWindow(self.master, height=30, orient=HORIZONTAL))
        for i in range(2):
            self.statusbars.append(PanedWindow(self.master, height=50, orient=HORIZONTAL))
        for i in range(6):
            statusbar_element_setting(num=i)


        # キーバインドを設定
        def test_focus_get(e):
            widget = self.master.focus_get()
            print(widget, 'has focus')
        def focus_to_right(e):
            if type(e.widget) == tkinter.Text:
                insert = e.widget.index(INSERT)
                e.widget.tk_focusNext().tk_focusNext().mark_set(INSERT, insert)
            e.widget.tk_focusNext().tk_focusNext().focus_set()
        def focus_to_left(e):
            if type(e.widget) == tkinter.Text:
                insert = e.widget.index(INSERT)
                e.widget.tk_focusPrev().tk_focusPrev().mark_set(INSERT, insert)
            e.widget.tk_focusPrev().tk_focusPrev().focus_set()
        def focus_to_bottom(e):
            e.widget.tk_focusNext().focus_set()
        def reload(e):
            statusbar_element_reload()
            if self.selection_line_highlight: self.highlight()
        self.master.bind('<KeyPress>', reload)
        self.master.bind('<KeyRelease>', reload)
        self.master.bind('<Button>', reload)
        self.master.bind('<ButtonRelease>', reload)
        # self.master.bind('<KeyPress>', lambda e:print(e), '+')
        # self.master.bind('<KeyRelease>', lambda e:print(e), '+')
        self.master.bind('<KeyRelease>', self.recode_edit_history, '+')
        self.master.bind('<Control-z>', self.undo)
        self.master.bind('<Control-Z>', self.repeat)
        self.menu_file.bind_all('<Control-o>', self.file_open)
        self.menu_file.bind_all('<Control-s>', self.file_over_write_save)
        self.menu_file.bind_all('<Control-S>', self.file_save_as)
        self.master.bind('<Control-P>', SettingWindow)
        self.master.bind('<Control-w>', focus_to_right)
        self.master.bind('<Control-.>', focus_to_right)
        self.master.bind('<Control-q>', focus_to_left)
        self.master.bind('<Control-,>', focus_to_left)
        self.master.bind('<Control-l>', self.select_line)
        self.master.bind('<Control-Return>', self.newline)
        for entry in self.entrys:
            entry.bind('<Down>', focus_to_bottom)
            entry.bind('<Return>', focus_to_bottom)
        self.master.bind('<Control-^>', test_focus_get)

        # 各パーツを設置
        for w in self.statusbars[0:4]:
            w.pack(fill=X, side=BOTTOM, padx=5, pady=3) if len(w.panes()) else print(f'{w} was not packed')
        for w in self.statusbars[4:6]:
            w.pack(fill=X) if len(w.panes()) else print(f'{w} was not packed')

        self.f2.pack(fill=Y, side=RIGHT, pady=18)
        # self.vbar.pack(fill=Y, expand=YES)
        self.f1.pack(fill=BOTH, expand=YES)

        # 設定ファイルに異常があり初期化した場合、その旨を通知する
        if self.settingFile_Error_md: self.settingFile_Error_md.show()

        # ファイルを渡されているとき、そのファイルを開く
        if len(sys.argv) > 1:
            self.file_open(file=sys.argv[1])


    # 以下、ファイルを開始、保存、終了に関するメソッド
    def file_open(self, event=None, file: str|None=None):
        '''
        '''
        self.file_close()
        # 開くファイルを指定されている場合、ファイル選択ダイアログをスキップする
        if file:
            self.filepath = file
        else:
            newfilepath = filedialog.askopenfilename(
                title = '編集ファイルを選択',
                initialdir = self.initialdir,
                initialfile='file.cep',
                filetypes=[('CastellaEditorプロジェクトファイル', '.cep'), ('YAMLファイル', '.yaml'), ('ThreeCrowsプロジェクトファイル', '.tcs')],
                defaultextension='cep')
            if newfilepath: self.filepath = newfilepath

        print(self.filepath)

        if self.filepath:
            try:
                with open(self.filepath, mode='rt', encoding='utf-8') as f:
                    data = self.data0.copy()
                    self.data = yaml.safe_load(f)
                    data.update(self.data)
                    if not 'columns' in data.keys():
                        data['columns'] = {}
                    if not 'number' in data['columns'].keys():
                        data['columns']['number'] = self.settings['columns']['number']
                    if not 'percentage' in data['columns'].keys():
                        data['columns']['percentage'] = self.settings['columns']['percentage']

                    self.number_of_columns = data['columns']['number']
                    self.column_percentage = [x*0.01 for x in data['columns']['percentage']]
                    for _ in range(10-len(self.column_percentage)):
                        self.column_percentage.append(0)

                    self.make_text_editor()

                    self.edit_history.clear()

                    for i in range(self.number_of_columns, 10):
                        mg_exist_hidden_data = MessageDialog(f'列{i+1}にデータがありますが、表示されません\n'
                                                            f'title: {data[i]["title"]}\n'
                                                            f'text: {data[i]["text"][0:10]}...\n'
                                                            '確認するにはプロジェクトファイルを普通のテキストエディタで開いて直接データを見るか、'
                                                            'ファイル - プロジェクト設定 からこのファイルで表示する列を変更し、'
                                                            f'列{self.number_of_columns+i-1}が見えるようにしてください', 'CastellaEditor - 隠されたデータ')
                        if data[i]['title']: mg_exist_hidden_data.show()
                        elif data[i]['text']: mg_exist_hidden_data.show()
                    self.entrys[0].focus_set()
                    self.edit_history.appendleft(data)

                self.initialdir = path.dirname(self.filepath)
                self.master.title(f'CastellaEditor - {self.filepath}')

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
                with open('./settings.yaml', mode='wt', encoding='utf-8') as f:
                    yaml.dump(self.settings, f, allow_unicode=True)

            except (KeyError, UnicodeDecodeError, yaml.scanner.ScannerError): # ファイルが読み込めなかった場合
                md = MessageDialog(title='TreeCrows - エラー', alert=True, buttons=['OK'], message='ファイルが読み込めませんでした。\nファイルが破損している、またはCastellaEditorプロジェクトファイルではない可能性があります')
                md.show()

    def file_save_as(self, event=None):
        self.filepath = filedialog.asksaveasfilename(
            title='名前をつけて保存',
            initialdir=self.initialdir,
            initialfile='noname',
            filetypes=[('CastellaEditorプロジェクトファイル', '.cep'), ('YAMLファイル', '.yaml')],
            defaultextension='cep'
        )
        if self.filepath:
            print(self.filepath)
            self.save_file(self.filepath)
            self.master.title(f'CastellaEditor - {self.filepath}')
            return True
        else: return False

    def file_over_write_save(self, event=None):
        if self.filepath:
            self.save_file(self.filepath)
            return True
        else:
            if self.file_save_as(): return True
            else: return False

    def save_file(self, path):
        try:
            with open(path, mode='wt', encoding='utf-8') as f:
                current_data = self.get_current_data()
                self.data = current_data
                yaml.safe_dump(current_data, f, encoding='utf-8', allow_unicode=True)
        except FileNotFoundError as e:
            print(e)

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
        for _ in self.recently_files:
            self.menu_file_recently.delete(0)
        try:
            self.menu_file_recently.add_command(label=self.recently_files[0] + ' (現在のファイル)')
            self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file=self.recently_files[1]))
            self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file=self.recently_files[2]))
            self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file=self.recently_files[3]))
            self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file=self.recently_files[4]))
        except IndexError:
            pass

    def file_close(self, shouldExit=False):

        current_data = self.get_current_data()

        if shouldExit:
            message = '更新内容が保存されていません。アプリを終了しますか。'
            title = 'CastellaEditor - 終了確認'
            buttons=['保存して終了:success', '保存せず終了:danger', 'キャンセル']
        else:
            message = '更新内容が保存されていません。ファイルを閉じ、ファイルを変更しますか'
            title = 'CastellaEditor - 確認'
            buttons=['保存して変更:success', '保存せず変更:danger', 'キャンセル']

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
                if self.file_over_write_save():
                    if shouldExit:
                        self.master.destroy()
                    else: print(buttons[2])
            else:
                pass

    def open_last_file(self, event=None):
        if not self.filepath:
            self.file_open(file=self.recently_files[0])

    # 以下、エディターの編集に関するメソッド
    def letter_count(self, obj:ScrolledText|Text|str=''):
        if type(obj) == ttk.scrolled.ScrolledText or ttk.Text: return len(obj.get('1.0', 'end-1c').replace('\n', ''))
        elif type(obj) == str: return len(obj)

    def line_count(self, obj:ScrolledText|Text|str='', debug=False):
        if type(obj) == ScrolledText or Text:
            obj = obj.get('1.0', 'end-1c')
        elif type(obj) == str:
            pass
        if not debug:
            obj = obj.rstrip('\r\n')
        return obj.count('\n') + 1 if obj else 0

    def get_current_data(self):
        current_data = {}
        # 辞書の元を作る
        for i in range(10):
            current_data[i] = {'text': '', 'title': ''}
        # ファイルの初期データで上書きする（表示されていない列のデータ消滅対策）
        current_data.update(self.data)
        # 表示されている列のデータで上書きする
        for i in range(self.number_of_columns):
            current_data[i] = {}
            current_data[i]['text'] = self.maintexts[i].get('1.0',END).rstrip('\r\n')
            current_data[i]['title'] = self.entrys[i].get()
        # 設定部分を書き込む
        current_data['columns'] = {'number': self.number_of_columns, 'percentage': [int(x*100) for x in self.column_percentage]}
        for _ in range(10-len(current_data['columns']['percentage'])):
            current_data['columns']['percentage'].append(0)
        current_data['version'] = __projversion__
        return current_data

    def recode_edit_history(self, event=None):
        current_data = self.get_current_data()
        if current_data != self.edit_history[0]:
            print('recode')
            self.edit_history.appendleft(current_data)
            self.undo_history.clear()

    def undo(self, event=None):
        if len(self.edit_history) <= 1:
            return
        current_data = self.get_current_data()
        if type(self.master.focus_get()) == tkinter.Text:
            insert = self.master.focus_get().index(INSERT)
        for i in range(self.number_of_columns):
            if current_data[i]['text'] != self.edit_history[1][i]['text']:
                print('undo')
                self.maintexts[i].delete('1.0', END)
                self.maintexts[i].insert(END, self.edit_history[1][i]['text'])
            if current_data[i]['title'] != self.edit_history[1][i]['title']:
                print('undo')
                self.entrys[i].delete('0', END)
                self.entrys[i].insert(END, self.edit_history[1][i]['title'])
        self.undo_history.appendleft(self.edit_history.popleft())
        self.master.focus_get().mark_set(INSERT, insert)
        self.master.focus_get().see(INSERT)

    def repeat(self, event=None):
        if len(self.undo_history) == 0:
            return
        current_data = self.get_current_data()
        if type(self.master.focus_get()) == tkinter.Text:
            insert = self.master.focus_get().index(INSERT)
        for i in range(self.number_of_columns):
            if current_data[i]['text'] != self.undo_history[0][i]['text']:
                print('repeat')
                self.maintexts[i].delete('1.0', END)
                self.maintexts[i].insert(END, self.undo_history[0][i]['text'])
            if current_data[i]['title'] != self.undo_history[0][i]['title']:
                print('repeat')
                self.entrys[i].delete('0', END)
                self.entrys[i].insert(END, self.undo_history[0][i]['title'])
        self.edit_history.appendleft(self.undo_history.popleft())
        self.master.focus_get().mark_set(INSERT, insert)
        self.master.focus_get().see(INSERT)

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

    def newline(self, e):
        if type(e.widget) == tkinter.Text:
            insert = e.widget.index(INSERT+ '-1lines linestart')
            e.widget.delete(insert+'-1c', insert)
            for w in self.maintexts:
                w.insert(insert, '\n')
                w.see(e.widget.index(INSERT))
                Text
            print(e.widget.index(INSERT))

    # テキストエディタ部分の要素の準備
    def make_text_editor(self):

        try:
            self.f1_2.destroy()
            self.line_number_box.destroy()
            self.line_number_box_entry.destroy()
            self.f1_1.destroy()
        except AttributeError:
            pass
        for w in self.columns:
            w.destroy()
        for w in self.entrys:
            w.destroy()
        for w in self.maintexts:
            w.destroy()

        ## 行番号
        if self.display_line_number:

            self.f1_1 = Frame(self.f1, padding=0)
            self.line_number_box_entry = Entry(self.f1_1, width=3,
                                                state=DISABLED, takefocus=NO)
            self.line_number_box = Text(self.f1_1, font=self.font,
                                        width=4,
                                        spacing3=self.between_lines,
                                        takefocus=NO,
                                        yscrollcommand=self.yscrollcommand)
            self.line_number_box.tag_config('right', justify=RIGHT)
            for i in range(9999):
                i = i + 1
                self.line_number_box.insert(END, f'{i}\n', 'right')
            self.line_number_box.config(state=DISABLED)

            self.f1_1.pack(fill=Y, side=LEFT, pady=10)
            self.line_number_box_entry.pack(fill=X, expand=False, padx=5, pady=0)
            self.line_number_box.pack(fill=BOTH, expand=YES, padx=5, pady=10)

        ## テキストエディタ本体
        self.f1_2 = Frame(self.f1, padding=0)
        self.f1_2.pack(fill=BOTH, expand=YES, pady=10)

        self.columns = []
        self.entrys = []
        self.maintexts = []

        for i in range(self.number_of_columns):
            self.columns.append(Frame(self.f1_2, padding=0))
            self.entrys.append(Entry(self.columns[i], font=self.font))
            self.maintexts.append(Text(self.columns[i], wrap=self.wrap,
                                    font=self.font, yscrollcommand=self.yscrollcommand,
                                    spacing3=self.between_lines))

        for i in range(self.number_of_columns):
            # 列の枠を作成
            j, relx = i, 0.0
            while j: relx, j = relx + self.column_percentage[j-1], j - 1
            self.columns[i].place(relx=relx, rely=0.0, relwidth=self.column_percentage[i], relheight=1.0)
            # Entryを作成
            self.entrys[i].pack(fill=X, expand=False, padx=5, pady=0)
            # ScrolledTextを作成
            self.maintexts[i].pack(fill=BOTH, expand=YES, padx=5, pady=10)

        for i in range(self.number_of_columns):
            self.entrys[i].insert(END, self.data[i]['title'])
            self.maintexts[i].insert(END, self.data[i]['text'])

        self.align_the_number_of_rows()

# 各列の行数をそろえるメソッド
    def align_the_number_of_rows(self, e=None):
        data = self.get_current_data()
        for i in range(self.number_of_columns):
            newdata = data[i]['text']+'\n'*(9999-self.line_count(self.maintexts[i]))
            self.maintexts[i].delete(1.0, END)
            self.maintexts[i].insert(END, newdata)
        if e:
            self.maintexts[0].see(INSERT)

# カーソルのある行を強調するメソッド
    def highlight(self):
        if type(self.master.focus_get()) == tkinter.Text:
            insert = self.master.focus_get().index(INSERT)
            hilight_font = self.font.copy()
            hilight_font.config(weight='normal', slant='roman', underline=True)
            for w in self.maintexts:
                w.mark_set(INSERT, insert)
                w.tag_delete('insert_line')
                w.tag_add('insert_line', insert+' linestart', insert+' lineend')
                w.tag_config('insert_line', underline=False, font=hilight_font)

    # 縦スクロールバーが動いたときのメソッド
    def vbar_command(self, e=None, a=None, b=None):
        boxes = self.maintexts.copy()
        try:
            boxes.append(self.line_number_box)
        except AttributeError:
            pass
        for w in boxes:
            w.yview(e, a, b)
        self.yscrollcommand()

    # テキストウィジェットを直接動かしたときのメソッド
    def yscrollcommand(self, a=None, b=None):
        # マウスカーソルがあるウィジェットを特定する
        widget = None
        pointer = self.master.winfo_pointerxy()
        boxes = self.maintexts.copy()
        try:
            boxes.append(self.line_number_box)
        except AttributeError:
            pass
        for w in boxes:
            x_range = (w.winfo_rootx(), w.winfo_rootx()+w.winfo_width())
            y_range = (w.winfo_rooty(), w.winfo_rooty()+w.winfo_height())
            if x_range[0]<=pointer[0]<=x_range[1] and y_range[0]<=pointer[1]<=y_range[1]:
                widget = w
        self.align_the_lines(widget=widget)

    # 位置を合わせる
    def align_the_lines(self, e=None, widget=None):
        if widget:
            height = widget.winfo_height()
            top = widget.index('@0,0')
            bottom = widget.index(f'@0,{height}')
            boxes = self.maintexts.copy()
            try:
                boxes.append(self.line_number_box)
            except AttributeError:
                pass
            for w in boxes:
                w.see(bottom)
                w.see(top)
            self.show_cursor_at_bottom_line()
        # スクロールバーの調整（調整用の改行を無視するようにする）
        l = [self.line_count(w) for w in self.maintexts]
        l2 = [self.line_count(w, True) for w in self.maintexts]
        x = l.index(max(l))
        try:
            t = [a/l[x]*l2[x] for a in self.maintexts[x].yview()]
        except ZeroDivisionError:
            t = [0.0, 1.0]
        for i in range(2):
            if t[i] > 1.0: t[i] = 1.0
        self.vbar.set(*t)

    # カーソルが最下端の行にあり、最下端の行が隠れているとき、表示する
    def show_cursor_at_bottom_line(self, e=None):
        widget = self.master.focus_get()
        if not widget:
            widget = self.maintexts[0]
        if widget.winfo_class() == 'Text':
            height = widget.winfo_height()
            bottom = widget.index(f'@0,{height}')
            for w in self.maintexts:
                if bottom == widget.index(INSERT+' linestart'):
                    w.see(float(bottom)+1.0)


    # 以下、設定ウィンドウに関する設定
class SettingWindow(Toplevel):
    def __init__(self, title="CastellaEditor - 設定", iconphoto='', size=(1200, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        # 設定ファイルの読み込み
        with open(r'.\settings.yaml', mode='rt', encoding='utf-8') as f:
            self.settings = yaml.safe_load(f)
        number_of_columns = self.settings['columns']['number']
        percentage_of_columns = self.settings['columns']['percentage']
        self.themename = self.settings['themename']
        self.font_family = self.settings['font']['family']
        if self.font_family == 'nomal':
            self.font_family = font.nametofont("TkDefaultFont").actual()['family']
        self.font_size = self.settings['font']['size']
        self.statusbar_elements_dict = self.settings['statusbar_element_settings']
        self.between_lines = self.settings['between_lines']
        self.wrap = self.settings['wrap']
        display_line_number = self.settings['display_line_number']
        selection_line_highlight = self.settings['selection_line_highlight']

        self.statusbar_elements_dict_converted = {}
        for i in range(6):
            l = self.statusbar_elements_dict[i]
            self.statusbar_elements_dict_converted[i] = [self.convert_statusbar_elements(val) for val in l]
            while len(self.statusbar_elements_dict_converted[i]) < 4:
                self.statusbar_elements_dict_converted[i].append('')

        self.tv = None

        under = Frame(self)
        under.pack(side=BOTTOM, fill=X, pady=10)
        Button(under, text='キャンセル', command=self.destroy).pack(side=RIGHT, padx=10)
        Button(under, text='OK', command=lambda:self.save(close=True)).pack(side=RIGHT, padx=10)
        Button(under, text='適用', command=lambda:self.save()).pack(side=RIGHT, padx=10)

        f = Labelframe(self, text='設定')
        f.pack(fill=BOTH, expand=True, padx=5, pady=5)
        nt = Notebook(f, padding=5)
        nt.pack(expand=True, fill=BOTH)

        # 以下、設定項目
        f1 = Frame(nt, padding=5)
        nt.add(f1, padding=5, text='    列    ')

        ## 列について
        Label(f1, text='この設定は、新規プロジェクトのデフォルト設定です。'+
            '現在編集中のプロジェクトの列数や列幅を変更するには、\n'+
                'ファイル(F) - プロジェクト設定(F)から行ってください。').pack()
        lf1 = Labelframe(f1, text='デフォルト列', padding=5)
        lf1.pack(fill=X)

        ### 列数
        Label(lf1, text='列数: ').grid(row=0, column=0, columnspan=2, sticky=E)
        lf1.grid_columnconfigure(2, weight=1)
        self.setting_number_of_columns = Spinbox(lf1, from_=1, to=10,)
        self.setting_number_of_columns.insert(0, number_of_columns)
        self.setting_number_of_columns.grid(row=0, column=2, columnspan=1, padx=5, pady=5)

        ### 区切り線
        Separator(lf1, orient=HORIZONTAL).grid(row=1 ,column=0, columnspan=3, sticky=W+E, padx=0, pady=10)

        ### 列幅
        self.setting_column_percentage = []
        self.setting_column_percentage_label = []
        Label(lf1, text='列幅（合計100にしてください）: ').grid(row=2, column=0, rowspan=number_of_columns)
        for i in range(number_of_columns):
            self.setting_column_percentage_label.append(Label(lf1, text=f'列{i+1}: '))
            self.setting_column_percentage.append(Spinbox(lf1, from_=1, to=100))
            try:
                self.setting_column_percentage[i].insert(0, percentage_of_columns[i])
            except IndexError:
                self.setting_column_percentage[i].insert(0, '0')
            self.setting_column_percentage_label[i].grid(row=i+2, column=1)
            self.setting_column_percentage[i].grid(row=i+2, column=2, padx=5, pady=5)

        def reload_setting_column_percentage(e=None, x:str='up'):
            for i in range(len(self.setting_column_percentage)):
                self.setting_column_percentage[i].destroy()
                self.setting_column_percentage_label[i].destroy()
            self.setting_column_percentage.clear()
            self.setting_column_percentage_label.clear()
            number_of_columns = self.setting_number_of_columns.get()
            if x == 'up':
                number_of_columns = int(number_of_columns) + 1
                if number_of_columns == 11: number_of_columns = 10
            elif x == 'down':
                number_of_columns = int(number_of_columns) - 1
                if number_of_columns == 0: number_of_columns = 1
            for i in range(int(number_of_columns)):
                self.setting_column_percentage_label.append(Label(lf1, text=f'列{i+1}: '))
                self.setting_column_percentage.append(Spinbox(lf1, from_=1, to=100))
                try:
                    self.setting_column_percentage[i].insert(0, percentage_of_columns[i])
                except IndexError:
                    self.setting_column_percentage[i].insert(0, '0')
                self.setting_column_percentage_label[i].grid(row=i+2, column=1)
                self.setting_column_percentage[i].grid(row=i+2, column=2, padx=5, pady=5)

        # タブ - 表示
        f2 = Frame(nt, padding=5)
        nt.add(f2, padding=5, text='   表示   ')
        lf2 = Labelframe(f2, text='フォント', padding=5)
        lf2.pack(fill=X)

        ## フォント
        ### フォントファミリー
        Label(lf2, text='ファミリー').pack(anchor=W)
        self.setting_font_family = self.setting_font_family_treeview(lf2)
        self.setting_font_family.pack(fill=X)
        lf2_1 = Frame(lf2, padding=5)
        lf2_2 = Frame(lf2, padding=5)
        lf2_3 = Frame(lf2, padding=5)
        ### フォントサイズ
        Label(lf2_1, text='サイズ').pack(anchor=W)
        self.setting_font_size = Spinbox(lf2_1, from_=1, to=100, wrap=True)
        self.setting_font_size.insert(0, self.font_size)
        self.setting_font_size.pack(anchor=W)
        lf2_1.pack(side=LEFT)
        ### 行間
        Label(lf2_2, text='行間').pack(anchor=W)
        self.setting_between_lines = Spinbox(lf2_2, from_=0, to=100, wrap=True)
        self.setting_between_lines.insert(0, self.between_lines)
        self.setting_between_lines.pack(anchor=W)
        lf2_2.pack(side=LEFT)
        ### 折り返し
        Label(lf2_3, text='*折り返し').pack(anchor=W)
        self.setting_wrap = StringVar()
        self.setting_wrap_menu = OptionMenu(lf2_3, self.setting_wrap, self.wrap, *['none', 'char', 'word'])
        self.setting_wrap_menu.pack(anchor=W)
        lf2_3.pack(side=LEFT)

        ## テーマ
        lf3 = Labelframe(f2, text='テーマ', padding=5)
        lf3.pack(fill=X)

        theme_names = []
        light_theme_names = [t for t in STANDARD_THEMES.keys() if STANDARD_THEMES[t]['type'] == 'light']
        dark_theme_names = [t for t in STANDARD_THEMES.keys() if STANDARD_THEMES[t]['type'] == 'dark']
        theme_names.extend(light_theme_names)
        theme_names.append('------')
        theme_names.extend(dark_theme_names)
        self.setting_theme = StringVar()
        self.setting_theme_menu = OptionMenu(lf3, self.setting_theme, self.themename, *theme_names)
        self.setting_theme_menu.grid(row=0, column=0, pady=5)

        ##その他
        lf4 = Labelframe(f2, text='その他', padding=5)
        lf4.pack(fill=X)
        ### 行番号
        self.setting_display_line_number = BooleanVar()
        self.setting_display_line_number.set(display_line_number)
        Checkbutton(lf4, text='*行番号', variable=self.setting_display_line_number, bootstyle="round-toggle", padding=5).pack(anchor=W)
        ### 選択行の強調
        self.setting_selection_line_highlight = BooleanVar()
        self.setting_selection_line_highlight.set(selection_line_highlight)
        Checkbutton(lf4, text='選択行の強調', variable=self.setting_selection_line_highlight, bootstyle="round-toggle", padding=5).pack(anchor=W)

        # タブ - ステータスバー
        f3 = Frame(nt, padding=5)
        for i in range(4):
            f3.grid_columnconfigure(i+1, weight=1)
        nt.add(f3, padding=5, text='ツールバー・ステータスバー')
        statusbar_elements=[None, '文字数カウンター - 1列目', '文字数カウンター - 2列目',
                            '文字数カウンター - 3列目', '文字数カウンター - 4列目',
                            '文字数カウンター - 5列目', '文字数カウンター - 6列目',
                            '文字数カウンター - 7列目', '文字数カウンター - 8列目',
                            '文字数カウンター - 9列目', '文字数カウンター - 10列目',
                            '行数カウンター - 1列目', '行数カウンター - 2列目',
                            '行数カウンター - 3列目', '行数カウンター - 4列目',
                            '行数カウンター - 5列目', '行数カウンター - 6列目',
                            '行数カウンター - 7列目', '行数カウンター - 8列目',
                            '行数カウンター - 9列目', '行数カウンター - 10列目',
                            'カーソルの現在位置', 'ショートカットキー1',
                            'ショートカットキー2', 'ショートカットキー3',
                            'ボタン - ファイルを開く', 'ボタン - 上書き保存',
                            'ボタン - 名前をつけて保存', 'ステータスバー初期メッセージ']
        self.setting_statusbar_elements_dict = {0: {'var': [], 'menu': []}, 1: {'var': [], 'menu': []}, 2: {'var': [], 'menu': []}, 3: {'var': [], 'menu': []}, 4: {'var': [], 'menu': []}, 5: {'var': [], 'menu': []}}

        Label(f3, text='ツールバー1').grid(row=1, column=0, columnspan=2, sticky=W)
        Label(f3, text='ツールバー2').grid(row=3, column=0, columnspan=2, sticky=W)
        Label(f3, text='ステータスバー1').grid(row=5, column=0, columnspan=2, sticky=W)
        Label(f3, text='ステータスバー2').grid(row=7, column=0, columnspan=2, sticky=W)
        Label(f3, text='ステータスバー3').grid(row=9, column=0, columnspan=2, sticky=W)
        Label(f3, text='ステータスバー4').grid(row=11, column=0, columnspan=2, sticky=W)
        for i in range(24):
            k = i % 4
            if i < 4:
                self.setting_statusbar_elements_dict[4]['var'].append(StringVar())
                self.setting_statusbar_elements_dict[4]['menu'].append(OptionMenu(f3, self.setting_statusbar_elements_dict[4]['var'][k], self.statusbar_elements_dict_converted[4][k], *statusbar_elements))
                self.setting_statusbar_elements_dict[4]['menu'][k].grid(row=2, column=k, padx=5, pady=5, sticky=W+E)
            elif i < 8:
                self.setting_statusbar_elements_dict[5]['var'].append(StringVar())
                self.setting_statusbar_elements_dict[5]['menu'].append(OptionMenu(f3, self.setting_statusbar_elements_dict[5]['var'][k], self.statusbar_elements_dict_converted[5][k], *statusbar_elements))
                self.setting_statusbar_elements_dict[5]['menu'][k].grid(row=4, column=k, padx=5, pady=5, sticky=W+E)
            elif i < 12:
                self.setting_statusbar_elements_dict[3]['var'].append(StringVar())
                self.setting_statusbar_elements_dict[3]['menu'].append(OptionMenu(f3, self.setting_statusbar_elements_dict[3]['var'][k], self.statusbar_elements_dict_converted[3][k], *statusbar_elements))
                self.setting_statusbar_elements_dict[3]['menu'][k].grid(row=6, column=k, padx=5, pady=5, sticky=W+E)
            elif i < 16:
                self.setting_statusbar_elements_dict[2]['var'].append(StringVar())
                self.setting_statusbar_elements_dict[2]['menu'].append(OptionMenu(f3, self.setting_statusbar_elements_dict[2]['var'][k], self.statusbar_elements_dict_converted[2][k], *statusbar_elements))
                self.setting_statusbar_elements_dict[2]['menu'][k].grid(row=8, column=k, padx=5, pady=5, sticky=W+E)
            elif i < 20:
                self.setting_statusbar_elements_dict[1]['var'].append(StringVar())
                self.setting_statusbar_elements_dict[1]['menu'].append(OptionMenu(f3, self.setting_statusbar_elements_dict[1]['var'][k], self.statusbar_elements_dict_converted[1][k], *statusbar_elements))
                self.setting_statusbar_elements_dict[1]['menu'][k].grid(row=10, column=k, padx=5, pady=5, sticky=W+E)
            elif i < 24:
                self.setting_statusbar_elements_dict[0]['var'].append(StringVar())
                self.setting_statusbar_elements_dict[0]['menu'].append(OptionMenu(f3, self.setting_statusbar_elements_dict[0]['var'][k], self.statusbar_elements_dict_converted[0][k], *statusbar_elements))
                self.setting_statusbar_elements_dict[0]['menu'][k].grid(row=12, column=k, padx=5, pady=5, sticky=W+E)

        # キーバインド
        self.setting_number_of_columns.bind('<<Increment>>', lambda e:reload_setting_column_percentage(e, 'up'))
        self.setting_number_of_columns.bind('<<Decrement>>', lambda e:reload_setting_column_percentage(e, 'down'))

        # 設定ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)

    # 設定編集GUIのOKや適用を押した際の動作
    def save(self, e=None, close=False):
        print('save_settings')
        number_of_columns = int(self.setting_number_of_columns.get())
        percentage_of_columns = []
        for i in range(len(self.setting_column_percentage)):
            percentage_of_columns.append(int(self.setting_column_percentage[i].get()))
        for _ in range(10-len(percentage_of_columns)):
            percentage_of_columns.append(0)
        font_family = self.tv.selection()[0]
        font_size = int(self.setting_font_size.get())
        between_lines = int(self.setting_between_lines.get())
        wrap = self.setting_wrap.get()
        themename = self.setting_theme.get()
        display_line_number = self.setting_display_line_number.get()
        selection_line_highlight = self.setting_selection_line_highlight.get()
        if themename == '------': themename = self.themename
        statusbar_method = {0: [], 1: [], 2: [], 3: [], 4: [], 5: []}
        for i in range(24):
            k = i % 4
            if   i < 4 :
                e = self.convert_statusbar_elements(self.setting_statusbar_elements_dict[4]['var'][k].get())
                if e: statusbar_method[4].append(e)
            elif i < 8 :
                e = self.convert_statusbar_elements(self.setting_statusbar_elements_dict[5]['var'][k].get())
                if e: statusbar_method[5].append(e)
            elif i < 12:
                e = self.convert_statusbar_elements(self.setting_statusbar_elements_dict[3]['var'][k].get())
                if e: statusbar_method[3].append(e)
            elif i < 16:
                e = self.convert_statusbar_elements(self.setting_statusbar_elements_dict[2]['var'][k].get())
                if e: statusbar_method[2].append(e)
            elif i < 20:
                e = self.convert_statusbar_elements(self.setting_statusbar_elements_dict[1]['var'][k].get())
                if e: statusbar_method[1].append(e)
            elif i < 24:
                e = self.convert_statusbar_elements(self.setting_statusbar_elements_dict[0]['var'][k].get())
                if e: statusbar_method[0].append(e)
        self.settings['columns']['number'] = number_of_columns
        self.settings['columns']['percentage'] = percentage_of_columns
        self.settings['font']['family'] = font_family
        self.settings['font']['size'] = font_size
        self.settings['between_lines'] = between_lines
        self.settings['wrap'] = wrap
        self.settings['themename'] = themename
        self.settings['display_line_number'] = display_line_number
        self.settings['statusbar_element_settings'] = statusbar_method
        self.settings['selection_line_highlight'] = selection_line_highlight
        self.settings['version'] = __version__
        app.font.config(family=font_family, size=font_size)
        app.between_lines = between_lines
        app.windowstyle.theme_use(themename)
        # app.statusbar
        app.selection_line_highlight = selection_line_highlight
        with open(r'.\settings.yaml', mode='wt', encoding='utf-8') as f:
            yaml.dump(self.settings, f, allow_unicode=True)
        MessageDialog(  '一部の設定（*アスタリスク付きの項目）の反映にはアプリの再起動が必要です\n'+
                        '設定をすぐに反映したい場合、いったんアプリを終了して再起動してください',
                        'CastellaEditor - 設定',
                        ['OK']).show()
        if close: self.destroy()

    # フォントファミリー選択ツリービュー
    def setting_font_family_treeview(self, master):
        f1 = Frame(master)
        f1.pack(fill=BOTH, expand=YES)
        self.tv = Treeview(f1, height=10, show=[], columns=[0])
        # 使用可能なフォントを取得
        families = set(font.families())
        # 縦書きフォントを除外
        for f in families.copy():
            if f.startswith('@'):
                families.remove(f)
        # 名前順に並び替え
        families = sorted(families)
        # ツリービューに挿入
        for f in families:
            self.tv.insert('', iid=f, index=END, tags=[f], values=[f])
            self.tv.tag_configure(f, font=(f, 10))
        # 初期位置を設定
        self.tv.selection_set(self.font_family)
        self.tv.see(self.font_family)
        # スクロールバーを設定
        vbar = Scrollbar(f1, command=self.tv.yview, orient=VERTICAL, bootstyle='rounded')
        vbar.pack(side=RIGHT, fill=Y)
        self.tv.configure(yscrollcommand=vbar.set)
        return self.tv

    def convert_statusbar_elements(self, val: str):
        l =[['文字数カウンター - 1列目', 'letter_count_1'],
            ['文字数カウンター - 2列目', 'letter_count_2'],
            ['文字数カウンター - 3列目', 'letter_count_3'],
            ['文字数カウンター - 4列目', 'letter_count_4'],
            ['文字数カウンター - 5列目', 'letter_count_5'],
            ['文字数カウンター - 6列目', 'letter_count_6'],
            ['文字数カウンター - 7列目', 'letter_count_7'],
            ['文字数カウンター - 8列目', 'letter_count_8'],
            ['文字数カウンター - 9列目', 'letter_count_9'],
            ['文字数カウンター - 10列目', 'letter_count_10'],
            ['行数カウンター - 1列目', 'line_count_1'],
            ['行数カウンター - 2列目', 'line_count_2'],
            ['行数カウンター - 3列目', 'line_count_3'],
            ['行数カウンター - 4列目', 'line_count_4'],
            ['行数カウンター - 5列目', 'line_count_5'],
            ['行数カウンター - 6列目', 'line_count_6'],
            ['行数カウンター - 7列目', 'line_count_7'],
            ['行数カウンター - 8列目', 'line_count_8'],
            ['行数カウンター - 9列目', 'line_count_9'],
            ['行数カウンター - 10列目', 'line_count_10'],
            ['カーソルの現在位置', 'current_place'],
            ['ショートカットキー1', 'hotkeys1'],
            ['ショートカットキー2', 'hotkeys2'],
            ['ショートカットキー3', 'hotkeys3'],
            ['ボタン - ファイルを開く', 'toolbutton_open'],
            ['ボタン - 上書き保存', 'toolbutton_over_write_save'],
            ['ボタン - 名前をつけて保存', 'toolbutton_save_as'],
            ['ステータスバー初期メッセージ', 'statusbar_message']]
        for x in l:
            if val == x[0]: return x[1]
            elif val == x[1]: return x[0]
        return ''

class ProjectFileSettingWindow(Toplevel):
    def __init__(self, title="CastellaEditor - プロジェクト設定", iconphoto='', size=(600, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        if app.get_current_data() != app.data:
            self.withdraw()
            mg = MessageDialog('プロジェクトファイル設定を変更する前にプロジェクトファイルを保存します',
                        'CastellaEditor - プロジェクトファイル設定',
                        ['OK:success', 'キャンセル:secondary'],)
            mg.show()
            print(mg.result)
            if mg.result == 'OK':
                x = app.file_over_write_save()
                if x: self.deiconify()
                else:
                    self.destroy()
                    return
            elif mg.result == 'キャンセル':
                self.destroy()
                return

        under = Frame(self)
        under.pack(side=BOTTOM, fill=X, pady=10)
        Button(under, text='キャンセル', command=self.destroy).pack(side=RIGHT, padx=10)
        Button(under, text='OK', command=lambda:self.save(close=True)).pack(side=RIGHT, padx=10)
        Button(under, text='適用', command=lambda:self.save()).pack(side=RIGHT, padx=10)

        self.f0 = Frame(self, padding=5)
        self.f0.pack(fill=BOTH)
        self.f1 = Frame(self.f0)
        self.f1.pack(fill=BOTH)
        Label(self.f1, text='列数: ').grid(row=0, column=0, columnspan=2, sticky=E)
        self.number_of_columns = StringVar()
        self.number_of_columns.set(value=app.number_of_columns)
        self.colmun_percentage = []
        for i in range(10):
            self.colmun_percentage.append(StringVar())
            self.colmun_percentage[i].set(int(app.column_percentage[i]*100))

        ### 列数
        sb_n = Spinbox(self.f1, from_=1, to=10, textvariable=self.number_of_columns, command=self.make_sb_p, state=READONLY, takefocus=False)
        sb_n.grid(row=0, column=2, columnspan=1, padx=5, pady=5)

        ### 区切り線
        Separator(self.f1, orient=HORIZONTAL).grid(row=1 ,column=0, columnspan=3, sticky=W+E, padx=0, pady=10)

        ### 列幅
        self.sb_p = []
        self.sb_p_l = []
        Label(self.f1, text='列幅（合計100にしてください）: ').grid(row=2, column=0, rowspan=self.number_of_columns.get())
        self.make_sb_p()


        # ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)

    def make_sb_p(self, e=None):
        for i in range(len(self.sb_p)):
            self.sb_p[i].destroy()
            self.sb_p_l[i].destroy()
        self.sb_p.clear()
        self.sb_p_l.clear()
        for i in range(int(self.number_of_columns.get())):
            self.sb_p_l.append(Label(self.f1, text=f'列{i+1}: '))
            self.sb_p_l[i].grid(row=i+2, column=1)
            self.sb_p.append(Spinbox(self.f1, from_=0, to=100, textvariable=self.colmun_percentage[i], state=READONLY, takefocus=False))
            self.sb_p[i].grid(row=i+2, column=2, padx=5, pady=5)

    def save(self, e=None, close=False):
        print('\nsave')
        number_of_columns = int(self.number_of_columns.get())
        percentage_of_columns = []
        for i in range(len(self.colmun_percentage)):
            percentage_of_columns.append(int(self.colmun_percentage[i].get())/100)
        for _ in range(10-len(percentage_of_columns)):
            percentage_of_columns.append(0)
        print(number_of_columns)
        print(percentage_of_columns)
        app.number_of_columns = number_of_columns
        app.column_percentage = percentage_of_columns
        app.make_text_editor()
        if close: self.destroy()


class ThirdPartyNoticesWindow(Toplevel):
    def __init__(self, title="CastellaEditor - ライセンス情報", iconphoto='', size=(1200, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        text = __thirdpartynotices__
        familys = font.families(self)
        if 'Consolas' in familys: family = 'Consolas'
        elif 'SF Mono' in familys: family = 'SF Mono'
        elif 'DejaVu Sans Mono' in familys: family = 'DejaVu Sans Mono'
        else: family = 'Courier'
        widget = ScrolledText(self, 10)
        widget.insert(END, text)
        pattern = re.compile('\n---------------------------------------------------------\n\n(.+)\n(.+)')
        widget.tag_config('name', font=font.Font(family=family, size=18, weight='bold'), spacing2=10)
        def open_url(url):
            def inner(_):
                webbrowser.open_new(url)
                print(url)
            return inner
        for m in pattern.finditer(text):
            index = widget.search(m.group(), 0.0)
            widget.tag_add('name', index+'+3lines', index+'+3lines lineend')
            widget.tag_config(m.groups()[1], font=font.Font(family=family, size=14, underline=True), spacing2=10)
            widget.tag_add(m.groups()[1], index+'+4lines', index+'+4lines lineend')
            command = open_url(m.groups()[1])
            widget.tag_bind(m.groups()[1], '<Button-1>', command)
        print(widget.tag_names())
        widget.winfo_children()[0].config(font=font.Font(family=family, size=12), spacing2=10, state=DISABLED)
        widget.pack(fill=BOTH, expand=True)

        # ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)


class HelpWindow(Toplevel):
    def __init__(self, title="ヘルプ", iconphoto='', size=(600, 500), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        f = Frame(self, padding=5)
        f.pack(fill=BOTH, expand=True)

        # ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)


class AboutWindow(Toplevel):
    def __init__(self, title="CastellaEditorについて", iconphoto='', size=(600, 500), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        frame = Frame(self, padding=5)
        frame.pack(fill=BOTH, expand=True)
        main = Text(frame)

        main.tag_config('title', justify=CENTER, spacing2=10, font=font.Font(size=15, weight='bold'))
        main.tag_config('text', justify=CENTER, spacing2=10, font=font.Font(size=12, weight='normal'))
        main.tag_config('link', justify=CENTER, spacing2=10, font=font.Font(size=12, weight='normal', underline=True))
        github = 'https://github.com/joppincal/CastellaEditor'
        main.tag_config('github')
        main.tag_bind('github', '<Button-1>', lambda e: webbrowser.open_new(github))
        homepage = ''
        main.tag_config('homepage')
        main.tag_bind('homepage', '<Button-1>', lambda e: webbrowser.open_new(homepage))

        main.insert(END, 'ここにアイコンを表示\n')
        main.insert(END, 'カステラエディタ\nCastellaEditor\n', 'title')
        main.insert(END, f'\nバージョン: {__version__}\n', 'text')
        main.insert(END, f'プロジェクトファイルバージョン: {__projversion__}\n', 'text')
        main.insert(END, '\nGithub: ', 'text')
        main.insert(END, github+'\n', ('link', 'github'))
        main.insert(END, 'Homepage: ', 'text')
        main.insert(END, homepage+'\n', ('link', 'homepage'))

        main.config(state=DISABLED, takefocus=False)
        main.pack(fill=BOTH, expand=True)

        # ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)


if __name__ == '__main__':
    root = ttk.Window(title='CastellaEditor', minsize=(1200, 800))
    app = Main(master=root)
    app.mainloop()
