'''
SoroEditor - Joppincal
This software is distributed under the MIT License. See LICENSE for details.
See ThirdPartyNotices.txt for third party libraries.
'''
from collections import deque
import datetime
import difflib
import logging
import logging.handlers
import os
from random import choice
import re
import sys
import tkinter
from tkinter import BooleanVar, IntVar, StringVar, filedialog, font
import webbrowser
from ttkbootstrap import Button, Checkbutton, Entry, Frame, Label, Labelframe,\
    Menu, Notebook, OptionMenu, PanedWindow, Radiobutton, Scrollbar, Separator, Spinbox, \
    Style, Text, Toplevel, Treeview, Window
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs.dialogs import MessageDialog
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.themes.standard import *
import yaml

__version__ = '0.3.4'
__projversion__ = '0.2.0'
with open(os.path.join(os.path.dirname(__file__), 'ThirdPartyNotices.txt'), 'rt', encoding='utf-8') as f:
    __thirdpartynotices__ = f.read()
__icon__ = ''

def log_setting():
    global log
    if not os.path.exists('./log'):
        os.mkdir('./log')
    log = logging.getLogger(__name__)
    log.setLevel(logging.DEBUG)
    formater = logging.Formatter('{asctime} {name:<8s} {levelname:<8s} {message}', style='{')
    handler = logging.handlers.RotatingFileHandler(
        filename='./log/soroeditor.log',
        encoding='utf-8',
        maxBytes=102400)
    handler.setFormatter(formater)
    log.addHandler(handler)

class Main(Frame):

    def __init__(self, master=None):

        super().__init__(master)
        self.master.geometry('1600x1000')

        self.master.protocol('WM_DELETE_WINDOW', lambda:self.file_close(True))

        sys.argv[0] = self.convert_drive_to_uppercase(sys.argv[0])
        global __file__
        __file__ = self.convert_drive_to_uppercase(__file__)

        # 設定ファイルを読み込み
        os.chdir(os.path.dirname(sys.argv[0]))
        self.settingFile_Error_md = None
        initialization = False
        self.settings ={'autosave': False,
                        'autosave_frequency': 60000,
                        'backup': True,
                        'backup_frequency': 300000,
                        'between_lines': 10,
                        'columns': {'number': 3, 'percentage': [10, 60, 30]},
                        'display_line_number': True,
                        'font': {'family': 'nomal', 'size': 12},
                        'geometry': '1600x1000',
                        'ms_align_the_lines': 100,
                        'recently_files': [],
                        'selection_line_highlight': True,
                        'statusbar_element_settings':
                            {0: ['statusbar_message'],
                            1: ['hotkeys2', 'hotkeys3'],
                            2: ['hotkeys1'],
                            3: ['current_place', 'line_count_1', 'line_count_2', 'line_count_3'],
                            4: ['toolbutton_open', 'toolbutton_over_write_save', 'toolbutton_save_as'],
                            5: [None]},
                        'themename': '',
                        'version': __version__,
                        'wrap': NONE}
        try:
            with open('./settings.yaml', mode='rt', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                if set(data.keys()) & set(self.settings.keys()) != set(self.settings.keys()):
                    raise KeyError('Settings file is incomplete')
        ## 内容に不備がある場合新規作成し、古い設定ファイルを別名で保存
        except (KeyError, yaml.YAMLError) as e:
            log.error(e)
            self.update_setting_file()
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
            log.error(e)
            self.update_setting_file()
            settingFile_Error_message = '設定ファイルが存在しなかったため作成しました'
            initialization = True
            self.settingFile_Error_md = MessageDialog(message=settingFile_Error_message, buttons=['OK'])
        else:
            self.settings = data
            log.info('Settings loaded successfully')

        # 設定ファイルから各設定を読み込み
        ## ウィンドウサイズ
        if self.settings['geometry'] == 'Full Screen':
            self.master.state('zoomed')
        else:
            try:
                self.master.geometry(self.settings['geometry'])
            except tkinter.TclError as e:
                print(e)
        ## テーマを設定
        self.windowstyle = Style()
        self.windowstyle.theme_use(self.settings['themename'])
        ## フォント設定
        self.font = font.Font(family=self.settings['font']['family'], size=self.settings['font']['size'])
        ## 列数
        self.number_of_columns: int = self.settings['columns']['number']
        ## 列比率
        self.column_percentage = [x*0.01 for x in self.settings['columns']['percentage']]
        for _ in range(10-len(self.column_percentage)):
            self.column_percentage.append(0)
        ## 折り返し
        self.wrap: str = self.settings['wrap']
        ## 行間
        self.between_lines: int = self.settings['between_lines']
        ## 行番号
        self.display_line_number: bool = self.settings['display_line_number']
        ## 選択行の強調
        self.selection_line_highlight: bool = self.settings['selection_line_highlight']
        ## 自動保存機能
        self.do_autosave: bool = self.settings['autosave']
        ## 自動保存頻度
        self.autosave_frequency: int = self.settings['autosave_frequency']
        ## バックアップ機能
        self.do_backup: bool = self.settings['backup']
        ## バックアップ頻度
        self.backup_frequency: int = self.settings['backup_frequency']

        self.filepath = ''
        self.data0 = {}
        for i in range(10):
            self.data0[i] = {'text': '', 'title': ''}
        self.data0['columns'] = {'number': self.number_of_columns, 'percentage': [int(x*100) for x in self.column_percentage]}
        self.data0['version'] = __projversion__
        self.data = self.data0
        self.edit_history = deque([self.data])
        self.undo_history: deque[dict] = deque([])
        self.recently_files: list = self.settings['recently_files']
        try:
            self.initialdir = os.path.dirname(self.recently_files[0])
        except (IndexError, TypeError):
            self.initialdir = ''
        self.text_place = (1.0, 1.0)
        self.ms_align_the_lines = self.settings['ms_align_the_lines']
        self.md_position = (int(self.winfo_screenwidth()/2-250), int(self.winfo_screenheight()/2-100))

        # メニューバーの作成
        menubar = Menu()

        ## メニュバー - ファイル
        self.menu_file = Menu(menubar, tearoff=False)

        self.menu_file.add_command(label='ファイルを開く(O)', command=self.file_open, accelerator='Ctrl+O', underline=8)
        self.menu_file.add_command(label='上書き保存(S)', command=self.file_over_write_save, accelerator='Ctrl+S', underline=6)
        self.menu_file.add_command(label='名前をつけて保存(A)', command=self.file_save_as, accelerator='Ctrl+Shift+S', underline=9)
        self.menu_file.add_command(label='エクスポート(E)', command=ExportWindow, accelerator='Ctrl+Shift+E', underline=7)
        self.menu_file.add_command(label='プロジェクト設定(F)', command=ProjectFileSettingWindow, underline=9)

        ###メニューバー - ファイル - 最近使用したファイル
        self.menu_file_recently = Menu(self.menu_file, tearoff=False)

        try:
            self.menu_file_recently.add_command(label=self.recently_files[0], command=lambda:self.file_open(file=self.recently_files[0]), accelerator='Ctrl+R')
            self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file=self.recently_files[1]))
            self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file=self.recently_files[2]))
            self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file=self.recently_files[3]))
            self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file=self.recently_files[4]))
        except IndexError:
            pass

        self.menu_file.add_cascade(label='最近使用したファイル(R)', menu=self.menu_file_recently, underline=11)

        self.menu_file.add_separator()
        self.menu_file.add_command(label='終了(Q)', command=lambda:self.file_close(shouldExit=True), accelerator='Alt+F4', underline=3)
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
        self.menu_help.add_command(label='初回起動メッセージを表示(F)', command=lambda: self.file_open(file=os.path.join(os.path.dirname(__file__), 'hello.txt')), underline=13)
        self.menu_help.add_command(label='SoroEditorについて(A)', command=AboutWindow, underline=16)
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

        # フォーカスを列1のタイトルにセット
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
        self.hotkeys2 = ('label', '[Ctrl+Enter]: 1行追加(下)  [Ctrl+Shift+Enter]: 1行追加(上)  [Ctrl+^]: 検索・置換  [Ctrl+Z]: 取り消し  [Ctrl+Shift+Z]: 取り消しを戻す')
        self.hotkeys3 = ('label', '[Ctrl+Q][Alt+,]: 左に移る  [Ctrl+W][Alt+.]: 右に移る')
        ## 各機能情報
        self.infomation = ('label', f'自動保存: {self.do_autosave}, バックアップ: {self.do_backup}')
        ## 顔文字
        self.kaomoji = ('label', choice(['( ﾟ∀ ﾟ)', 'ヽ(*^^*)ノ ', '(((o(*ﾟ▽ﾟ*)o)))', '(^^)', '(*^○^*)']))
        ## ツールボタン
        self.toolbutton_open = ('button', 'ファイルを開く[Ctrl+O]', self.file_open)
        self.toolbutton_over_write_save = ('button', '上書き保存[Ctrl+S]', self.file_over_write_save)
        self.toolbutton_save_as = ('button', '名前をつけて保存[Ctrl+Shift+S]', self.file_save_as)
        ## 初期ステータスバーメッセージ
        self.statusbar_message = ('label', 'ステータスバーの表示項目はメニューバーの 設定-ステータスバー から変更できます')

        # ステータスバー作成メソッド
        def make_statusbar_element(elementtype=None, text=None, commmand=None):
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
            elif name == 'infomation': return self.infomation
            elif name == 'kaomoji': return self.kaomoji
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
                    self.statusbar_element_dict[num][i] = make_statusbar_element(*e)
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
                next_text_index = self.maintexts.index(e.widget)+1
                if next_text_index >= len(self.maintexts): next_text_index = 0
                next_text = self.maintexts[next_text_index]
                insert = e.widget.index(INSERT)
                next_text.mark_set(INSERT, insert)
                next_text.focus_set()
        def focus_to_left(e):
            if type(e.widget) == tkinter.Text:
                prev_text_index = self.maintexts.index(e.widget)-1
                prev_text = self.maintexts[prev_text_index]
                insert = e.widget.index(INSERT)
                prev_text.mark_set(INSERT, insert)
                prev_text.focus_set()
        def focus_to_bottom(e):
            if e.widget.winfo_class() == 'TEntry':
                print('O')
                index = self.entrys.index(e.widget)
                self.maintexts[index].focus_set()
        def reload(e):
            statusbar_element_reload()
            if self.selection_line_highlight: self.highlight()
            self.set_text_widget_editable()
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
        self.master.bind('<Alt-.>', focus_to_right)
        self.master.bind('<Control-q>', focus_to_left)
        self.master.bind('<Alt-,>', focus_to_left)
        self.master.bind('<Control-l>', self.select_line)
        self.master.bind('<Control-Return>', lambda e: self.newline(e, 1))
        self.master.bind('<Control-Shift-Return>', lambda e: self.newline(e, 0))
        self.master.bind('<Down>', focus_to_bottom)
        self.master.bind('<Return>', focus_to_bottom)
        self.master.bind('<Control-^>', test_focus_get)
        self.master.bind('<Control-Y>', self.print_history)
        self.master.bind('<F1>', lambda e: HelpWindow())
        self.master.bind('<Control-E>', lambda e: ExportWindow())

        # 各パーツを設置
        for w in self.statusbars[0:4]:
            if len(w.panes()):
                w.pack(fill=X, side=BOTTOM, padx=5, pady=3)
        for w in self.statusbars[4:6]:
            if len(w.panes()):
                w.pack(fill=X)

        self.f2.pack(fill=Y, side=RIGHT, pady=18)
        self.vbar.pack(fill=Y, expand=YES)
        self.f1.pack(fill=BOTH, expand=YES)

        # 設定ファイルに異常があり初期化した場合、その旨を通知する
        if self.settingFile_Error_md:
            self.settingFile_Error_md.show(self.md_position)

        # 設定ファイルが存在しなかった場合、初回起動と扱う
        if initialization:
            self.file_open(file=os.path.join(os.path.dirname(__file__), 'hello.txt'))

        # ファイルを渡されているとき、そのファイルを開く
        if len(sys.argv) > 1:
            self.file_open(file=sys.argv[1])

        self.master.after(1000, self.change_window_title)
        self.master.after(self.ms_align_the_lines, self.align_the_lines)
        self.master.after(self.autosave_frequency, self.autosave)
        self.master.after(self.backup_frequency, self.backup)


    def convert_drive_to_uppercase(self, file_path: str) -> str:
        """
        ファイルパスの先頭にあるドライブ名を大文字に変換した文字列を返す関数。
        """
        drive, tail = os.path.splitdrive(file_path)
        drive = drive.upper()
        return os.path.join(drive, tail)

    def make_text_editor(self):
        '''
        テキストエディタ部分の要素の準備
        '''
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
            self.line_number_box_entry = Entry(self.f1_1, font=self.font, width=3,
                                                state=DISABLED, takefocus=NO)
            self.line_number_box = Text(self.f1_1, font=self.font,
                                        width=4,
                                        spacing3=self.between_lines,
                                        takefocus=NO)
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
                                    font=self.font,
                                    spacing3=self.between_lines))

        for i in range(self.number_of_columns):
            # 列の枠を作成
            j, relx = i, 0.0
            while j: relx, j = relx + self.column_percentage[j-1], j - 1
            self.columns[i].place(relx=relx, rely=0.0, relwidth=self.column_percentage[i], relheight=1.0)
            # Entryを作成
            self.entrys[i].pack(fill=X, expand=False, padx=5, pady=0)
            # Textを作成
            self.maintexts[i].pack(fill=BOTH, expand=YES, padx=5, pady=10)

        ## ダミーテキスト（スクロールバーのための）
        self.dummy_column = (Frame(self.f1_2, padding=0))
        self.dummy_entry = (Entry(self.dummy_column, font=self.font))
        self.dummy_maintext = (Text(self.dummy_column, wrap=self.wrap,
                                font=self.font,
                                spacing3=self.between_lines))
        self.dummy_column.place(relx=1.0, rely=0.0, relwidth=1.0, relheight=1.0)
        self.dummy_entry.pack(fill=X, expand=False, padx=5, pady=0)
        self.dummy_maintext.pack(fill=BOTH, expand=YES, padx=5, pady=10)

        for i in range(self.number_of_columns):
            self.entrys[i].insert(END, self.data[i]['title'])
            self.maintexts[i].insert(END, self.data[i]['text'])

        self.align_number_of_rows()

        self.textboxes = self.maintexts + [self.dummy_maintext]
        try:
            self.textboxes.append(self.line_number_box)
        except AttributeError as e:
            print(e)

        for w in self.maintexts:
            w.bind('<Button-3>', self.popup)

        self.set_text_widget_editable(mode=2)

    def align_number_of_rows(self, e=None):
        '''
        各列の行数を揃えるメソッド
        '''
        self.set_text_widget_editable(mode=1)
        # 各列の行数を揃える
        for i in range(self.number_of_columns):
            # 行数の差分を計算する
            line_count_diff = 10000 - self.line_count(self.maintexts[i])
            if line_count_diff > 0:
                # 行数が足りない場合、不足分の改行を追加する
                self.maintexts[i].insert(END, '\n' * line_count_diff)
        # カーソルを最初の列に移動する
        if e:
            self.maintexts[0].see(INSERT)
        self.set_text_widget_editable()

    # 以下、ファイルを開始、保存、終了に関するメソッド
    def file_open(self, e=None, file: str|None=None):
        '''
        '''
        self.file_close()
        # 開くファイルを指定されている場合、ファイル選択ダイアログをスキップする
        if file:
            # ドライブ文字を大文字にする
            newfilepath = self.convert_drive_to_uppercase(file)
        else:
            newfilepath = filedialog.askopenfilename(
                title = '編集ファイルを選択',
                initialdir = self.initialdir,
                initialfile='file.sep',
                filetypes=[('SoroEditorプロジェクトファイル', '.sep'), ('YAMLファイル', '.yaml'), ('その他', '.*'), ('ThreeCrowsプロジェクトファイル', '.tcs'), ('CastellaEditorプロジェクトファイル', '.cep')],
                defaultextension='sep')

        if newfilepath:
            if os.path.exists(newfilepath):
                self.filepath = newfilepath
            else:
                return

        log.info(f'Opening file: {newfilepath}')

        if self.filepath:
            try:
                with open(self.filepath, mode='rt', encoding='utf-8') as f:
                    newdata: dict = yaml.safe_load(f)
                    if 'version' in newdata.keys():
                        self.data = newdata
                    else:
                        raise(KeyError)
                    data = self.data0.copy()
                    data.update(newdata)
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

                    try:
                        self.make_text_editor()
                    except KeyError:
                        self.data = self.edit_history[0]
                        self.make_text_editor()
                        raise KeyError

                    self.edit_history.clear()

                    for i in range(self.number_of_columns, 10):
                        mg_exist_hidden_data = MessageDialog(f'列{i+1}にデータがありますが、表示されません\n'
                                                            f'title: {data[i]["title"]}\n'
                                                            f'text: {data[i]["text"][0:10]}...\n'
                                                            '確認するにはプロジェクトファイルを普通のテキストエディタで開いて直接データを見るか、'
                                                            'ファイル - プロジェクト設定 からこのファイルで表示する列を変更し、'
                                                            f'列{self.number_of_columns+i-1}が見えるようにしてください', 'SoroEditor - 隠されたデータ')
                        if data[i]['title']: mg_exist_hidden_data.show(self.md_position)
                        elif data[i]['text']: mg_exist_hidden_data.show(self.md_position)
                    self.entrys[0].focus_set()
                    self.edit_history.appendleft(data)

                self.initialdir = os.path.dirname(self.filepath)

                # 最近使用したファイルのリストを修正し、settings.yamlに反映
                self.recently_files.insert(0, self.filepath)
                # '\'を'/'に置換
                self.recently_files = [v.replace('\\', '/') for v in self.recently_files]
                # 重複を削除
                self.recently_files = list(dict.fromkeys(self.recently_files))
                # 5つ以内になるよう削除
                del self.recently_files[4:-1]
                self.settings['recently_files'] = self.recently_files
                # 設定ファイルに書き込み
                self.update_setting_file()
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

            except FileNotFoundError as e:
                log.error(e)
                md = MessageDialog(title='TreeCrows - エラー', alert=True, buttons=['OK'], message='ファイルが見つかりません')
                md.show(self.md_position)
                # 最近使用したファイルに見つからなかったファイルが入っている場合、削除しsettins.yamlに反映する
                if newfilepath in self.recently_files:
                    self.recently_files.remove(newfilepath)
                    self.settings['recently_files'] = self.recently_files
                # 設定ファイルに書き込み
                self.update_setting_file()
                # 更新
                try:
                    for _ in range(5):
                        self.menu_file_recently.delete(0)
                except tkinter.TclError:
                    pass
                try:
                    if self.filepath:
                        self.menu_file_recently.add_command(label=self.filepath + ' (現在のファイル)')
                    else:
                        self.menu_file_recently.add_command(label=self.recently_files[0], command=lambda:self.file_open(file=self.recently_files[0]))
                    self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file=self.recently_files[1]))
                    self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file=self.recently_files[2]))
                    self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file=self.recently_files[3]))
                    self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file=self.recently_files[4]))
                except IndexError:
                    pass
                return

            except (KeyError, UnicodeDecodeError, yaml.scanner.ScannerError) as e:
                log.error(f'Cannot open file: {e}')
                # ファイルが読み込めなかった場合
                md = MessageDialog(
                    title='TreeCrows - エラー',
                    alert=True,
                    buttons=['OK'],
                    message='ファイルが読み込めませんでした。\nファイルが破損している、またはSoroEditorプロジェクトファイルではない可能性があります')
                md.show(self.md_position)

            else:
                log.info(f'Success opening file: {self.filepath}')
                self.align_the_lines(1.0, 1.0, False)

    def file_save_as(self, e=None) -> bool:
        self.filepath = filedialog.asksaveasfilename(
            title='名前をつけて保存',
            initialdir=self.initialdir,
            initialfile='noname',
            filetypes=[('SoroEditorプロジェクトファイル', '.sep'), ('YAMLファイル', '.yaml'), ('その他', '.*')],
            defaultextension='sep'
        )
        if self.filepath:
            print(self.filepath)
            self.save_file(self.filepath)
            return True
        else:
            return False

    def file_over_write_save(self, e=None) -> bool:
        if self.filepath:
            self.save_file(self.filepath)
            return True
        else:
            if self.file_save_as():
                return True
            else:
                return False

    def save_file(self, file_path):
        '''
        データを保存する
        '''
        try:
            with open(file_path, mode='wt', encoding='utf-8') as f:
                # 現在のデータを読み込む
                current_data = self.get_current_data()
                # ファイルに書き込む
                yaml.safe_dump(current_data, f, encoding='utf-8', allow_unicode=True, canonical=True)
                # 変更前のデータを更新する（変更検知に用いられる）
                self.data = current_data
        except (FileNotFoundError, UnicodeDecodeError, yaml.YAMLError) as e:
            error_type = type(e).__name__
            error_message = str(e)
            log.error(f"An error of type {error_type} occurred while writing to the file {file_path}: {error_message}")
        else:
            log.info(f'Save file: {file_path}')

        self.update_recently_files()

    def update_setting_file(self):
        '''
        設定ファイルをMain.settingsの内容に書き換える
        '''
        try:
            with open('./settings.yaml', mode='wt', encoding='utf-8') as f:
                yaml.dump(self.settings, f, allow_unicode=True)
        except (FileNotFoundError, UnicodeDecodeError, yaml.YAMLError) as e:
            error_type = type(e).__name__
            error_message = str(e)
            log.error(f"An error of type {error_type} occurred while saving setting file: {error_message}")
        else:
            log.info('Update setting file')

    def update_recently_files(self):
        '''
        最近使用したファイルのリストを修正し、settings.yamlに反映する
        '''
        self.recently_files.insert(0, self.filepath)
        # 重複を削除
        self.recently_files = list(dict.fromkeys(self.recently_files))
        # 5つ以内になるよう削除
        del self.recently_files[5:]
        # 設定に反映
        self.settings['recently_files'] = self.recently_files
        # 設定ファイルに書き込み
        self.update_setting_file()
        # メニューを更新
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
            title = 'SoroEditor - 終了確認'
            buttons=['保存して終了:success', '保存せず終了:danger', 'キャンセル']
        else:
            message = '更新内容が保存されていません。ファイルを閉じ、ファイルを変更しますか'
            title = 'SoroEditor - 確認'
            buttons=['保存して変更:success', '保存せず変更:danger', 'キャンセル']

        if self.data == current_data:
            if shouldExit:
                self.master.destroy()
        elif self.data != current_data:
            mg = MessageDialog(message=message, title=title, buttons=buttons)
            mg.show(self.md_position)
            if mg.result in ('保存せず終了', '保存せず変更'):
                if shouldExit:
                    self.master.destroy()
            elif mg.result in ('保存して終了', '保存して変更'):
                if self.file_over_write_save():
                    if shouldExit:
                        self.master.destroy()
            else:
                pass

    def open_last_file(self, e=None):
        if not self.filepath:
            self.file_open(file=self.recently_files[0])

    # 以下、エディターの編集に関するメソッド
    def letter_count(self, obj:ScrolledText|Text|str=''):
        if type(obj) == ScrolledText or Text:
            return len(obj.get('1.0', 'end-1c').replace('\n', ''))
        elif type(obj) == str:
            return len(obj)

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

    def recode_edit_history(self, e=None):
        current_data = self.get_current_data()
        if current_data != self.edit_history[0]:
            self.edit_history.appendleft(current_data)
            self.undo_history.clear()

    def undo(self, e=None):

        if len(self.edit_history) <= 1:
            return

        data_new = self.edit_history[0]
        data_old = self.edit_history[1]

        self.set_text_widget_editable(mode=1)

        for i in range(10):
            if data_new[i] != data_old[i]:
                if data_new[i]['title'] != data_old[i]['title']:
                    self.entrys[i].delete(0, END)
                    self.entrys[i].insert(END, data_old[i]['title'])
                elif data_new[i]['text'] != data_old[i]['text']:
                    diff = difflib.SequenceMatcher(None, data_new[i]['text'].splitlines(True), data_old[i]['text'].splitlines(True))
                    for t in diff.get_opcodes():
                        if t[0] in ('replace', 'delete', 'insert'):
                            if t[0] == 'replace':
                                self.maintexts[i].delete(f'{float(t[1]+1)} linestart', f'{float(t[2]+1)} linestart')
                                for s in reversed(data_old[i]['text'].splitlines(True)[t[3]:t[4]]):
                                    self.maintexts[i].insert(f'{float(t[1]+1)} linestart', s)
                            elif t[0] == 'delete':
                                self.maintexts[i].delete(f'{float(t[1]+1)} linestart', f'{float(t[2]+1)} linestart')
                            elif t[0] == 'insert':
                                for s in reversed(data_old[i]['text'].splitlines(True)[t[3]:t[4]]):
                                    self.maintexts[i].insert(f'{float(t[1]+1)} linestart', s)
                        elif t[0] == 'equal':
                            pass

        self.undo_history.appendleft(self.edit_history.popleft())
        self.set_text_widget_editable(mode=0)

    def repeat(self, e=None):

        if len(self.undo_history) == 0:
            return

        data_new = self.edit_history[0]
        data_old = self.undo_history[0]

        self.set_text_widget_editable(mode=1)

        for i in range(10):
            if data_new[i] != data_old[i]:
                if data_new[i]['title'] != data_old[i]['title']:
                    self.entrys[i].delete(0, END)
                    self.entrys[i].insert(END, data_old[i]['title'])
                elif data_new[i]['text'] != data_old[i]['text']:
                    diff = difflib.SequenceMatcher(None, data_new[i]['text'].splitlines(True), data_old[i]['text'].splitlines(True))
                    for t in diff.get_opcodes():
                        if t[0] == 'replace':
                            self.maintexts[i].delete(f'{float(t[1]+1)} linestart', f'{float(t[2]+1)} linestart')
                            for s in reversed(data_old[i]['text'].splitlines(True)[t[3]:t[4]]):
                                self.maintexts[i].insert(f'{float(t[1]+1)} linestart', s)
                        elif t[0] == 'delete':
                            self.maintexts[i].delete(f'{float(t[1]+1)} linestart', f'{float(t[2]+1)} linestart')
                        elif t[0] == 'insert':
                            for s in reversed(data_old[i]['text'].splitlines(True)[t[3]:t[4]]):
                                self.maintexts[i].insert(f'{float(t[1]+1)} linestart', s)
                        elif t[0] == 'equal':
                            pass

        self.edit_history.appendleft(self.undo_history.popleft())
        self.set_text_widget_editable(mode=0)

    def print_history(self, e=None):
        print(self.edit_history)

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

    def newline(self, e, mode=0):
        '''
        行を追加する

        mode: int=[0, 1]

        mode=0: 上に追加
        mode=1: 下に追加
        '''
        if type(e.widget) == tkinter.Text:

            self.set_text_widget_editable(mode=1)

            e.widget.delete(INSERT+'-1c', INSERT)
            if mode == 0: insert = e.widget.index(INSERT+ ' linestart')
            elif mode == 1: insert = e.widget.index(INSERT+ '+1line linestart')
            for w in self.maintexts:
                w.insert(insert, '\n')
                w.see(e.widget.index(INSERT))
            if mode == 0: e.widget.mark_set(INSERT, INSERT+'-1line linestart')
            elif mode == 1: e.widget.mark_set(INSERT, INSERT+'+1line linestart')

            self.set_text_widget_editable(mode=0)

    def highlight(self):
        '''
        カーソルのある行を強調するメソッド
        '''
        if self.master.focus_get() in self.maintexts:
            insert = self.master.focus_get().index(INSERT)
            hilight_font = self.font.copy()
            hilight_font.config(weight='normal', slant='roman', underline=True)
            for w in self.maintexts: # self.textboxesにし、Main.line_number_boxにも適用しようとしたところ、日本語入力の再変換候補があらぶってしまうため断念
                w.mark_set(INSERT, insert)
                w.tag_delete('insert_line')
                w.tag_add('insert_line', insert+' linestart', insert+' lineend')
                w.tag_config('insert_line', underline=False, font=hilight_font)
            if self.display_line_number: # 上記の問題はこのようにする事で解決。Main.textboxesに含まれるMain.dummy_maintextの影響か？
                self.line_number_box.mark_set(INSERT, insert)
                self.line_number_box.tag_delete('insert_line')
                self.line_number_box.tag_add('insert_line', insert+' linestart', insert+' lineend')
                self.line_number_box.tag_config('insert_line', underline=False, font=hilight_font)

    def vbar_command(self, e=None, a=None, b=None):
        '''
        縦スクロールバーが動いたときのメソッド
        '''
        self.dummy_maintext.yview(e, a, b)

    def align_the_lines(self, top:float=None, bottom:float=None, repeat=True):
        '''
        位置を合わせる
        '''
        height = self.textboxes[0].winfo_height()

        # 各行の表示位置を確認し、変更を検知する
        if top and bottom:
            self.text_place = (top, bottom)
        else:
            l = [(float((w.index('@0,0'))), float(w.index(f'@0,{height}')))
                            for w in self.textboxes]
            text_places = []
            for t in l:
                if t[0] != self.text_place[0]:
                    text_places.append(t)
            if not text_places:
                self.text_place = l[0]

            if len(text_places) == 1: self.text_place = text_places[0]

        # ダミーテキストの行数を調整
        self.set_dummy_text_lines()

        if self.text_place[1] > self.line_count(self.dummy_maintext):
            self.text_place = list(self.text_place)
            self.text_place[1] = float(self.line_count(self.dummy_maintext))
            self.text_place = tuple(self.text_place)

        # 場所を設定
        for w in self.textboxes:
            w.see(self.text_place[1])
            w.see(self.text_place[1])
            w.see(self.text_place[1])
            w.see(self.text_place[0])
            w.see(self.text_place[0])
            w.see(self.text_place[0])

        # ずれを検知し、修正
        if not all([True
                    if f == self.maintexts[0].index('@0,0')
                    else False
                    for f
                    in [w.index('@0,0') for w in self.textboxes]]):
            for w in self.textboxes:
                w.see('0.0')
                w.see(self.text_place[1])
                w.see(self.text_place[1])
                w.see(self.text_place[1])
                w.see(self.text_place[0])
                w.see(self.text_place[0])
                w.see(self.text_place[0])
            self.text_place = (float((w.index('@0,0'))), float(w.index(f'@0,{height}')))
        self.show_cursor_at_bottom_line()
        if repeat:
            self.master.after(self.ms_align_the_lines, self.align_the_lines)

        # 縦スクロールバーを設定
        self.vbar.set(*self.dummy_maintext.yview())

    def show_cursor_at_bottom_line(self, e=None):
        '''
        カーソルが最下端の行にあり、最下端の行が隠れているとき、表示する
        '''
        widget = self.master.focus_get()
        if not widget:
            widget = self.maintexts[0]
        if widget.winfo_class() == 'Text':
            height = widget.winfo_height()
            bottom = widget.index(f'@0,{height}')
            if bottom == widget.index(INSERT+' linestart'):
                for w in self.textboxes:
                    w.see(float(bottom)+0.0)

    def set_dummy_text_lines(self):
        '''
        ダミーテキストの行数を調整するメソッド
        '''
        # メインテキストの中で最も多い行数に50行を加えた値を計算する
        max_line_count = max([self.line_count(w) for w in self.maintexts]) + 50
        # ダミーテキストの行数を計算する
        dummy_line_count = self.line_count(self.dummy_maintext)
        # メインテキストの行数とダミーテキストの行数の差分を計算する
        line_count_diff = dummy_line_count - max_line_count
        # ダミーテキストの行数がメインテキストの行数よりも多い場合、超過分だけ行を削除する
        if line_count_diff > 0:
            for _ in range(line_count_diff):
                self.dummy_maintext.delete(0.0, float(line_count_diff))
        # ダミーテキストの行数がメインテキストの行数よりも少ない場合、不足分だけ行を追加する
        if line_count_diff < 0:
            line_count_diff = -line_count_diff
            for _ in range(line_count_diff):
                self.dummy_maintext.insert(END, f'd\n')

    def set_text_widget_editable(self, e=None, mode=0, widget=None):
        '''
        mode: int=[0, 1, 2]
        mode=0: 編集中でないテキストを無効にする
        mode=1: すべてのテキストを有効にする
        mode=2: すべてのテキストを無効にする
        '''
        if mode == 0:
            if not widget: widget = self.master.focus_get()
            if not widget: return
            for w in self.maintexts:
                if w != widget:
                    w.config(state=DISABLED)
                if w == widget:
                    w.config(state=NORMAL)
        elif mode == 1:
            for w in self.maintexts:
                w.config(state=NORMAL)
        elif mode == 2:
            for w in self.maintexts:
                w.config(state=DISABLED)

    def change_window_title(self):
        '''
        更新がある場合に更新ありと表示する
        '''
        current_data = self.get_current_data()
        name = 'SoroEditor'
        if self.filepath:
            name = name + ' - ' + self.filepath
        else:
            name = name + ' - ' + '(無題)'
        if current_data != self.data:
            name = name + ' (更新あり)'
        if not self.do_autosave:
            name = name + ' (自動保存無効)'
        elif self.autosave and not self.filepath:
            name = name + ' (ファイル未保存のため自動保存無効)'
        self.master.title(name)
        self.master.after(500, self.change_window_title)

    def backup(self):
        '''
        定期的にバックアップファイルにデータを保存する
        '''

        if not self.do_backup:
            return
        backup_max = 50

        # バックアップファイル名・パスを生成する
        backup_filename = 'backup'
        if self.filepath:
            backup_filename += '-' + os.path.basename(self.filepath)
        backup_filename += '.$ep'
        backup_filepath = os.path.join(os.path.dirname(self.filepath), backup_filename)

        # 現在のデータを取得する
        current_data = self.get_current_data()

        # データをyaml形式の文字列に変換する
        new_data = yaml.safe_dump(current_data, allow_unicode=True, canonical=True)

        # 新しいデータにタイムスタンプを追加する
        timestamp = datetime.datetime.now().strftime('# %Y-%m-%d-%H-%M-%S\n')
        new_data: list[str] = list(new_data.splitlines(True))
        new_data.insert(1, timestamp)
        new_data.append('\n')

        # バックアップファイルを読み込み、データを取得する
        try:
            with open(backup_filepath, 'rt', encoding='utf-8') as f:
                old_data = f.readlines()
        except FileNotFoundError:
            old_data = []

        # 新しいデータをバックアップデータに追加する
        old_data = new_data + old_data

        # 最大50個に削る
        old_data.reverse()
        old_data = iter(old_data)
        num_of_separator = 0
        backup_data = []
        while num_of_separator < backup_max:
            data = next(old_data)
            backup_data.append(data)
            if data == '---\n':
                num_of_separator = num_of_separator + 1
        backup_data.reverse()


        # バックアップファイルにデータを書き込む
        try:
            with open(backup_filepath, 'wt', encoding='utf-8') as f:
                f.writelines(backup_data)
        except OSError as e:
            log.error(f'Failed to write backup file: {e}')
            return

        # バックアップ完了メッセージをログに出力する
        log.info(f'Backup completed: {backup_filepath}')

        # backup_frequency秒後に再度バックアップを実行する
        self.master.after(self.backup_frequency, self.backup)

    def autosave(self):
        '''
        定期的にファイルを保存する
        '''
        if not self.do_autosave:
            return
        if not self.filepath:
            return
        if self.file_over_write_save():
            log.info(f'Autosave completed: {self.filepath}')
        else:
            log.error(f'Failed Autosave: {self.filepath}')
        self.master.after(self.autosave_frequency, self.autosave)

    def popup(self, e=None):
        e.widget.mark_set(INSERT, CURRENT)
        self.menu_edit.post(e.x_root,e.y_root)

class SettingWindow(Toplevel):
    '''
    設定ウィンドウに関する設定
    '''
    def __init__(self, e=None, title="SoroEditor - 設定", iconphoto='', size=(1200, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        self.protocol('WM_DELETE_WINDOW', self.close)

        log.info('Open SettingWindow')
        # 設定ファイルの読み込み
        with open('./settings.yaml', mode='rt', encoding='utf-8') as f:
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
        self.newwindow_geometry = self.settings['geometry']
        autosave = self.settings['autosave']
        self.autosave_frequency = self.settings['autosave_frequency']
        backup = self.settings['backup']
        self.backup_frequency = self.settings['backup_frequency']

        self.statusbar_elements_dict_converted = {}
        for i in range(6):
            l = self.statusbar_elements_dict[i]
            self.statusbar_elements_dict_converted[i] = [self.convert_statusbar_elements(val) for val in l]
            while len(self.statusbar_elements_dict_converted[i]) < 4:
                self.statusbar_elements_dict_converted[i].append('')

        self.font_treeview = None

        under = Frame(self)
        under.pack(side=BOTTOM, fill=X, pady=10)
        Button(under, text='キャンセル', command=self.destroy).pack(side=RIGHT, padx=10)
        Button(under, text='OK', command=lambda:self.save(close=True)).pack(side=RIGHT, padx=10)
        Button(under, text='適用', command=lambda:self.save(close=False)).pack(side=RIGHT, padx=10)

        f = Labelframe(self, text='設定')
        f.pack(fill=BOTH, expand=True, padx=5, pady=5)
        nt = Notebook(f, padding=5)
        nt.pack(expand=True, fill=BOTH)

        # 以下、設定項目
        f1 = Frame(nt, padding=5)

        ## 列について
        Label(f1, text='この設定は、新規プロジェクトのデフォルト設定です。'+
            '現在編集中のプロジェクトの列数や列幅を変更するには、\n'+
                'ファイル(F) - プロジェクト設定(F)から行ってください。').pack()
        lf1 = Labelframe(f1, text='デフォルト列', padding=5)
        lf1.pack(fill=X)

        ### 列数
        Label(lf1, text='列数: ').grid(row=0, column=0, columnspan=2, sticky=E)
        lf1.grid_columnconfigure(2, weight=1)
        self.setting_number_of_columns = Spinbox(lf1, from_=1, to=10)
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
        ### 起動サイズ
        Label(lf4, text='サイズ: ').pack(side=LEFT)
        geometry_list = ['Full Screen', '960x540', '1600x900', '1920x1080', '1200x800', '1800x1200', '800x500', '1600x1000']
        self.setting_geometry = StringVar()
        self.setting_geometry_menu = OptionMenu(lf4, self.setting_geometry, self.newwindow_geometry, *geometry_list)
        self.setting_geometry_menu.pack(side=LEFT)
        ### テーマ
        Label(lf4, text='  テーマ: ').pack(side=LEFT)
        theme_names = []
        light_theme_names = [t for t in STANDARD_THEMES.keys() if STANDARD_THEMES[t]['type'] == 'light']
        dark_theme_names = [t for t in STANDARD_THEMES.keys() if STANDARD_THEMES[t]['type'] == 'dark']
        theme_names.extend(light_theme_names)
        theme_names.append('------')
        theme_names.extend(dark_theme_names)
        self.setting_theme = StringVar()
        self.setting_theme_menu = OptionMenu(lf4, self.setting_theme, self.themename, *theme_names)
        self.setting_theme_menu.pack(side=LEFT)

        # タブ - ステータスバー
        f3 = Frame(nt, padding=5)
        for i in range(4):
            f3.grid_columnconfigure(i+1, weight=1)
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
                            '各機能情報', '顔文字',
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

        # タブ - 機能
        f4 = Frame(nt, padding=5)
        ## 自動保存機能
        lf5 = Labelframe(f4, text='自動保存機能', padding=5)
        lf5.pack(fill=X)
        explain_autosave = '''
自動保存機能は、編集中のファイルをバックグラウンドで自動保存する機能です
ファイルを一度も保存していない場合（画面上部のタイトルバーに無題と表示されている場合）
機能しないので注意してください
'''
        Label(lf5, text=explain_autosave).pack(anchor=W)
        self.autosave = BooleanVar()
        self.autosave.set(autosave)
        Checkbutton(lf5, text='*自動保存機能', variable=self.autosave, bootstyle="round-toggle", padding=5).pack(anchor=W)
        ### 自動保存頻度
        Label(lf5, text='\n*自動保存頻度（分）').pack(anchor=W)
        self.setting_autosave_frequency = Spinbox(lf5, from_=1, to=60, wrap=True)
        self.setting_autosave_frequency.insert(0, self.autosave_frequency/60000)
        self.setting_autosave_frequency.pack(anchor=W)

        ## バックアップ機能
        lf6 = Labelframe(f4, padding=5, text='バックアップ機能')
        lf6.pack(fill=X)
        explain_backup = '''
バックアップ機能は、編集中のファイルのデータをバックアップファイルに自動保存する機能です
backup-{ファイル名}.$epに保存されます
ファイルを保存していない場合でも、実行ファイルと同じディレクトリのbackup.$epに保存されます
メモ帳などのテキストエディタでバックアップファイルとプロジェクトファイルを開き、
バックアップファイルにある[---]の間一つををコピーしてプロジェクトファイルに上書きする事でデータを再利用できます
'''
        Label(lf6, text=explain_backup).pack(anchor=W)
        self.backup = BooleanVar()
        self.backup.set(backup)
        Checkbutton(lf6, text='*バックアップ機能', variable=self.backup, bootstyle="round-toggle", padding=5).pack(anchor=W)
        ### バックアップ頻度
        Label(lf6, text='\n*バックアップ頻度（分）').pack(anchor=W)
        self.setting_backup_frequency = Spinbox(lf6, from_=1, to=60, wrap=True)
        self.setting_backup_frequency.insert(0, self.backup_frequency/60000)
        self.setting_backup_frequency.pack(anchor=W)


        nt.add(f2, padding=5, text='            表示            ')
        nt.add(f4, padding=5, text='            機能            ')
        nt.add(f1, padding=5, text='             列             ')
        nt.add(f3, padding=5, text='  *ツールバー・ステータスバー  ')


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

        # 設定値をGUIから取得
        try:
            number_of_columns = int(self.setting_number_of_columns.get())
        except ValueError as e:
            log.exception(f'invalid value in number of columns: {e}')
            number_of_columns = self.settings['columns']['number']
        percentage_of_columns = []
        for i in range(len(self.setting_column_percentage)):
            try:
                percentage_of_columns.append(int(self.setting_column_percentage[i].get()))
            except ValueError as e:
                log.exception(f'invalid value in percentage of columns: {e}')
                percentage_of_columns.append(self.settings['columns']['percentage'][i])
        for _ in range(10-len(percentage_of_columns)):
            percentage_of_columns.append(0)
        font_family = self.font_treeview.selection()[0]
        try:
            font_size = int(self.setting_font_size.get())
        except ValueError as e:
            log.exception(f'invalid value in font_size: {e}')
            font_size = self.settings['font']['size']
        try:
            between_lines = int(self.setting_between_lines.get())
        except ValueError as e:
            log.exception(f'invalid value in between lines: {e}')
            between_lines = self.setitngs['between_lines']
        wrap = self.setting_wrap.get()
        themename = self.setting_theme.get()
        display_line_number = self.setting_display_line_number.get()
        selection_line_highlight = self.setting_selection_line_highlight.get()
        geometry = self.setting_geometry.get()
        autosave = self.autosave.get()
        try:
            autosave_frequency = int(float(self.setting_autosave_frequency.get())*60000)
        except ValueError as e:
            log.exception(f'invalid value in autosave frequency: {e}')
            autosave_frequency = self.settings['autosave_frequency']
        backup = self.backup.get()
        try:
            backup_frequency = int(float(self.setting_backup_frequency.get())*60000)
        except ValueError as e:
            log.exception(f'invalid value in backup frequency: {e}')
            backup_frequency = self.settings['backup_frequency']
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

        # 設定を辞書に格納する
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
        self.settings['geometry'] = geometry
        self.settings['autosave'] = autosave
        self.settings['autosave_frequency'] = autosave_frequency
        self.settings['backup'] = backup
        self.settings['backup_frequency'] = backup_frequency
        self.settings['version'] = __version__

        # 一部設定は再起動を待たずプログラムに反映する
        try:
            app.font.config(family=font_family, size=font_size)
            app.between_lines = between_lines
            app.windowstyle.theme_use(themename)
            app.selection_line_highlight = selection_line_highlight
            if geometry == 'Full Screen':
                app.master.state('zoomed')
            else:
                app.master.state('normal')
                app.master.geometry(geometry)
            app.do_autosave = autosave
            app.autosave_frequency = autosave_frequency
            app.do_backup = backup
            app.backup_frequency = backup_frequency
            for w in app.textboxes:
                w.config(wrap=wrap, spacing2=between_lines)
            app.make_text_editor()
        except tkinter.TclError as e:
            log.error(f'Failed to write to the settigs dict: {e}')

        # 設定ファイルに更新された辞書を保存する
        try:
            with open('./settings.yaml', mode='wt', encoding='utf-8') as f:
                yaml.dump(self.settings, f, allow_unicode=True)
        except (FileNotFoundError, UnicodeDecodeError, yaml.YAMLError) as e:
            error_type = type(e).__name__
            error_message = str(e)
            log.error(f"An error of type {error_type} occurred while saving setting file: {error_message}")
        else:
            log.info('Update setting file')
            MessageDialog('一部の設定（*アスタリスク付きの項目）の反映にはアプリの再起動が必要です\n'+
                        '設定をすぐに反映したい場合、いったんアプリを終了して再起動してください',
                        'SoroEditor - 設定',
                        ['OK']).show(app.md_position)
        if close:
            self.close()

    # フォントファミリー選択ツリービュー
    def setting_font_family_treeview(self, master):
        f1 = Frame(master)
        f1.pack(fill=BOTH, expand=YES)
        self.font_treeview = Treeview(f1, height=10, show=[], columns=[0])
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
            self.font_treeview.insert('', iid=f, index=END, tags=[f], values=[f])
            self.font_treeview.tag_configure(f, font=(f, 10))
        # 初期位置を設定
        self.font_treeview.selection_set(self.font_family)
        self.font_treeview.see(self.font_family)
        # スクロールバーを設定
        vbar = Scrollbar(f1, command=self.font_treeview.yview, orient=VERTICAL, bootstyle='rounded')
        vbar.pack(side=RIGHT, fill=Y)
        self.font_treeview.configure(yscrollcommand=vbar.set)
        return self.font_treeview

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
            ['各機能情報', 'infomation'],
            ['顔文字', 'kaomoji'],
            ['ボタン - ファイルを開く', 'toolbutton_open'],
            ['ボタン - 上書き保存', 'toolbutton_over_write_save'],
            ['ボタン - 名前をつけて保存', 'toolbutton_save_as'],
            ['ステータスバー初期メッセージ', 'statusbar_message']]
        for x in l:
            if val == x[0]: return x[1]
            elif val == x[1]: return x[0]
        return ''

    def close(self):
        log.info('Closed SettingWindow')
        self.destroy()

class ProjectFileSettingWindow(Toplevel):
    def __init__(self, title="SoroEditor - プロジェクト設定", iconphoto='', size=(600, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        self.protocol('WM_DELETE_WINDOW', self.close)

        log.info('Open ProjectFileSettingWindow')
        if app.get_current_data() != app.data:
            self.withdraw()
            mg = MessageDialog('プロジェクトファイル設定を変更する前にプロジェクトファイルを保存します',
                        'SoroEditor - プロジェクトファイル設定',
                        ['OK:success', 'キャンセル:secondary'],)
            mg.show(app.md_position)
            if mg.result == 'OK':
                x = app.file_over_write_save()
                if x:
                    self.deiconify()
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
        number_of_columns = int(self.number_of_columns.get())
        percentage_of_columns = []
        for i in range(len(self.colmun_percentage)):
            percentage_of_columns.append(int(self.colmun_percentage[i].get())/100)
        for _ in range(10-len(percentage_of_columns)):
            percentage_of_columns.append(0)
        app.number_of_columns = number_of_columns
        app.column_percentage = percentage_of_columns
        app.make_text_editor()
        if close:
            self.close()

    def close(self):
        log.info('Close ProjectFileSettingWindow')
        self.destroy()


class ThirdPartyNoticesWindow(Toplevel):
    def __init__(self, title="SoroEditor - ライセンス情報", iconphoto='', size=(1200, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
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
        widget.winfo_children()[0].config(font=font.Font(family=family, size=12), spacing2=10, state=DISABLED)
        widget.pack(fill=BOTH, expand=True)

        # ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)


class HelpWindow(Toplevel):
    def __init__(self, title="SoroEditor - ヘルプ", iconphoto='', size=(600, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        f = Frame(self, padding=5)
        f.pack(fill=BOTH, expand=True)
        t = Text(f, font=app.font)

        font_title = app.font.copy()
        font_title.config(size=15)
        font_h1 = app.font.copy()
        font_h1.config(size=13)
        font_text = app.font.copy()
        font_text.config(size=10)

        def open_url(url):
            def innner(_):
                webbrowser.open_new(url)
            return innner

        t.tag_config('title', font=font_title, justify=CENTER,
                    spacing1=10, spacing3=10)
        t.tag_config('h1', font=font_h1, underline=True,
                    spacing1=20, spacing3=15)
        t.tag_config('text', font=font_text)
        t.tag_config('github', underline=True)
        t.tag_config('github_issue', underline=True)
        t.tag_config('mail', underline=True)
        github = open_url('https://github.com/joppincal/SoroEditor')
        github_issue = open_url('https://github.com/joppincal/SoroEditor/issue')
        mail = open_url('mailto://Joppincal@mailo.com?subject=SoroEditorについて')
        t.tag_bind('github', '<Button-1>', github)
        t.tag_bind('github_issue', '<Button-1>', github_issue)
        t.tag_bind('mail', '<Button-1>', mail)

        t.insert(END, 'そろエディタ\n', 'title')
        t.insert(END, '概要\n', 'h1')
        t.insert(END,
'''そろエディタは「並列テキストエディタ」です
複数列が揃って並び、それぞれテキストエディタとして編集できます
スクロールを同期できます
同様の内容は、
ヘルプ(H)→ヘルプを表示(H) F1や
ヘルプ(H)→初回起動メッセージを表示(F)
もしくは\n''', 'text')
        t.insert(END, 'https://github.com/joppincal/SoroEditor\n', ('text', 'github'))
        t.insert(END, 'で確認できます\n', 'text')
        t.insert(END, '設定可能項目\n', 'h1')
        t.insert(END,
'''設定(S)	から
・列数・列幅
・フォント（ファミリー・サイズ）
・表示テーマ
・画面端の折り返し（なし 文字で折り返し 単語で折り返し（英語向け））
・ステータスバー（ウィンドウ下部の、様々な情報が記述されている場所）
・ツールバー（ウィンドウ上部のボタンが配置された場所）
・編集中の行の強調表示（デフォルトでは下線）
・行番号
・起動時のウィンドウサイズ

*列数・列幅に関しては、ファイル(F)→プロジェクト設定(F)	から
プロジェクトファイルごとに設定できます\n''', 'text')
        t.insert(END, 'ホットキー\n', 'h1')
        t.insert(END,
'''Ctrl+O:           ファイルを開く
Ctrl+R:           前回使用したファイルを開く
Ctrl+S:           上書き保存
Ctrl+Shift+S:     名前をつけて保存
Ctrl+Z:           取り消し
Ctrl+Shift+Z:     取り消しを戻す
Ctrl+Enter:       1行追加（下）
Ctrl+Shift+Enter: 1行追加（上）
Ctrl+L:           1行選択
Ctrl+Q, Alt+<:    左の列に移動
Ctrl+W, Alt+>:    右の列に移動\n''', 'text')
        t.insert(END, 'プロジェクトファイル\n', 'h1')
        t.insert(END,
'''プロジェクトファイルはYAML形式のファイルの拡張子を*.sepに変更したものです
一般のテキストエディタでも閲覧・編集が可能です\n''', 'text')
        t.insert(END, '設定ファイル\n', 'h1')
        t.insert(END,
'''設定ファイルはYAML形式です（setting.yaml）
一般のテキストエディタでも閲覧・編集が可能です
実行ファイル（SoroEditor.exe または SoroEditor.py）と
同じフォルダに置いてください\n''', 'text')
        t.insert(END, '連絡先\n', 'h1')
        t.insert(END, '不具合報告・要望などは、', 'text')
        t.insert(END, 'Joppincal@mailo.com\n', ('text', 'mail'))
        t.insert(END, 'もしくは', 'text')
        t.insert(END, 'Githubのissue', ('text', 'github_issue'))
        t.insert(END, 'にて')

        t.config(state=DISABLED)
        t.pack(fill=BOTH, expand=True)

        # ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)


class AboutWindow(Toplevel):
    def __init__(self, title="SoroEditorについて", iconphoto='', size=(600, 500), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        frame = Frame(self, padding=5)
        frame.pack(fill=BOTH, expand=True)
        main = Text(frame)

        main.tag_config('title', justify=CENTER, spacing2=10, font=font.Font(size=15, weight='bold'))
        main.tag_config('text', justify=CENTER, spacing2=10, font=font.Font(size=12, weight='normal'))
        main.tag_config('link', justify=CENTER, spacing2=10, font=font.Font(size=12, weight='normal', underline=True))
        github = 'https://github.com/joppincal/SoroEditor'
        main.tag_config('github')
        main.tag_bind('github', '<Button-1>', lambda e: webbrowser.open_new(github))
        homepage = ''
        main.tag_config('homepage')
        main.tag_bind('homepage', '<Button-1>', lambda e: webbrowser.open_new(homepage))

        main.insert(END, 'ここにアイコンを表示\n')
        main.insert(END, 'そろエディタ\nSoroEditor\n', 'title')
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


class ExportWindow(Toplevel):

    def __init__(self, title="SoroEditor - エクスポート", iconphoto='', size=(1000, 700), position=None, minsize=None, maxsize=None, resizable=(0, 0), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        log.info('Open ExportWindow')

        self.protocol('WM_DELETE_WINDOW', self.close)

        self.current_data = app.get_current_data()

        f = Frame(self, padding=10)
        f.pack(fill=BOTH, expand=True)

        self.f1 = Frame(f, padding=5)
        Label(self.f1, text='ファイル形式').pack(side=LEFT, padx=2)
        self.file_format = StringVar(value='CSV')
        file_formats = ['CSV', 'TSV', 'テキスト', 'テキスト(Shift-JIS)']
        OptionMenu(self.f1, self.file_format, self.file_format.get(), *file_formats, command=self.make_setting_frame).pack(side=LEFT, padx=2)
        self.shift_jis_alert = Label(self.f1)
        self.shift_jis_alert.pack(side=LEFT, padx=10)
        self.f1.pack(fill=X)

        f2 = Frame(f, padding=(5, 5))
        Button(f2, text='キャンセル', command=self.destroy, bootstyle=SECONDARY).pack(side=RIGHT, padx=2)
        Button(f2, text='出力', command=self.export).pack(side=RIGHT, padx=2)
        Separator(f2).pack(fill=X)
        f2.pack(side=BOTTOM, fill=X, anchor=S)

        self.lf1 = Labelframe(f, text='設定', padding=5)
        self.lf1.pack(fill=BOTH, expand=True)
        self.lf2 = Labelframe(self.lf1, text='出力する列', padding=5)
        self.lf2.place(relx=0.0, rely=0.0, relheight=1.0, relwidth=0.3)
        self.lf3 = Labelframe(self.lf1, text='詳細', padding=5)
        self.lf3.place(relx=0.3, rely=0.0, relheight=1.0, relwidth=0.7)
        self.make_setting_frame(self.file_format.get())

        # ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)

    def make_setting_frame(self, file_format:str):
        # リセット
        for w in self.lf2.winfo_children():
            w.destroy()
        for w in self.lf3.winfo_children():
            w.destroy()

        # 出力する列
        ## 列のタイトルを取得、空欄の場合デフォルト設定
        line_names = [self.current_data[i]['title']
                        if self.current_data[i]['title']
                        else f'列{i+1}'
                        for i in range(app.number_of_columns)]
        ## CSV/TSV/PDFの場合の設定画面
        if file_format in ['CSV', 'TSV', 'PDF']:
            self.line_checkbuttons_variables = [BooleanVar(value=True)
                                            for _ in range(len(line_names))]
            line_checkbuttons = [Checkbutton(
                                    self.lf2, text=line_names[i], padding=10,
                                    variable=self.line_checkbuttons_variables[i],
                                    bootstyle='round-toggle')
                                    for i in range(len(line_names))]
            for w in line_checkbuttons:
                w.pack(anchor=W)
        ## テキスト/テキスト(Shift-JIS)の場合の設定画面
        elif file_format in ['テキスト', 'テキスト(Shift-JIS)']:
            self.line_radiobuttons_variable = IntVar(value=0)
            line_radiobuttons = [Radiobutton(
                                    self.lf2, text=line_names[i], padding=10,
                                    variable=self.line_radiobuttons_variable,
                                    value=i)
                                    for i in range(len(line_names))]
            for w in line_radiobuttons:
                w.pack(anchor=W)

        # 詳細
        # テキスト(Shift-JIS)の場合、データの一部が再現できない場合がある事を記述する

        if self.file_format.get() == 'テキスト(Shift-JIS)':
            self.shift_jis_alert.config(text='*Shift-JISで表現できない一部のUnicode文字は「?」に置き換えられます')
        else:
            self.shift_jis_alert.config(text='')

        def filepath_label_change(widget:Label, newfilepath=''):
            widget.config(text=newfilepath)

        def get_new_filepath(defaultextension, filetype, widget):

            def inner(defaultextension=defaultextension, filetype=filetype, widget=widget):
                if app.filepath:
                    initialdir = os.path.dirname(app.filepath)
                    initialfile = os.path.splitext(os.path.basename(app.filepath))[0] + filetype
                else:
                    initialdir = sys.argv[0]
                    initialfile = 'export' + filetype
                if filetype == '.csv':
                    filetype = (['CSVファイル', '*.csv'],
                                ['すべてのファイル', '*.*'])
                elif filetype == '.tsv':
                    filetype = (['TSVファイル', '*.tsv'],
                                ['すべてのファイル', '*.*'])
                elif filetype == '.pdf':
                    filetype = (['PDFファイル', '*.pdf'],
                                ['すべてのファイル', '*.*'])
                elif filetype == '.txt':
                    filetype = (['テキストファイル', '*.txt'],
                                ['すべてのファイル', '*.*'])
                newfilepath = filedialog.asksaveasfilename(
                                            defaultextension=defaultextension,
                                            filetypes=filetype,
                                            initialdir=initialdir,
                                            initialfile=initialfile)
                if newfilepath:
                    # ドライブ文字を大文字にする
                    drive, tail = os.path.splitdrive(newfilepath)
                    drive = drive.upper()
                    newfilepath = os.path.join(drive, tail)
                    filepath_label_change(widget, newfilepath)
            return inner

        self.include_title_in_output = BooleanVar(value=True)
        Checkbutton(self.lf3, padding=10,
                    text='タイトルを含めて出力する',
                    variable=self.include_title_in_output,
                    bootstyle='round-toggle').grid(column=0, columnspan=2, row=0, sticky=W, padx=10, pady=10)

        # 保存先の設定
        if file_format == 'CSV':
            extension = '.csv'
        elif file_format == 'TSV':
            extension = '.tsv'
        elif file_format == 'PDF':
            extension = '.pdf'
        elif file_format in ['テキスト', 'テキスト(Shift-JIS)']:
            extension = '.txt'

        if app.filepath:
            self.filepath = StringVar(value=os.path.dirname(app.filepath) + '/' + os.path.splitext(os.path.basename(app.filepath))[0] + extension)
        else:
            self.filepath = StringVar(value=os.path.join(os.path.dirname(sys.argv[0]), 'export'+extension))

        Label(self.lf3, text='保存先', padding=10).grid(column=0, row=2, sticky=W, padx=10, pady=5)
        filepath_label = Label(self.lf3)
        filepath_label.grid(column=1, columnspan=3, row=3, sticky=W, padx=10, pady=5)
        filepath_label_change(filepath_label, self.filepath.get())
        Button(self.lf3, text='変更', command=get_new_filepath(extension, extension, filepath_label)).grid(column=1, row=2, sticky=W, padx=10, pady=5)

    def export(self):
        # データを現在の表示状態に更新する
        self.data = [(self.current_data[i]['title'], self.current_data[i]['text'])
                        for i in range(app.number_of_columns)]

        # ファイルフォーマットがCSVまたはTSVの場合はCSV形式で出力する
        if self.file_format.get() in ['CSV', 'TSV']:
            data = self.export_csv()

        # ファイルフォーマットがテキストまたはテキスト(Shift-JIS)の場合はテキスト形式で出力する
        if self.file_format.get() in ['テキスト', 'テキスト(Shift-JIS)']:
            data = self.export_text()

        # 出力データをファイルに書き込む
        if self.write_file(data):
            messagedialog = MessageDialog(
                f'「{self.filepath.get()}」を作成しました',
                'SoroEditor - エクスポート',
                ['OK:primary'])
        else:
            messagedialog = MessageDialog(
                '出力に失敗しました。同名ファイルを他のソフトで開いていないか確認してください',
                'SoroEditor - エクスポート',
                ['OK:primary']
            )
        messagedialog.show()

    def export_csv(self) -> str:
        # 出力するデータを取得する
        data = self.data
        # データの列数を取得する
        data_length = len(data)

        output_data = ''

        # 出力するデータをチェックボックスで選択された行のみに絞り込む
        data = [data[i] for i in range(data_length)
                if self.line_checkbuttons_variables[i].get()]
        # 絞り込んだデータの列数を取得する
        data_length = len(data)

        # ファイルフォーマットに応じて区切り文字を設定する
        separator = {'CSV': ',', 'TSV': '\t'}.get(self.file_format.get())

        # ヘッダー行を追加する
        if self.include_title_in_output.get():
            # ヘッダー行のデータを取得する
            header = [data[i][0] for i in range(data_length)]
            # ヘッダー行を出力データに追加する
            output_data += f'{separator}'.join(header) + '\n'

        # データ行を追加する
        # 最大の行数を計算する
        max_row_length = max([len(text.split('\n')) for title, text in data])

        for i in range(max_row_length):
            # 1つの行のデータを表すリストを初期化
            row_data = []
            for title, text in data:
                # 1つのセルのデータを表す文字列を取得する
                cell_data = ''
                text_rows = text.split('\n')
                if i < len(text_rows):
                    cell_data = text_rows[i]
                row_data.append(cell_data)
                # 1つの行のデータを出力データに追加する
            output_data += f'{separator}'.join(row_data) + '\n'

        # 出力するデータを返す
        return output_data

    def export_text(self) -> str:
        data = self.data

        output_data = ''

        # 設定が有効の場合、タイトル行を出力
        if self.include_title_in_output.get():
            # ラジオボタンで選択された行のタイトルを取得し、改行を追加して出力する
            output_data += data[self.line_radiobuttons_variable.get()][0] + '\n'

        # ラジオボタンで選択された行のテキストを取得して出力する
        output_data += data[self.line_radiobuttons_variable.get()][1] + '\n'

        # 出力するデータを返す
        return output_data

    def write_file(self, data):
        # すでに同名ファイルが存在するか確認し、存在する場合上書きするか確認
        if os.path.exists(self.filepath.get()):
            messagedialog = MessageDialog(
                f'ファイル「{self.filepath.get()}」はすでに存在します',
                'SoroEditor - エクスポート',
                ['上書き:primary', 'キャンセル'])
            messagedialog.show()
            if messagedialog.result == 'キャンセル':
                return False
        # エンコード方法を設定する
        if self.file_format.get() == 'テキスト(Shift-JIS)':
            encodeing = 'cp932'
        elif self.file_format.get() in ['CSV', 'TSV']:
            encodeing = 'utf-8-sig'
        else:
            encodeing = 'utf-8'
        # ファイルを作成し、書き込む
        try:
            with open(self.filepath.get(), mode='wt', encoding=encodeing, errors='replace') as f:
                f.write(data)
        except PermissionError as e:
            log.error(f'Failed export {self.file_format.get()}: {self.filepath.get()} {type(e).__name__}: {e}')
            return False
        except Exception as e:
            log.error(f'Failed export {self.file_format.get()}: {self.filepath.get()} {type(e).__name__}: {e}')
            return False
        else:
            log.info(f'Success export {self.file_format.get()}: {self.filepath.get()}')
            # 書き込みを行い成功した場合Trueを返す
            return True

    def close(self):
        log.info('Close ExportWindow')
        self.destroy()


if __name__ == '__main__':
    log_setting()
    log.info('===Start Application===')
    root = Window(title='SoroEditor', minsize=(800, 500))
    app = Main(master=root)
    app.mainloop()
    log.info('===Close Application===')