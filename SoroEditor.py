'''
SoroEditor - Joppincal
This software is distributed under the MIT License. See LICENSE for details.
See ThirdPartyNotices.txt for third party libraries.
'''
import datetime
import difflib
import os
import re
import sys
import webbrowser
from collections import deque, namedtuple
from random import choice
import csv
import logging
import logging.handlers
import yaml

from PIL import Image, ImageTk
from tkinter import BooleanVar, IntVar, StringVar, PhotoImage, TclError, filedialog, font
from ttkbootstrap import (Button, Checkbutton, Entry, Frame, Label, Labelframe,
Menu, Notebook, OptionMenu, PanedWindow, Radiobutton, Scrollbar, Separator, Spinbox,
Style, Text, Toplevel, Treeview, Window)
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs.dialogs import MessageDialog
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.themes.standard import *


__version__ = '0.4.1'
__projversion__ = '0.3.8'
with open(os.path.join(os.path.dirname(__file__), 'ThirdPartyNotices.txt'), 'rt', encoding='utf-8') as f:
    __thirdpartynotices__ = f.read()

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
        # デフォルト設定
        self.settings ={'autosave': False,
                        'autosave_frequency': 60000,
                        'backup': True,
                        'backup_frequency': 300000,
                        'between_lines': 10,
                        'button_style': 'icon_with_text',
                        'columns': {'number': 3, 'percentage': [15, 55, 30]},
                        'display_line_number': True,
                        'font': {'family': 'nomal', 'size': 12},
                        'geometry': '1600x1000',
                        'ms_align_the_lines': 50,
                        'recently_files': [],
                        'selection_line_highlight': True,
                        'search_engines': {
                            'Google': 'https://www.google.com/search?q=%s',
                            'Bing': 'https://www.bing.com/search?q=%s',
                            'Yahoo!Japan': 'https://search.yahoo.co.jp/search?p=%s',
                            'Wikipedia ja': 'http://ja.wikipedia.org/wiki/%s'
                            },
                        'statusbar_element_settings': {
                            0: ['hotkeys3', 'statusbar_message'],
                            1: ['hotkeys2'],
                            2: ['hotkeys1'],
                            3: ['current_place'],
                            4: ['toolbutton_create', 'toolbutton_open', 'toolbutton_save', 'toolbutton_file_reload'],
                            5: ['toolbutton_bookmark', 'toolbutton_template', 'toolbutton_search', 'toolbutton_replace', 'toolbutton_setting']
                            },
                        'templates':{},
                        'themename': '',
                        'version': __version__,
                        'wrap': NONE}
        # 設定ファイルを開く
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
            log.info('Succeed loading settings')

        # 設定ファイルから各設定を読み込み
        ## ウィンドウサイズ
        if self.settings['geometry'] == 'Full Screen':
            self.master.state('zoomed')
        else:
            try:
                self.master.geometry(self.settings['geometry'])
            except TclError as e:
                log.error(f'Failed to set window size: {e}')
        ## テーマを設定
        self.windowstyle = Style()
        self.windowstyle.theme_use(self.settings['themename'])
        ## ボタンの外観設定
        self.button_style: str = self.settings['button_style']
        if not self.button_style in ('icon_with_text', 'icon_only', 'text_only'):
            self.button_style = 'icon_with_text'
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
        ## 定型文
        self.templates = {i: self.settings['templates'].get(i, '') for i in range(10)}
        ## 検索エンジン
        self.search_engines: dict = self.settings['search_engines']

        # サブウィンドウを入れる変数
        self.setting_window = None
        self.search_window = None
        self.template_window = None
        self.bookmark_window = None
        self.windows = {
            'setting': self.setting_window,
            'search': self.search_window,
            'template': self.template_window,
            'bookmark': self.bookmark_window,
        }

        self.Icons = Icons()
        self.filepath = ''
        self.data0 = {}
        for i in range(10):
            self.data0[i] = {'text': '', 'title': ''}
        self.data0['columns'] = {'number': self.number_of_columns, 'percentage': [int(x*100) for x in self.column_percentage]}
        self.data0['version'] = __projversion__
        self.data0['bookmarks'] = []
        self.data0['use_regex_in_bookmarks'] = False
        self.data = self.data0.copy()
        self.bookmarks = []
        self.use_regex_in_bookmarks = False
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
        self.menu_file = Menu(menubar)

        self.menu_file.add_command(label='新規作成(N)', command=self.file_create, accelerator='Ctrl+N', underline=5)
        self.menu_file.add_command(label='ファイルを開く(O)', command=self.file_open, accelerator='Ctrl+O', underline=8)
        self.menu_file.add_command(label='上書き保存(S)', command=self.file_over_write_save, accelerator='Ctrl+S', underline=6)
        self.menu_file.add_command(label='名前をつけて保存(A)', command=self.file_save_as, accelerator='Ctrl+Shift+S', underline=9)
        self.menu_file.add_command(label='エクスポート(E)', command=ExportWindow, accelerator='Ctrl+Shift+E', underline=7)
        self.menu_file.add_command(label='インポート(I)', command=ImportWindow, accelerator='Ctrl+Shift+I', underline=6)
        self.menu_file.add_command(label='プロジェクト設定(F)', command=ProjectFileSettingWindow, underline=9)
        self.menu_file.add_command(label='再読込(W)', command=self.make_text_editor, accelerator='F5', underline=4)

        ###メニューバー - ファイル - 最近使用したファイル
        self.menu_file_recently = Menu(self.menu_file)

        try:
            self.menu_file_recently.add_command(label=self.recently_files[0], command=lambda:self.file_open(file_path_to_open=self.recently_files[0]), accelerator='Ctrl+R')
            self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file_path_to_open=self.recently_files[1]))
            self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file_path_to_open=self.recently_files[2]))
            self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file_path_to_open=self.recently_files[3]))
            self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file_path_to_open=self.recently_files[4]))
        except IndexError:
            pass

        self.menu_file.add_cascade(label='最近使用したファイル(R)', menu=self.menu_file_recently, underline=11)

        self.menu_file.add_separator()
        self.menu_file.add_command(label='終了(Q)', command=lambda:self.file_close(exit_on_completion=True), accelerator='Alt+F4', underline=3)
        menubar.add_cascade(label='ファイル(F)', menu=self.menu_file, underline=5)

        ## メニューバー - 編集
        self.menu_edit = Menu(menubar)
        self.make_menu_edit(self.menu_edit)
        menubar.add_cascade(label='編集(E)', menu=self.menu_edit, underline=3)

        ## メニューバー - 検索
        self.menu_search = Menu(menubar)

        self.menu_search.add_command(label='検索(S)', command=self.open_SearchWindow, accelerator='Ctrl+F', underline=3)
        self.menu_search.add_command(label='置換(R)', command=lambda: self.open_SearchWindow('1'), accelerator='Ctrl+Shift+F', underline=3)
        for name, url in self.search_engines.items():
            self.menu_search.add_command(label=name+'で検索', command=self.search_on_web(url), underline=0)


        menubar.add_cascade(label='検索(S)', menu=self.menu_search, underline=3)

        ## メニューバー - 定型文
        self.menu_templates = Menu(menubar)
        self.menu_templates.add_command(label='定型文(T)', command=self.open_TemplateWindow, accelerator='Ctrl+T', underline=4)
        self.menu_templates.add_separator()
        self.make_menu_templates()
        menubar.add_cascade(label='定型文(T)', menu=self.menu_templates, underline=4)

        ## メニューバー - 付箋
        self.menu_bookmark = Menu(menubar)
        self.menu_bookmark.add_command(label='付箋(B)', command=self.open_BookmarkWindow, accelerator='Ctrl+B', underline=3)
        self.menu_bookmark.add_separator()
        self.make_menu_bookmarks()
        menubar.add_cascade(label='付箋(B)', menu=self.menu_bookmark, underline=3)

        ## メニューバー - 設定
        self.menu_option = Menu(menubar)
        self.menu_option.add_command(label='設定(O)', command=self.open_SettingWindow, accelerator='Ctrl+Shift+P', underline=3)
        self.menu_option.add_command(label='プロジェクト設定(F)', command=ProjectFileSettingWindow, underline=9)
        menubar.add_cascade(label='設定(O)', menu=self.menu_option, underline=3)

        ## メニューバー - ヘルプ
        self.menu_help = Menu(menubar)
        self.menu_help.add_command(label='ヘルプを表示(H)', command=HelpWindow, accelerator='F1', underline=7)
        self.menu_help.add_command(label='初回起動メッセージを表示(F)', command=lambda: self.file_open(file_path_to_open=os.path.join(os.path.dirname(__file__), 'hello.txt')), underline=13)
        self.menu_help.add_command(label='SoroEditorについて(A)', command=AboutWindow, underline=16)
        self.menu_help.add_command(label='ライセンス情報(L)', command=ThirdPartyNoticesWindow, underline=8)
        menubar.add_cascade(label='ヘルプ(H)', menu=self.menu_help, underline=4)

        # メニューバーの設置
        self.master.config(menu = menubar)

        # メニューバー関連のキーバインドを設定
        self.open_last_file_id = self.master.bind('<Control-r>', self.open_last_file)

        # 各パーツを製作
        self.f1 = Frame(self.master, padding=5)
        self.f2 = Frame(self.master)
        self.vbar = Scrollbar(self.f2, command=self.vbar_command, style=ROUND, takefocus=False)

        self.columns:list[Frame] = []
        self.entrys:list[Entry] = []
        self.maintexts:list[Text] = []

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
            widget = self.master.focus_lastfor()
            try:
                if widget.winfo_class() in ('Text', 'TEntry'):
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
        self.hotkeys2 = ('label', '[Ctrl+Enter]: 1行追加(下)  [Ctrl+Shift+Enter]: 1行追加(上)  [Ctrl+F]: 検索  [Ctrl+Shift+F]: 置換  [Ctrl+Z]: 取り消し  [Ctrl+Shift+Z]: 取り消しを戻す')
        self.hotkeys3 = ('label', '[Ctrl+Q][Alt+Q][Alt+,]: 左に移る  [Ctrl+W][Alt+W][Alt+.]: 右に移る')
        ## 各機能情報
        self.infomation = ('label', f'自動保存: {self.do_autosave}, バックアップ: {self.do_backup}')
        ## 顔文字
        kaomoji_list = ['( ﾟ∀ ﾟ)', 'ヽ(*^^*)ノ ', '(((o(*ﾟ▽ﾟ*)o)))', '(^^)', '(*^○^*)', '(`o´)', '(´・ω・`)', 'ヽ(`Д´)ﾉ ', '( *´・ω)/(；д； )', '( ；∀；)', '(⊃︎´▿︎` )⊃︎ ', '(・∀・)', '((o(^∇^)o))', 'ｷﾀ━━━━(ﾟ∀ﾟ)━━━━!!', ' 【審議中】　(　´・ω) (´・ω・) (・ω・｀) (ω・｀ ) ', '(*´ω｀*)', '(●▲●)', '(○▽○)', '(´・ω・)つ旦', ' (●ﾟ◇ﾟ●)', "（'ω`）"]
        self.kaomoji = ('label', choice(kaomoji_list))
        ## その他ステータスバー
        self.now = StringVar()
        self.clock_mode = self.settings.get('clock_mode', 'ymdhm')
        self.clock = ('label', None, self.clock_change, None, self.now)
        ## ツールボタン
        self.toolbutton_create = ('button', '新規作成', self.file_create, self.Icons.file_create)
        self.toolbutton_open = ('button', 'ファイルを開く', self.file_open, self.Icons.file_open)
        self.toolbutton_save = ('button', '上書き保存', self.file_over_write_save, self.Icons.file_save)
        self.toolbutton_save_as = ('button', '名前をつけて保存', self.file_save_as, self.Icons.file_save_as)
        self.toolbutton_file_reload = ('button', '再読込', self.file_reload, self.Icons.refresh)
        self.toolbutton_project_setting = ('button', 'プロジェクト設定', ProjectFileSettingWindow, self.Icons.project_settings)
        self.toolbutton_setting = ('button', '設定', self.open_SettingWindow, self.Icons.settings)
        self.toolbutton_search = ('button', '検索', self.open_SearchWindow, self.Icons.search)
        self.toolbutton_replace = ('button', '置換', lambda: self.open_SearchWindow('1'), self.Icons.replace)
        self.toolbutton_import = ('button', 'インポート', ImportWindow, self.Icons.import_)
        self.toolbutton_export = ('button', 'エクスポート', ExportWindow, self.Icons.export)
        self.toolbutton_template = ('button', '定型文', self.open_TemplateWindow, self.Icons.template)
        self.toolbutton_bookmark = ('button', '付箋', self.open_BookmarkWindow, self.Icons.bookmark)
        self.toolbutton_undo = ('button', '取り消し', self.undo, self.Icons.undo)
        self.toolbutton_repeat = ('button', '取り消しを戻す', self.repeat, self.Icons.repeat)
        ## 初期ステータスバーメッセージ
        self.statusbar_message = ('label', 'ツールバー・ステータスバーの項目は設定-ツールバー・ステータスバー から変更できます')

        # 設定ファイルからステータスバーの設定を読み込むメソッド
        def statusbar_element_setting_load(name: str=str):
            if re.compile(r'letter_count_\d+').search(name):
                return(letter_count(int(re.search(r'\d+', name).group()) - 1))
            if re.compile(r'line_count_\d+').search(name):
                return(line_count(int(re.search(r'\d+', name).group()) - 1))
            if re.compile(r'line_count_debug_\d+').search(name):
                return(line_count_debug(int(re.search(r'\d+', name).group()) - 1))
            pairs = {
                'current_place': current_place(),
                'hotkeys1': self.hotkeys1,
                'hotkeys2': self.hotkeys2,
                'hotkeys3': self.hotkeys3,
                'infomation': self.infomation,
                'kaomoji': self.kaomoji,
                'now':self.clock,
                'toolbutton_create': self.toolbutton_create,
                'toolbutton_open': self.toolbutton_open,
                'toolbutton_save': self.toolbutton_save,
                'toolbutton_save_as': self.toolbutton_save_as,
                'toolbutton_file_reload': self.toolbutton_file_reload,
                'toolbutton_project_setting': self.toolbutton_project_setting,
                'toolbutton_setting': self.toolbutton_setting,
                'toolbutton_search': self.toolbutton_search,
                'toolbutton_replace': self.toolbutton_replace,
                'toolbutton_import': self.toolbutton_import,
                'toolbutton_export': self.toolbutton_export,
                'toolbutton_template': self.toolbutton_template,
                'toolbutton_bookmark': self.toolbutton_bookmark,
                'toolbutton_undo': self.toolbutton_undo,
                'toolbutton_repeat': self.toolbutton_repeat,
                'statusbar_message': self.statusbar_message,
            }
            method = pairs.get(name)
            if method:
                return method

        # ステータスバー作成メソッド
        def make_statusbar_element(elementtype=None, text=None, commmand=None, image=None, textvariable=None):
            if elementtype == 'label':
                label = Label(text=text, image=image, textvariable=textvariable)
                label.bind('<Button>', commmand)
                return label
            if elementtype == 'button':
                if self.button_style == 'icon_only':
                    compound = None
                elif self.button_style == 'text_only':
                    compound = None
                    image = None
                else: # icon_with_textの場合及びそのほか不適切な文字列の場合
                    compound = LEFT
                button = Button(text=text, command=commmand, image=image, takefocus=False, compound=compound, textvariable=textvariable)
                if self.windowstyle.theme_use() in  [t for t in STANDARD_THEMES.keys() if STANDARD_THEMES[t]['type'] == 'light']:
                    button.config(bootstyle='light')
                elif self.windowstyle.theme_use() in [t for t in STANDARD_THEMES.keys() if STANDARD_THEMES[t]['type'] == 'dark']:
                    button.config(bootstyle='dark')
                return button

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
                for name in l:
                    l2.append(statusbar_element_setting_load(name))
            except TypeError as e:
                pass

            # ステータスバーを作る
            if num in range(7) and l2:
                self.statusbar_element_dict[num] = dict()
                for i, e in enumerate(l2):
                    if not e: continue
                    self.statusbar_element_dict[num][i] = make_statusbar_element(*e)
                for e in self.statusbar_element_dict[num].values():
                    if num < 4:
                        i = 1
                    else:
                        i = 0
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
        self.statusbar_element_reload = statusbar_element_reload

        self.statusbar_element_dict = dict()

        # ステータスバーを生成
        self.statusbars = list()
        for i in range(4):
            self.statusbars.append(PanedWindow(self.master, height=30, orient=HORIZONTAL))
        for i in range(3):
            self.statusbars.append(PanedWindow(self.master, height=40, orient=HORIZONTAL))
        for i in range(7):
            statusbar_element_setting(num=i)

        # 各パーツを設置
        for w in self.statusbars[0:4]:
            if len(w.panes()):
                w.pack(fill=X, side=BOTTOM, padx=5, pady=3)
        for w in self.statusbars[4:]:
            if len(w.panes()):
                w.pack(fill=X, pady=1)

        self.f2.pack(fill=Y, side=RIGHT, pady=18)
        self.vbar.pack(fill=Y, expand=YES)
        self.f1.pack(fill=BOTH, expand=YES)

        # キーバインドを設定
        def test_focus_get(e):
            widget = self.master.focus_get()
            print(widget, 'has focus')
        def focus_to_right(e):
            if type(e.widget) == Text:
                next_text_index = self.maintexts.index(e.widget)+1
                if next_text_index >= len(self.maintexts): next_text_index = 0
                next_text = self.maintexts[next_text_index]
                insert = e.widget.index(INSERT)
                next_text.mark_set(INSERT, insert)
                next_text.focus_set()
        def focus_to_left(e):
            if type(e.widget) == Text:
                prev_text_index = self.maintexts.index(e.widget)-1
                prev_text = self.maintexts[prev_text_index]
                insert = e.widget.index(INSERT)
                prev_text.mark_set(INSERT, insert)
                prev_text.focus_set()
        def focus_to_bottom(e):
            if e.widget.winfo_class() == 'TEntry':
                index = self.entrys.index(e.widget)
                self.maintexts[index].focus_set()
                self.maintexts[index].mark_set(INSERT, 1.0)
                self.maintexts[index].see(1.0)
        def reload(e):
            statusbar_element_reload()
            if self.selection_line_highlight: self.highlight()
            self.set_text_widget_editable()
        self.master.bind('<KeyPress>', reload)
        self.master.bind('<KeyRelease>', reload)
        self.master.bind('<Button>', reload)
        self.master.bind('<ButtonRelease>', reload)
        self.master.bind('<KeyRelease>', self.recode_edit_history, '+')
        self.master.bind('<Control-x>', lambda _: self.cut_copy(0))
        self.master.bind('<Control-c>', lambda _: self.cut_copy(1))
        self.master.bind('<Control-z>', self.undo)
        self.master.bind('<Control-Z>', self.repeat)
        self.master.bind('<Control-n>', self.file_create)
        self.master.bind('<Control-o>', self.file_open)
        self.master.bind('<Control-s>', self.file_over_write_save)
        self.master.bind('<Control-S>', self.file_save_as)
        self.master.bind('<F5>', self.file_reload)
        self.master.bind('<Control-P>', self.open_SettingWindow)
        self.master.bind('<Control-w>', focus_to_right)
        self.master.bind('<Alt-w>', focus_to_right)
        self.master.bind('<Alt-.>', focus_to_right)
        self.master.bind('<Alt-Right>', focus_to_right)
        self.master.bind('<Control-q>', focus_to_left)
        self.master.bind('<Alt-q>', focus_to_left)
        self.master.bind('<Alt-,>', focus_to_left)
        self.master.bind('<Alt-Left>', focus_to_left)
        self.master.bind('<Control-l>', self.select_line)
        self.master.bind('<Control-Return>', lambda e: self.newline(e, 1))
        self.master.bind('<Control-Shift-Return>', lambda e: self.newline(e, 0))
        self.master.bind('<Down>', focus_to_bottom)
        self.master.bind('<Return>', focus_to_bottom)
        self.master.bind('<Control-^>', test_focus_get)
        self.master.bind('<Control-Y>', self.print_history)
        self.master.bind('<F1>', lambda _: HelpWindow())
        self.master.bind('<Control-E>', lambda _: ExportWindow())
        self.master.bind('<Control-I>', lambda _: ImportWindow())
        self.master.bind('<Control-f>', self.open_SearchWindow)
        self.master.bind('<Control-F>', lambda _: self.open_SearchWindow('1'))
        self.master.bind('<Control-t>', self.open_TemplateWindow)
        self.master.bind('<Control-b>', self.open_BookmarkWindow)
        for i in range(10):
            num = 0 if i == 9 else i + 1
            self.master.bind(f'<Alt-Key-{num}>', self.menu_templates_clicked(i))
            self.master.bind(f'<Control-Key-{num}>', self.menu_templates_clicked(i))

        # 設定ファイルに異常があり初期化した場合、その旨を通知する
        if self.settingFile_Error_md:
            self.settingFile_Error_md.show(self.md_position)

        # 設定ファイルが存在しなかった場合、初回起動と扱う
        if initialization:
            self.file_open(file_path_to_open=os.path.join(os.path.dirname(__file__), 'hello.txt'))

        # ファイルを渡されているとき、そのファイルを開く
        if len(sys.argv) > 1:
            self.file_open(file_path_to_open=sys.argv[1])

        self.master.after(500, self.change_window_title)
        self.now_time_set()
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

    def make_text_editor(self, e=None):
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
                                        spacing3=self.between_lines,)
            self.line_number_box.tag_config('right', justify=RIGHT)
            for i in range(9999):
                i = i + 1
                self.line_number_box.insert(END, f'{i}\n', 'right')
            self.line_number_box.config(state=DISABLED, takefocus=NO)

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
        for w in self.maintexts:
            w.tag_config('search', background='blue', foreground='white')
            w.tag_config('search_selected', background='cyan', foreground='white')

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

        self.textboxes = self.maintexts + [self.dummy_maintext]
        try:
            self.textboxes.append(self.line_number_box)
        except AttributeError as e:
            pass

        self.align_number_of_rows()

        for w in self.maintexts:
            w.bind('<Button-3>', self.popup)
            w.bind('<Control-o>', self.file_open)
            for event in ['<Alt-Up>', '<Alt-Control-Up>', '<Alt-Down>', '<Alt-Control-Down>']:
                w.bind(event, self.handle_KeyPress_event_of_swap_lines)

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
        else:
            for w in self.textboxes:
                w.mark_set(INSERT, 1.0)
        self.set_text_widget_editable()

    # 以下、ファイルの開始、保存、終了に関するメソッド
    def file_create(self, e=None):
        if not self.file_close():
            return
        self.filepath = ''
        self.bookmarks = []
        self.use_regex_in_bookmarks = False
        self.data = self.data0.copy()
        self.edit_history = deque([self.data])
        self.undo_history = deque([])
        self.make_text_editor()

    def file_open(self, e=None, file_path_to_open: str|None=None):
        '''
        '''
        # 編集中のファイルに更新がある場合、ファイル終了ダイアログが表示される。
        # ファイル終了ダイアログでファイルを終了する選択を取らなかった場合、ファイル開始を中断する
        if self.data != self.get_current_data():
            if not self.file_close():
                return 'break'
        # 開くファイルを指定されている場合、ファイル選択ダイアログをスキップする
        if file_path_to_open:
            # ドライブ文字を大文字にする
            file_path_to_open = self.convert_drive_to_uppercase(file_path_to_open)
        else:
            file_path_to_open = filedialog.askopenfilename(
                title = '編集ファイルを選択',
                initialdir = self.initialdir,
                initialfile='file.sep',
                filetypes=[('SoroEditorプロジェクトファイル', '.sep'), ('YAMLファイル', '.yaml'), ('その他', '.*'), ('ThreeCrowsプロジェクトファイル', '.tcs'), ('CastellaEditorプロジェクトファイル', '.cep')],
                defaultextension='sep')

        log.info(f'Opening file: {file_path_to_open}')

        if file_path_to_open:
            try:
                with open(file_path_to_open, mode='rt', encoding='utf-8') as f:
                    newdata: dict = yaml.safe_load(f)
                    if 'version' in newdata.keys():
                        self.data = newdata
                    else:
                        raise(KeyError)
                    data = self.data0.copy()
                    data.update(newdata)

                    self.number_of_columns = data['columns']['number']
                    self.column_percentage = [x*0.01 for x in data['columns']['percentage']]
                    self.bookmarks = data['bookmarks']
                    self.use_regex_in_bookmarks = data['use_regex_in_bookmarks']
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

                self.initialdir = os.path.dirname(file_path_to_open)

                # 最近使用したファイルのリストを修正し、settings.yamlに反映
                self.recently_files.insert(0, file_path_to_open)
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
                    self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file_path_to_open=self.recently_files[1]))
                    self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file_path_to_open=self.recently_files[2]))
                    self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file_path_to_open=self.recently_files[3]))
                    self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file_path_to_open=self.recently_files[4]))
                except IndexError:
                    pass

                self.make_menu_bookmarks()

            except FileNotFoundError as e:
                log.error(e)
                md = MessageDialog(title='TreeCrows - エラー', alert=True, buttons=['OK'], message='ファイルが見つかりません')
                md.show(self.md_position)
                # 最近使用したファイルに見つからなかったファイルが入っている場合、削除しsettins.yamlに反映する
                if file_path_to_open in self.recently_files:
                    self.recently_files.remove(file_path_to_open)
                    self.settings['recently_files'] = self.recently_files
                # 設定ファイルに書き込み
                self.update_setting_file()
                # 更新
                try:
                    for _ in range(5):
                        self.menu_file_recently.delete(0)
                except TclError:
                    pass
                try:
                    if self.filepath:
                        self.menu_file_recently.add_command(label=self.filepath + ' (現在のファイル)')
                    else:
                        self.menu_file_recently.add_command(label=self.recently_files[0], command=lambda:self.file_open(file_path_to_open=self.recently_files[0]))
                    self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file_path_to_open=self.recently_files[1]))
                    self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file_path_to_open=self.recently_files[2]))
                    self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file_path_to_open=self.recently_files[3]))
                    self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file_path_to_open=self.recently_files[4]))
                except IndexError:
                    pass

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
                self.filepath = file_path_to_open
                log.info(f'Succeed opening file: {self.filepath}')
                self.align_the_lines(1.0, 1.0, False)

        return 'break'

    def file_save_as(self, e=None) -> bool:
        self.filepath = filedialog.asksaveasfilename(
            title='名前をつけて保存',
            initialdir=self.initialdir,
            initialfile='noname',
            filetypes=[('SoroEditorプロジェクトファイル', '.sep'), ('YAMLファイル', '.yaml'), ('その他', '.*')],
            defaultextension='sep'
        )
        if self.filepath:
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
                yaml.safe_dump(current_data, f, encoding='utf-8', allow_unicode=True)
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
            self.menu_file_recently.add_command(label=self.recently_files[1], command=lambda:self.file_open(file_path_to_open=self.recently_files[1]))
            self.menu_file_recently.add_command(label=self.recently_files[2], command=lambda:self.file_open(file_path_to_open=self.recently_files[2]))
            self.menu_file_recently.add_command(label=self.recently_files[3], command=lambda:self.file_open(file_path_to_open=self.recently_files[3]))
            self.menu_file_recently.add_command(label=self.recently_files[4], command=lambda:self.file_open(file_path_to_open=self.recently_files[4]))
        except IndexError:
            pass

    def ask_file_close(self, exit_on_completion=False) -> bool:
        '''
        ファイルを閉じるか使用者に確認するメソッド
        ファイルを閉じる場合はTrueを、中止する場合はFalseを返す
        '''
        current_data = self.get_current_data()

        if exit_on_completion:
            message = '更新内容が保存されていません。アプリを終了しますか。'
            title = 'SoroEditor - 終了確認'
            buttons=['保存して終了:success', '保存せず終了:danger', 'キャンセル']
        else:
            message = '更新内容が保存されていません。ファイルを閉じ、ファイルを変更しますか'
            title = 'SoroEditor - 確認'
            buttons=['保存して変更:success', '保存せず変更:danger', 'キャンセル']

        if self.data == current_data:
            return True
        if self.data != current_data:
            messagedialog = MessageDialog(message=message, title=title, buttons=buttons)
            messagedialog.show(self.md_position)
            if messagedialog.result in ('保存せず終了', '保存せず変更'):
                return True
            if messagedialog.result in ('保存して終了', '保存して変更'):
                if self.file_over_write_save():
                    return True
            if messagedialog.result == ('キャンセル', None):
                pass
            return False

    def file_close(self, exit_on_completion=False, need_confirmation=True) -> bool:
        '''
        ファイルを閉じるメソッド
        ファイルが閉じられた場合はTrueを、閉じられなかった場合はFalseを返す
        '''
        if need_confirmation:
            if not self.ask_file_close(exit_on_completion):
                return False

        self.filepath = ''
        self.bookmarks = []
        self.use_regex_in_bookmarks = False
        self.column_percentage = [x*0.01 for x in self.settings['columns']['percentage']]
        self.data = self.data0.copy()
        self.edit_history = deque([self.data])
        self.undo_history = deque([])
        self.make_text_editor()
        if exit_on_completion:
            self.master.destroy()
        return True

    def file_reload(self, e=None):
        first_data = self.data
        self.data = self.get_current_data()
        self.make_text_editor()
        self.data = first_data

    def open_last_file(self, e=None):
        if not self.filepath:
            self.file_open(file_path_to_open=self.recently_files[0])

    def import_from_csv(self, e=None, mode=0, filepath=None, encoding='utf-8', delimiter=','):
        '''
        mode:int[0, 1]
            - mode=0: CSVの1行目をタイトルとする
            - mode=1: すべて本文としてインポートする
        '''
        if not filepath:
            filepath = filedialog.askopenfilename(
                title='CSV形式のファイルを選択',
                initialdir=self.initialdir,
                defaultextension='csv',
                filetypes=(('CSVファイル', 'csv'), ('その他', '.*'))
                )
        try:
            with open(filepath, 'rt', encoding=encoding, newline='') as f:
                reader = csv.reader(f, delimiter=delimiter)
                data = []
                for row in reader:
                    data.append(row)
        except FileNotFoundError:
            pass
        except PermissionError:
            pass
        except UnicodeDecodeError:
            pass
        except csv.Error:
            pass
        except ValueError:
            pass
        else:
            self.file_create()
            newdata = self.data0.copy()
            if mode == 0:
                for j, row in enumerate(data):
                    for i, text in enumerate(row):
                        if j == 0:
                            newdata[i]['title'] = text
                        else:
                            newdata[i]['text'] += text
                            newdata[i]['text'] += '\n'
            if mode == 1:
                for row in data:
                    for i, text in enumerate(row):
                        newdata[i]['text'] += text
                        newdata[i]['text'] += '\n'
            self.data = newdata
            self.make_text_editor()

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

    def now_time_set(self):
        mode = self.clock_mode
        if mode == 'ymdhms':
            self.now.set(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S'))
        if mode =='ymdhm':
            self.now.set(datetime.datetime.now().strftime('%Y/%m/%d %H:%M'))
        if mode =='mdhms':
            self.now.set(datetime.datetime.now().strftime('%m/%d %H:%M:%S'))
        if mode =='mdhm':
            self.now.set(datetime.datetime.now().strftime('%m/%d %H:%M'))
        if mode =='hms':
            self.now.set(datetime.datetime.now().strftime('%H:%M:%S'))
        if mode =='hm':
            self.now.set(datetime.datetime.now().strftime('%H:%M'))

        self.master.after(100, self.now_time_set)

    def clock_change(self, _):
        modes = ['ymdhm', 'ymdhms', 'mdhm', 'mdhms', 'hm', 'hms']
        try:
            index = modes.index(self.clock_mode)
        except ValueError:
            index = 5
        if index == 5:
            self.clock_mode = modes[0]
        else:
            self.clock_mode = modes[index+1]
        self.settings['clock_mode'] = self.clock_mode
        self.update_setting_file()
        self.now_time_set()

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
        current_data['bookmarks'] = self.bookmarks
        current_data['use_regex_in_bookmarks'] = self.use_regex_in_bookmarks
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

    def search_on_web(self, search_engine_url):
        '''
        選択中のテキストを読み取って検索エンジンに入力し、ウェブブラウザを開くメソッドを返す
        search_engine_url: str 検索エンジンのURL。例: 'https://www.google.com/search?q=%s'
        '''
        def inner(*_):
            widget = self.master.focus_get()
            text = ''
            if widget.winfo_class() =='Text':
                if widget.tag_ranges(SEL):
                    text = widget.get(SEL_FIRST, SEL_LAST)
            elif widget.winfo_class() =='TEntry':
                text = widget.get()

            if not text:
                return

            try:
                webbrowser.open(search_engine_url % text)
            except TypeError:
                MessageDialog(
                    '検索エンジンエラー\nURLの検索クエリ部分に\'%s\'を入れる\n例: \'https://www.google.com/search?q=%s\'',
                    'SoroEditor - 検索エンジンエラー',
                    ['OK']
                    ).show()

        return inner

    def search(self, search_text:str, use_regular_expression=False, title='', entry=None) -> list:
        if not search_text:
            return []
        current_data = self.get_current_data()
        texts = [current_data[i]['text'] for i in range(10) if current_data[i]['text']]
        if not use_regular_expression:
            search_text = re.escape(search_text)
        results = []
        try:
            for line, text in enumerate(texts):
                newrow_charactors_matches:list[tuple[int, int]] = [(0, 0)]
                newrow_charactors_matches.extend([m.span() for m in list(re.finditer('\n', text))])
                search_text_matches:list[tuple[int, int]] = [m.span() for m in list(re.finditer(search_text, text))]
                # 各行の開始文字数・終了文字数をリスト化する
                rows_span = [(newrow_charactors_matches[i][1], newrow_charactors_matches[i+1][0]) for i in range(len(newrow_charactors_matches)-1)]
                rows_span.append((rows_span[-1][1]+1, len(text)))
                # 検索結果が何行目の何文字目かを確認する
                for start, stop in search_text_matches:
                    row_first = -1
                    for row, t in enumerate(rows_span):
                        i, j = t
                        if i <= start <= j:
                            row_first = row
                            break
                    start = start - rows_span[row][0]
                    row_last = -1
                    for row, t in enumerate(rows_span):
                        i, j = t
                        if i <= stop <= j:
                            row_last = row
                            break
                    stop = stop - rows_span[row][0]
                    # 結果を追加する
                    results.append((line, (row_first, row_last), (start, stop)))
        except re.error:
            results = []
            if title and entry:
                messagedialog = MessageDialog(
                    '正規表現エラー\n検索内容を確認してください',
                    title + ' 正規表現エラー',
                    alert=True,
                    buttons=['OK'],
                    parent=entry)
                messagedialog.show(app.md_position)
                entry.focus()
        return results

    def cut_copy(self, mode=0):
        '''
        mode: int[0, 1]
            mode=0: cut
            mode=1: copy
        '''
        widget = self.master.focus_get()
        widget_class = widget.winfo_class()

        if widget_class == 'Text':
            if widget.tag_ranges(SEL):
                t = widget.get(SEL_FIRST, SEL_LAST)
                if mode == 0:
                    widget.delete(SEL_FIRST, SEL_LAST)
            else:
                return
                # issue58
                # 選択されたテキストがない場合、カーソルのある行をコピー・カットするコード
                # ショートカットキーでmode=0を呼び出す時正常に動作しない
                # (OSのカット機能と競合しクリップボードに贈られるテキストが壊れる)原因になっているため無効化
                # 解決策が見つかり次第修正したい
                t = widget.get(INSERT+' linestart', INSERT+'+1line linestart')
                if mode == 0:
                    widget.delete(INSERT+' linestart', INSERT+'+1line linestart')

        elif widget_class == 'TEntry':
            t = widget.get()
            if mode == 0:
                widget.delete(0, END)

        self.clipboard_clear()
        self.clipboard_append(t)
        if mode == 0:
            self.recode_edit_history()

    def paste(self, e=None):
        widget = self.master.focus_get()
        widget_class = widget.winfo_class()
        try:
            t = self.clipboard_get()
        except TclError:
            return
        if widget_class == 'Text':
            if widget.tag_ranges(SEL):
                widget.delete(SEL_FIRST, SEL_LAST)
        widget.insert(INSERT, t)
        self.recode_edit_history()

    def swap_lines(self, mode='up', widget=None):
        '''
        Swap selected lines or current line up or down in a text widget.

        Parameters:
        - mode (str): Swap direction, either 'up' or 'down'.
        - widget (tk.Text): The Text widget to perform the swap in.

        '''
        if not widget:
            return

        sel_first, sel_last = '', ''
        insert = widget.index(INSERT)

        if widget.tag_ranges(SEL):
            sel_first, sel_last = widget.index(SEL_FIRST), widget.index(SEL_LAST)
            swap_index = (SEL_FIRST+' linestart', SEL_LAST+'+1line linestart')
        else:
            swap_index = (INSERT+' linestart', INSERT+'+1line linestart')

        text = widget.get(*swap_index)
        widget.delete(*swap_index)

        if mode == 'up':
            move_to_index = INSERT+'-1line linestart'
            insert_index = insert+'-1line'
            sel_first, sel_last = sel_first+'-1line', sel_last+'-1line'
        if mode == 'down':
            move_to_index = INSERT+'+1line linestart'
            insert_index = insert+'+1line'
            sel_first, sel_last = sel_first+'+1line', sel_last+'+1line'

        widget.insert(move_to_index, text)
        widget.mark_set(INSERT, insert_index)
        try:
            if sel_first and sel_last:
                widget.tag_add(SEL, sel_first, sel_last)
        except TclError:
            pass

    def swap_lines_in_all_boxes(self, mode='up', widget=None):
        '''
        Swap selected lines or current line up or down in all text boxes.

        Parameters:
        - widget (tk.Text): The Text widget to perform the swap in.
        - mode (str): Swap direction, either 'up' or 'down'.

        '''
        if not widget:
            return

        insert = widget.index(INSERT)
        sel_first, sel_last = insert, insert

        if widget.tag_ranges(SEL):
            sel_first, sel_last = widget.index(SEL_FIRST), widget.index(SEL_LAST)

        self.set_text_widget_editable(mode=1)
        for box in self.maintexts:
            box.mark_set(INSERT, insert)
            box.tag_add(SEL, sel_first, sel_last)
            self.swap_lines(mode, box)
            if box != widget:
                box.tag_remove(SEL, 1.0, END)
        self.set_text_widget_editable(mode=0)

    def handle_KeyPress_event_of_swap_lines(self, event):
        if event.keysym == 'Up' and event.state == 393216:# Alt-Up
            self.swap_lines('up', event.widget)
        elif event.keysym == 'Up' and event.state == 393220:# Alt-Ctrl-Up
            self.swap_lines_in_all_boxes('up', event.widget)
        elif event.keysym == 'Down' and event.state == 393216:# Alt-Down
            self.swap_lines('down', event.widget)
        elif event.keysym == 'Down' and event.state == 393220:# Alt-Ctrl-Down
            self.swap_lines_in_all_boxes('down', event.widget)

        return 'break'

    def select_all(self, e=None):
        widget = self.master.focus_get()
        widget_class = widget.winfo_class()
        if widget_class == 'Text':
            widget.tag_add(SEL, '1.0', END+'-1c')
        elif widget_class == 'TEntry':
            pass

    def select_line(self, e=None):
        widget = self.master.focus_get()
        widget_class = widget.winfo_class()
        if widget_class == 'Text':
            widget.tag_add(SEL, INSERT+' linestart', INSERT+' lineend')
        elif widget_class == 'TEntry':
            pass

    def newline(self, e, mode=0):
        '''
        行を追加する

        mode: int=[0, 1]
            mode=0: 上に追加
            mode=1: 下に追加
        '''
        if type(e.widget) == Text:

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

    def insert_template_to_maintext(self, num):
        text = self.templates[num]
        widget = self.focus_lastfor()
        if widget in self.maintexts+self.entrys:
            widget.insert(INSERT, text)

    def menu_templates_clicked(self, num):
        def inner(*_):
            self.insert_template_to_maintext(num)
        return inner

    def open_SearchWindow_with_bookmarks(self, bookmark):
        self.open_SearchWindow('0', bookmark)

    def menu_bookmarks_clicked(self, bookmark):
        def inner(*_):
            self.open_SearchWindow_with_bookmarks(bookmark)
        return inner

    def make_menu_edit(self, parent:Menu):
        parent.add_command(label='切り取り(T)', command=lambda: self.cut_copy(), accelerator='Ctrl+X', underline=5)
        parent.add_command(label='コピー(C)', command=lambda: self.cut_copy(1), accelerator='Ctrl+C', underline=4)
        parent.add_command(label='貼り付け(P)', command=self.paste, accelerator='Ctrl+V', underline=5)
        parent.add_command(label='すべて選択(A)', command=self.select_all, accelerator='Ctrl+A', underline=6)
        parent.add_command(label='1行選択(L)', command=self.select_line, underline=5)
        parent.add_command(label='取り消し(U)', command=self.undo, accelerator='Ctrl+Z', underline=5)
        parent.add_command(label='取り消しを戻す(R)', command=self.repeat, accelerator='Ctrl+Shift+Z', underline=8)

    def make_menu_templates(self):
        self.menu_templates.delete(2, 12)
        for i in range(10):
            num = i + 1 if i != 9 else 0
            try:
                text = repr(self.templates[i])[1:-1]
            except SyntaxError:
                text = repr(self.templates[i][:-1])[1:-1]
            label = f'({num}): {text[:10]}...' if len(text) > 10 else f'({num}): {text}'
            self.menu_templates.add_command(label=label,
                                            command=self.menu_templates_clicked(i),
                                            accelerator=f'Ctrl or Alt+{i}',
                                            underline=1)

    def make_menu_bookmarks(self):
        self.menu_bookmark.delete(2, 100)
        if not self.bookmarks:
            self.menu_bookmark.add_command(label='付箋なし')
        for i, bookmark in enumerate(self.bookmarks, 1):
            i = i if i < 10 else 0
            label = f'({i}): {bookmark}を検索'
            command = self.menu_bookmarks_clicked(bookmark)
            self.menu_bookmark.add_command(label=label, command=command, underline=1)

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
        if self.wrap != NONE:
            return

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
        log.info('---Backup---')

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
        new_data = yaml.safe_dump(current_data, allow_unicode=True)

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
            log.error(f'---Failed Backup---\n: {e}')
            return

        # バックアップ完了メッセージをログに出力する
        log.info(f'{backup_filepath}\n---Succeed Backup---')

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
        log.info(f'---Autosave---')
        if self.file_over_write_save():
            log.info(f'---Succeed Autosave---')
        else:
            log.error(f'---Failed Autosave---')
        self.master.after(self.autosave_frequency, self.autosave)

    def popup(self, e=None):
        '''右クリックメニュー'''
        # 表示するメニューの生成
        menu = Menu()
        self.make_menu_edit(menu)
        menu.add_separator()
        menu.add_cascade(label='検索(S)', menu=self.menu_search)
        menu.add_cascade(label='付箋(B)', menu=self.menu_bookmark)
        menu.add_cascade(label='定型文(T)', menu=self.menu_templates)

        e.widget.focus_set()
        menu.post(e.x_root,e.y_root)

    def open_SettingWindow(self, *_):
        '''重複を防ぐ'''
        return self.open_sub_window('setting')

    def open_SearchWindow(self, mode='0', searchtext=''):
        '''重複を防ぐ'''
        if not mode in ('0', '1'):
            mode = '0'
        return self.open_sub_window('search', mode, searchtext)

    def open_TemplateWindow(self, *_):
        '''重複を防ぐ'''
        return self.open_sub_window('template')

    def open_BookmarkWindow(self, *_):
        '''重複を防ぐ'''
        return self.open_sub_window('bookmark')

    def open_sub_window(self, sub_window:str, *args):
        '''重複を防ぐ'''
        def get_window(sub_window):
            return {
                'setting': SettingWindow,
                'search': SearchWindow,
                'template': TemplateWindow,
                'bookmark': BookmarkWindow,
            }[sub_window]

        if self.windows[sub_window] is None or not self.windows[sub_window].winfo_exists():
            self.windows[sub_window] = get_window(sub_window)(*args)
        else:
            self.windows[sub_window].focus()

    def close_sub_window(self, sub_window:Toplevel):
        def inner(*_):
            log.info(f'---Close {sub_window.__class__.__name__}---')
            sub_window.destroy()
        return inner

    def set_icon_sub_window(self, sub_window:Toplevel):
        sub_window.iconphoto(False, Icons().icon)


class SettingWindow(Toplevel):
    '''
    設定ウィンドウに関する設定
    '''
    def __init__(self, e=None, title="SoroEditor - 設定", iconphoto='', size=(1200, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        self.close = app.close_sub_window(self)

        self.protocol('WM_DELETE_WINDOW', self.close)

        app.set_icon_sub_window(self)

        log.info('---Open SettingWindow---')
        # 設定ファイルの読み込み
        with open('./settings.yaml', mode='rt', encoding='utf-8') as f:
            self.settings = yaml.safe_load(f)
        number_of_columns = self.settings['columns']['number']
        percentage_of_columns = self.settings['columns']['percentage']
        self.themename = self.settings['themename']
        self.button_style = self.settings['button_style']
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

        self.pair_of_status_bar_elements = [
            ['文字数カウンター - 1列目', 'letter_count_1'],
            ['文字数カウンター - 2列目', 'letter_count_2'],
            ['文字数カウンター - 3列目', 'letter_count_3'],
            ['文字数カウンター - 4列目', 'letter_count_4'],
            ['文字数カウンター - 5列目', 'letter_count_5'],
            ['行数カウンター - 1列目', 'line_count_1'],
            ['行数カウンター - 2列目', 'line_count_2'],
            ['行数カウンター - 3列目', 'line_count_3'],
            ['行数カウンター - 4列目', 'line_count_4'],
            ['行数カウンター - 5列目', 'line_count_5'],
            ['カーソルの現在位置', 'current_place'],
            ['ショートカットキー1', 'hotkeys1'],
            ['ショートカットキー2', 'hotkeys2'],
            ['ショートカットキー3', 'hotkeys3'],
            ['各機能情報', 'infomation'],
            ['顔文字', 'kaomoji'],
            ['ステータスバー初期メッセージ', 'statusbar_message'],
            ]
        self.pair_of_tool_bar_elements = [
            ['ボタン - 新規作成', 'toolbutton_create'],
            ['ボタン - ファイルを開く', 'toolbutton_open'],
            ['ボタン - 上書き保存', 'toolbutton_save'],
            ['ボタン - 名前をつけて保存', 'toolbutton_save_as'],
            ['ボタン - 再読込', 'toolbutton_file_reload'],
            ['ボタン - プロジェクト設定', 'toolbutton_project_setting'],
            ['ボタン - 設定', 'toolbutton_setting'],
            ['ボタン - 検索', 'toolbutton_search'],
            ['ボタン - 置換', 'toolbutton_replace'],
            ['ボタン - インポート', 'toolbutton_import'],
            ['ボタン - エクスポート', 'toolbutton_export'],
            ['ボタン - 定型文', 'toolbutton_template'],
            ['ボタン - 付箋', 'toolbutton_bookmark'],
            ['ボタン - 取り消し', 'toolbutton_undo'],
            ['ボタン - 取り消しを戻す', 'toolbutton_repeat'],
            ]
        self.statusbar_elements_dict_converted = {}
        for i in range(7):
            try:
                l = self.statusbar_elements_dict[i]
                if l == None:
                    raise KeyError
            except KeyError:
                l = []
            self.statusbar_elements_dict_converted[i] = [self.convert_statusbar_elements(val) for val in l]
            while len(self.statusbar_elements_dict_converted[i]) < 5:
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
        Label(lf2_3, text='*折り返し (none以外にすると同期スクロールが無効になります)').pack(anchor=W)
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
        ### ボタンスタイル
        Label(lf4, text='  *ボタンスタイル: ').pack(side=LEFT)
        button_styles = ['icon_with_text', 'icon_only', 'text_only']
        self.setting_button_style = StringVar()
        self.setting_button_style_menu = OptionMenu(lf4, self.setting_button_style, self.button_style, *button_styles)
        self.setting_button_style_menu.pack(side=LEFT)

        # タブ - ステータスバー
        f3 = Frame(nt, padding=5)
        for i in range(4):
            f3.grid_columnconfigure(i+1, weight=1)
        statusbar_elements = [None] + [l[0] for l in self.pair_of_status_bar_elements]
        toolbar_elements = [None] + [l[0] for l in self.pair_of_tool_bar_elements]
        self.setting_statusbar_elements_dict = {i: {'var': [], 'menu': []} for i in range(7)}

        for i, j in enumerate([4, 5, 6, 3, 2, 1, 0]):
            row = i*2 + 1
            if i < 3:
                text = f'ツールバー{i+1}'
            else:
                text = f'ステータスバー{i-2}'
            Label(f3, text=text).grid(row=row, column=0, columnspan=2, sticky=W)
            i = (i + 1) * 2
            for k in range(5):
                if j < 4:
                    l = statusbar_elements
                else:
                    l = toolbar_elements
                self.setting_statusbar_elements_dict[j]['var'].append(StringVar())
                self.setting_statusbar_elements_dict[j]['menu'].append(OptionMenu(f3, self.setting_statusbar_elements_dict[j]['var'][k], self.statusbar_elements_dict_converted[j][k], *l))
                self.setting_statusbar_elements_dict[j]['menu'][k].grid(row=i, column=k, padx=5, pady=5, sticky=W+E)

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

        self.set_window()

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
        button_style = self.setting_button_style.get()
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
        statusbar_method = {i: [] for i in range(7)}
        for i in range(7):
            for j in range(5):
                e = self.convert_statusbar_elements(self.setting_statusbar_elements_dict[i]['var'][j].get())
                if e:
                    statusbar_method[i].append(e)

        # 設定を辞書に格納する
        self.settings['columns']['number'] = number_of_columns
        self.settings['columns']['percentage'] = percentage_of_columns
        self.settings['font']['family'] = font_family
        self.settings['font']['size'] = font_size
        self.settings['between_lines'] = between_lines
        self.settings['wrap'] = wrap
        self.settings['themename'] = themename
        self.settings['button_style'] = button_style
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
        except TclError as e:
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
        else:
            self.set_window()

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
        for x in self.pair_of_status_bar_elements + self.pair_of_tool_bar_elements:
            if val == x[0]: return x[1]
            elif val == x[1]: return x[0]
        return ''

    def set_window(self):
        # 設定ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)


class ProjectFileSettingWindow(Toplevel):
    def __init__(self, title="SoroEditor - プロジェクト設定", iconphoto='', size=(600, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        self.close = app.close_sub_window(self)

        self.protocol('WM_DELETE_WINDOW', self.close)

        app.set_icon_sub_window(self)

        log.info('---Open ProjectFileSettingWindow---')
        if app.get_current_data() != app.data:
            self.withdraw()
            messagedialog = MessageDialog('プロジェクトファイル設定を変更する前にプロジェクトファイルを保存します',
                        'SoroEditor - プロジェクトファイル設定',
                        ['OK:success', 'キャンセル:secondary'],)
            messagedialog.show(app.md_position)
            if messagedialog.result == 'OK':
                x = app.file_over_write_save()
                if x:
                    self.deiconify()
                else:
                    self.destroy()
                    return
            elif messagedialog.result == 'キャンセル':
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


class ThirdPartyNoticesWindow(Toplevel):
    def __init__(self, title="SoroEditor - ライセンス情報", iconphoto='', size=(1200, 800), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        app.set_icon_sub_window(self)

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
                log.info(f'Open {url}')
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

        app.set_icon_sub_window(self)

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
        t.tag_config('text', font=font_text, spacing1=5, spacing2=5)
        t.tag_config('github', underline=True)
        t.tag_config('github_issue', underline=True)
        t.tag_config('mail', underline=True)
        github = open_url('https://github.com/joppincal/SoroEditor')
        github_issue = open_url('https://github.com/joppincal/SoroEditor/issues')
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
・列数/列幅
・フォント（ファミリー・サイズ）
・表示テーマ
・画面端の折り返し（なし 文字で折り返し 単語で折り返し（英語向け））
・ステータスバー（ウィンドウ下部の、様々な情報が記述されている場所）
・ツールバー（ウィンドウ上部のボタンが配置された場所）
・編集中の行の強調表示（デフォルトでは下線）
・行番号
・起動時のウィンドウサイズ
・自動保存
・バックアップ

*列数・列幅に関しては、ファイル(F)→プロジェクト設定(F)	から
プロジェクトファイルごとに設定できます\n''', 'text')
        t.insert(END, 'ホットキー\n', 'h1')
        t.insert(END,
'''Ctrl+O:           ファイルを開く
Ctrl+R:           前回使用したファイルを開く
Ctrl+S:           上書き保存
Ctrl+Shift+S:     名前をつけて保存
Ctrl+Shift+E:     エクスポート
Ctrl+Shift+P:     設定

Ctrl+F:           検索
Ctrl+Shift+F:     置換

Ctrl+Z:           取り消し
Ctrl+Shift+Z:     取り消しを戻す
Ctrl+Enter:       1行追加（下）

Ctrl+Shift+Enter: 1行追加（上）
Ctrl+L:           1行選択
Ctrl+Q, Alt+<:    左の列に移動
Ctrl+W, Alt+>:    右の列に移動\n''', 'text')
        t.insert(END, '検索/置換\n', 'h1')
        t.insert(END,
'''検索/置換機能はメニューバーから、またはショートカットキーからアクセスできます
検索/置換にはPythonの正規表現を用いる事もできます\n''', 'text')
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
    def __init__(self, title="SoroEditorについて", iconphoto='', size=(550, 700), position=None, minsize=None, maxsize=None, resizable=(False, False), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        app.set_icon_sub_window(self)

        frame = Frame(self, padding=5)
        frame.pack(fill=BOTH, expand=True)
        main = Text(frame)

        icon = Image.open('src/icon/icon.png')
        icon = icon.resize((300, 300))
        icon = ImageTk.PhotoImage(icon)
        iconlabel = Label(self, image=icon)

        main.tag_config('title', justify=CENTER, spacing2=10, font=font.Font(size=15, weight='bold'))
        main.tag_config('text', justify=CENTER, spacing2=10, font=font.Font(size=12, weight='normal'))
        main.tag_config('link', justify=CENTER, spacing2=10, font=font.Font(size=12, weight='normal', underline=True))
        github = 'https://github.com/joppincal/SoroEditor'
        main.tag_config('github')
        main.tag_bind('github', '<Button-1>', lambda _: webbrowser.open_new(github))
        homepage = ''
        main.tag_config('homepage')
        main.tag_bind('homepage', '<Button-1>', lambda _: webbrowser.open_new(homepage))

        main.insert(END, '\nそろエディタ\nSoroEditor\n', 'title')
        main.window_create(END, window=iconlabel, pady=30)
        main.tag_add('title', iconlabel)
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


class ImportWindow(Toplevel):
    def __init__(self, title="SoroEditor - インポート", iconphoto='', size=(500, 700), position=None, minsize=None, maxsize=None, resizable=(0, 0), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        log.info('---Open ImportWindow---')

        self.close = app.close_sub_window(self)

        self.protocol('WM_DELETE_WINDOW', self.close)

        app.set_icon_sub_window(self)

        self.current_data = app.get_current_data()

        f = Frame(self, padding=10)
        f.pack(fill=BOTH, expand=True)

        self.f1 = Frame(f, padding=5)
        self.f1.pack(side=TOP, fill=X)

        Label(self.f1, text='ファイル形式: ').grid(row=0, column=0, sticky=E)
        self.file_format = StringVar(value='CSV')
        file_formats = ['CSV', 'TSV']
        OptionMenu(self.f1, self.file_format, self.file_format.get(), *file_formats).grid(row=0, column=1, sticky=EW)

        Label(self.f1, text='文字コード: ').grid(row=1, column=0, sticky=E)

        Label(self.f1, text='文字コード: ').grid(row=1, column=0, sticky=E)
        self.encoding = StringVar(value='UTF-8')
        self.encoding.trace_add('write', self.set_optionmenu_title)
        encodings = ['UTF-8', 'Shift-JIS']
        OptionMenu(self.f1, self.encoding, self.encoding.get(), *encodings).grid(row=1, column=1, sticky=EW)

        self.set_optionmenu_title()

        Separator(self.f1).grid(row=3, column=0, columnspan=2, sticky=EW)

        Label(self.f1, text='ファイル: ').grid(row=4, column=0, sticky=E)
        self.filepath = StringVar(value="P:\My Video\編集データ\台本\看取ってください、私のマスター！.csv")
        self.filepath_label = Label(self.f1, textvariable=self.filepath, wraplength=350)
        self.filepath_label.grid(row=4, column=1, sticky=EW)
        Button(self.f1, text='ファイル変更', command=self.change_filepath).grid(row=5, column=1, sticky=E)

        self.f1.grid_columnconfigure(1, weight=1)
        self.f1.grid_rowconfigure(0, pad=10)
        self.f1.grid_rowconfigure(1, pad=10)
        self.f1.grid_rowconfigure(2, pad=10)
        self.f1.grid_rowconfigure(3, pad=10)
        self.f1.grid_rowconfigure(4, pad=10)

        f2 = Frame(f, padding=(5, 5))
        Button(f2, text='キャンセル', command=self.close, bootstyle=SECONDARY).pack(side=RIGHT, padx=2)
        Button(f2, text='インポート', command=self.import_).pack(side=RIGHT, padx=2)
        Separator(f2).pack(fill=X)
        f2.pack(side=BOTTOM, fill=X, anchor=S)

        # ウィンドウの設定
        self.grab_set()
        self.focus_set()
        self.transient(app)
        app.wait_window(self)

    def set_optionmenu_title(self, *_):
        file_format = self.file_format.get()
        Label(self.f1, text='タイトル: ').grid(row=2, column=0, sticky=E)
        self.title = StringVar(value=f'{file_format}の1行目をタイトルとする')
        titles = [f'{file_format}の1行目をタイトルとする', f'{file_format}の1行目も本文とする']
        OptionMenu(self.f1, self.title, self.title.get(), *titles).grid(row=2, column=1, sticky=EW)

    def change_filepath(self):
        filepath = filedialog.askopenfilename(
            defaultextension='csv',
            filetypes=(['CSVファイル', '.csv'], ['TSVファイル', '.tsv'], ['その他', '.*']),
            initialdir=app.initialdir,
            title='インポートするファイルを選択'
            )
        if filepath:
            self.filepath.set(filepath)
        else:
            return

    def import_(self):
        filepath = self.filepath.get()
        file_format = self.file_format.get()
        encoding = self.encoding.get()
        title = self.title.get()
        delimiter = ','
        mode = 0

        if file_format == 'CSV':
            delimiter = ','
        if file_format == 'TSV':
            delimiter = '\t'
        if title == f'{file_format}の1行目をタイトルとする':
            mode = 0
        if title == f'{file_format}の1行目も本文とする':
            mode = 1

        app.import_from_csv(mode=mode, filepath=filepath, encoding=encoding, delimiter=delimiter)


class ExportWindow(Toplevel):

    def __init__(self, title="SoroEditor - エクスポート", iconphoto='', size=(1000, 700), position=None, minsize=None, maxsize=None, resizable=(0, 0), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        log.info('---Open ExportWindow---')

        self.close = app.close_sub_window(self)

        self.protocol('WM_DELETE_WINDOW', self.close)

        app.set_icon_sub_window(self)

        self.current_data = app.get_current_data()

        f = Frame(self, padding=10)
        f.pack(fill=BOTH, expand=True)

        f1 = Frame(f, padding=5)
        Label(f1, text='ファイル形式').pack(side=LEFT, padx=2)
        self.file_format = StringVar(value='CSV')
        file_formats = ['CSV', 'TSV', 'テキスト', 'テキスト(Shift-JIS)']
        OptionMenu(f1, self.file_format, self.file_format.get(), *file_formats, command=self.make_setting_frame).pack(side=LEFT, padx=2)
        self.shift_jis_alert = Label(f1)
        self.shift_jis_alert.pack(side=LEFT, padx=10)
        f1.pack(fill=X)

        f2 = Frame(f, padding=(5, 5))
        Button(f2, text='キャンセル', command=self.close, bootstyle=SECONDARY).pack(side=RIGHT, padx=2)
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
        delimiter = {'CSV': ',', 'TSV': '\t'}.get(self.file_format.get())

        # ヘッダー行を追加する
        if self.include_title_in_output.get():
            # ヘッダー行のデータを取得する
            header = [data[i][0] for i in range(data_length)]
            # ヘッダー行を出力データに追加する
            output_data += f'{delimiter}'.join(header) + '\n'

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
            output_data += f'{delimiter}'.join(row_data) + '\n'

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
            encoding = 'cp932'
        elif self.file_format.get() in ['CSV', 'TSV']:
            encoding = 'utf-8-sig'
        else:
            encoding = 'utf-8'
        # ファイルを作成し、書き込む
        try:
            with open(self.filepath.get(), mode='wt', encoding=encoding, errors='replace') as f:
                f.write(data)
        except PermissionError as e:
            log.error(f'Failed export {self.file_format.get()}: {self.filepath.get()} {type(e).__name__}: {e}')
            return False
        except Exception as e:
            log.error(f'Failed export {self.file_format.get()}: {self.filepath.get()} {type(e).__name__}: {e}')
            return False
        else:
            log.info(f'Succeed export {self.file_format.get()}: {self.filepath.get()}')
            # 書き込みを行い成功した場合Trueを返す
            return True


class SearchWindow(Toplevel):

    def __init__(self, title='', iconphoto='', size=None, position=None, minsize=None, maxsize=None, resizable=(0, 0), transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        log.info('---Open SearchWindow---')

        self.protocol('WM_DELETE_WINDOW', self.close)

        app.set_icon_sub_window(self)

        if self.title() == '0':
            self.win_title = 'SoroEditor - 検索'
            self.mode = 0
        if self.title() == '1':
            self.win_title = 'SoroEditor - 置換'
            self.mode = 1
        else:
            self.win_title = 'SoroEditor - 検索'
            self.mode = 0
        self.title(self.win_title)

        searchtext = iconphoto

        if self.mode == 0:
            log.info('Search Mode')
        if self.mode == 1:
            log.info('Replace Mode')

        self.config(padx=10, pady=10)

        Label(self, text='検索').pack(padx=10, pady=0, anchor=W)
        self.text_in_entry = StringVar()
        self.entry = Entry(self, textvariable=self.text_in_entry, width=50)
        self.entry.bind('<Return>', lambda _: self.entry_return())
        self.entry.bind('<Shift-Return>', lambda _: self.entry_return(True))
        self.entry.pack(fill=X, padx=10, pady=10)
        self.entry.focus()
        self.use_regular_expression = BooleanVar()
        Checkbutton(
            self,
            variable=self.use_regular_expression,
            text='正規表現を用いる',
            takefocus=False,
            bootstyle='round-toggle'
            ).pack(padx=10,pady=0,anchor=W)

        if self.mode == 1:
            Label(self, text='置換').pack(padx=10, pady=0, anchor=W)
            self.text_in_entry2 = StringVar()
            self.entry2 = Entry(self, textvariable=self.text_in_entry2, width=50)
            self.entry2.bind('<Return>', lambda _: self.entry2_return())
            self.entry2.bind('<Shift-Return>', lambda _: self.entry2_return(True))
            self.entry2.pack(fill=X, padx=10, pady=10)

        f2 = Frame(self, padding=5)
        f3 = Frame(self, padding=5)
        self.search_button = Button(f2, text='検索', command=self.search_button_clicked)
        self.search_button.pack(side=RIGHT, padx=2)
        if self.mode == 1:
            self.replace_next_button = Button(f3, text='置換して次へ', command=self.replace_button_clicked)
            self.replace_next_button.pack(side=RIGHT, padx=2)
            self.replace_prev_button = Button(f3, text='置換して前へ', command=lambda: self.replace_button_clicked(-1))
            self.replace_prev_button.pack(side=RIGHT, padx=2)
            self.replace_all_button = Button(f3, text='すべて置換', command=lambda: self.replace_button_clicked(1))
            self.replace_all_button.pack(side=RIGHT, padx=2)
        self.next_button = Button(f2, text='次へ', command=self.next_button_clicked, bootstyle='secondary')
        self.next_button.pack(side=RIGHT, padx=2)
        self.prev_button = Button(f2, text='前へ', command=self.prev_button_clicked, bootstyle='secondary')
        self.prev_button.pack(side=RIGHT, padx=2)
        if self.mode == 0:
            self.switch_button = Button(f2, text='置換へ', command=self.switch_mode)
        elif self.mode == 1:
            self.switch_button = Button(f2, text='検索へ', command=self.switch_mode)
        self.switch_button.pack(side=RIGHT, padx=2)
        f3.pack(side=BOTTOM, fill=X, anchor=S)
        f2.pack(side=BOTTOM, fill=X, anchor=S)

        self.results = deque([])

        self.bind('<Escape>', lambda e: self.close())

        self.loop_make_results()

        # ウィンドウの設定
        self.geometry('+{0}+{1}'.format(*app.md_position))
        self.transient(app)

        # 検索文字列が指定されている場合、検索する
        if searchtext:
            self.text_in_entry.set(searchtext)
            self.use_regular_expression.set(app.use_regex_in_bookmarks)

    def search_button_clicked(self):
        self.next_button_clicked()

    def loop_make_results(self):
        self.make_results()
        self.after(100, self.loop_make_results)

    def make_results(self) -> bool:
        results = deque(app.search(self.text_in_entry.get(), self.use_regular_expression.get(), self.win_title, self.entry))
        if results and set(results) == set(self.results):
            return False
        for w in app.maintexts:
            w.tag_remove('search', 1.0, END)
            w.tag_remove('search_selected', 1.0, END)
        if not self.text_in_entry.get() or not results:
            self.title(self.win_title + ' 0件')
            self.results = deque([])
            return False
        self.results = results
        self.results.rotate(1)
        self.last_result = self.results[0]
        for line, row, span in self.results:
            app.maintexts[line].tag_add('search', f'{row[0]+1}.{span[0]}', f'{row[1]+1}.{span[1]}')
        self.title(self.win_title + f'{len(self.results)}件')
        return True

    def select(self, index=0):
        if not self.results:
            return
        num_of_results = len(self.results)
        try:
            results_index = self.results.index(self.last_result)
        except IndexError:
            results_index = 1
        line, row, span = self.results[index]
        row = row[0] + 1
        for w in app.textboxes:
            w.tag_remove('search_selected', 1.0, END)
        app.maintexts[line].focus()
        app.maintexts[line].mark_set(INSERT, f'{row}.{span[1]}')
        app.maintexts[line].see(f'{row}.{span[1]}')
        app.align_the_lines(repeat=False)
        app.maintexts[line].see(f'{row}.{span[1]}')
        app.maintexts[line].tag_add('search_selected', f'{row}.{span[0]}', f'{row}.{span[1]}')
        app.statusbar_element_reload()
        app.highlight()
        self.title(self.win_title + f' {num_of_results - results_index}/{num_of_results}件')

    def replace_button_clicked(self, mode=0):
        self.replace(mode)

    def replace(self, mode=0):
        '''
        mode: int[0,-1,1]

        mode=0: Replace one and move on to the next.
        mode=-1: Replace one and go back to previous.
        mode=1: Replace All
        '''
        if not self.results:
            self.make_results()
            return
        app.set_text_widget_editable(mode=1)
        if mode == 1:
            for result in self.results:
                line, row, span = result
                # 変更点を抜き出し、置換する
                pattern = self.text_in_entry.get()
                if not self.use_regular_expression.get():
                    pattern = re.escape(pattern)
                repl = self.text_in_entry2.get()
                text = app.maintexts[line].get(f'{row+1}.{span[0]} linestart', f'{row+1}.{span[1]} lineend')
                text = re.sub(pattern, repl, text, 1)
                # 変更点を削除し、新しいテキストを差し込む
                app.maintexts[line].delete(f'{row+1}.{span[0]} linestart', f'{row+1}.{span[1]} lineend')
                app.maintexts[line].insert(f'{row+1}.{span[0]} linestart', text)
            self.results.clear()
            self.title(self.win_title + ' 0件')
        else:
            result = self.results.popleft()
            if result == self.last_result:
                try:
                    self.last_result = self.results[-1]
                except IndexError:
                    self.last_result = None
            line, row, span = result
            row = row[0] + 1
            # 変更点を抜き出し、置換する
            pattern = self.text_in_entry.get()
            if not self.use_regular_expression.get():
                pattern = re.escape(pattern)
            repl = self.text_in_entry2.get()
            text = app.maintexts[line].get(f'{row}.{span[0]} linestart', f'{row}.{span[1]} lineend')
            text = re.sub(pattern, repl, text, 1)
            # 変更点を削除し、新しいテキストを差し込む
            app.maintexts[line].delete(f'{row}.{span[0]} linestart', f'{row}.{span[1]} lineend')
            app.maintexts[line].insert(f'{row}.{span[0]} linestart', text)
            self.select(mode)
        app.set_text_widget_editable(mode=2)

    def entry_return(self, shift:bool=False):
        if shift:
            self.prev_button_clicked()
        else:
            self.next_button_clicked()
        self.entry.focus()

    def entry2_return(self, shift:bool=False):
        if shift:
            self.replace(-1)
        else:
            self.replace()
        self.entry2.focus()

    def next_button_clicked(self):
        updated = self.make_results()
        if not updated:
            self.results.rotate(-1)
        self.select()
        self.next_button.focus()

    def prev_button_clicked(self):
        updated = self.make_results()
        if not updated:
            self.results.rotate(1)
        self.select()
        self.prev_button.focus()

    def switch_mode(self):
        self.destroy()
        if self.mode == 0:
            mode = '1'
        elif self.mode == 1:
            mode = '0'
        searchtext = self.text_in_entry.get()
        app.open_SearchWindow(mode, searchtext)

    def close(self):
        for w in app.maintexts:
            w.tag_remove('search', 1.0, END)
            w.tag_remove('search_selected', 1.0, END)
        app.close_sub_window(self)()

class TemplateWindow(Toplevel):
    def __init__(self, title="SoroEditor - 定型文", iconphoto='', size=(800, 750), position=None, minsize=None, maxsize=None, resizable=None, transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        self.close = app.close_sub_window(self)

        self.protocol('WM_DELETE_WINDOW', self.close)

        app.set_icon_sub_window(self)

        log.info('---Open TemplateWindow---')

        self.is_topmost = BooleanVar()

        self.menubar = Menu(self)
        self.menubar.add_checkbutton(label='最前面(T)', variable=self.is_topmost, underline=4)
        self.menubar.add_command(label='閉じる(Q)', command=self.destroy, underline=4)
        self.config(menu=self.menubar)

        f1 = Frame(self, padding=5)
        f1.pack(fill=BOTH, expand=True)

        Label(f1, text='定型文を10個まで登録できます。\n右のボタンをクリック、メニューバーから選択、またはショートカットキーで挿入できます。\
            \nショートカットキーはこのウィンドウを開かず使用できます。').pack()

        f2 = Frame(f1, padding=5)
        f2.pack(fill=BOTH, expand=True, anchor=CENTER)

        # チェックボタンを設定
        self.is_topmost.trace_add('write', self.on_is_topmost_changed)
        self.is_topmost.set(True)
        Checkbutton(f1, text='最前面[Ctrl+T]', variable=self.is_topmost, padding=10).pack(side=BOTTOM, anchor=W)

        f2_1 = Frame(f2, padding=5)
        f2_2 = Frame(f2, padding=5)
        f2_1.place(relx=0.0, rely=0.0, relheight=1.0, relwidth=0.5)
        f2_2.place(relx=0.5, rely=0.0, relheight=1.0, relwidth=0.5)

        rows = [Frame(f2_1, padding=3) if i < 5 else Frame(f2_2, padding=3) for i in range(10)]
        for w in rows:
            w.pack(fill=BOTH)

        buttons = [Button(rows[i], text=f' Alt+{i+1}\nCtrl+{i+1}')
                    if i != 9
                    else Button(rows[i], text=' Alt+0\nCtrl+0')
                    for i in range(10)]
        for i, w in enumerate(buttons):
            w.config(command=self.button_clicked(i))
            w.pack(fill=Y, side=RIGHT)

        self.load()
        self.texts = [ScrolledText(rows[i], height=4, autohide=True) for i in range(10)]
        for i, w in enumerate(self.texts):
            w.pack(fill=X, expand=True)
            w.insert(1.0, self.templates[i])

        # 変更検知
        self.bind('<KeyRelease>', self.key_released)

        self.bind('<Escape>', self.close)
        self.bind('<Control-t>', lambda _: self.is_topmost.set(not self.is_topmost.get()))

        # ウィンドウの設定
        self.focus()

    def button_clicked(self, num):
        def inner():
            app.insert_template_to_maintext(num)
        return inner

    def load(self):
        try:
            with open('./settings.yaml', 'rt', encoding='utf-8') as f:
                self.settings:dict = yaml.safe_load(f)
        except (FileNotFoundError, UnicodeDecodeError, yaml.YAMLError) as e:
            error_type = type(e).__name__
            error_message = str(e)
            log.error(f"An error of type {error_type} occurred while loading setting file: {error_message}")
            MessageDialog('設定ファイルの読み込みに失敗したため定型文を利用できません\nsettings.yamlが存在するか確認してください',
                            'SoroEditor - 定型文').show(app.md_position)
            self.destroy()
        else:
            self.templates = {i: self.settings['templates'][i]
                                if i in self.settings['templates'].keys()
                                else ''
                                for i in range(10)}

    def save(self):
        # 設定ファイルに更新された辞書を保存する
        try:
            app.settings['templates'] = self.templates
            app.templates = self.templates
            app.make_menu_templates()
            with open('./settings.yaml', mode='wt', encoding='utf-8') as f:
                yaml.dump(self.settings, f, allow_unicode=True)
        except (FileNotFoundError, UnicodeDecodeError, yaml.YAMLError) as e:
            error_type = type(e).__name__
            error_message = str(e)
            log.error(f"An error of type {error_type} occurred while saving setting file: {error_message}")
        else:
            log.info('Succeed updating setting file by TemplateWindow')

    def key_released(self, e=None):
        # Entry内のテキストに変更があるか確認する
        templates = self.get_current_data()
        if templates != self.templates:
            self.templates = templates
            self.settings['templates'] = self.templates
            self.save()

    def on_is_topmost_changed(self, *_):
        if self.is_topmost.get():
            self.attributes('-topmost', True)
            self.menubar.entryconfig(0, label='✓最前面(T)', underline=5)
        else:
            self.attributes('-topmost', False)
            self.menubar.entryconfig(0, label='　最前面(T)', underline=4)

    def get_current_data(self):
        return {i: re.sub('(\n|\r|\r\n)$', '', w.get(1.0, END), 1) for i, w in enumerate(self.texts)}


class BookmarkWindow(Toplevel):

    def __init__(self, title="SoroEditor - 付箋", iconphoto='', size=(900, 750), position=None, minsize=None, maxsize=None, resizable=None, transient=None, overrideredirect=False, windowtype=None, topmost=False, toolwindow=False, alpha=1, **kwargs):
        super().__init__(title, iconphoto, size, position, minsize, maxsize, resizable, transient, overrideredirect, windowtype, topmost, toolwindow, alpha, **kwargs)

        self.close = app.close_sub_window(self)

        self.protocol('WM_DELETE_WINDOW', self.close)

        app.set_icon_sub_window(self)

        log.info('---Open BookmarkWindow---')

        self.is_topmost = BooleanVar()

        self.menubar = Menu(self)
        self.menubar.add_checkbutton(label='最前面(T)', variable=self.is_topmost, underline=4)
        self.menubar.add_command(label='更新(R)', command=self.reset, underline=3)
        self.menubar.add_command(label='付箋設定(O)', command=self.bookmark_setting, underline=5)
        self.menubar.add_command(label='閉じる(Q)', command=self.close, underline=4)
        self.config(menu=self.menubar)

        self.is_topmost.trace_add('write', self.on_is_topmost_changed)
        self.is_topmost.set(True)

        bottombar = Frame(self, padding=10)
        bottombar.pack(fill=X, side=BOTTOM)
        Button(bottombar, text='更新[F5]', command=self.reset, takefocus=NO).pack(side=RIGHT, anchor=E, padx=10)
        Button(bottombar, text='付箋設定[Ctrl+P]', command=lambda: self.bookmark_setting(self), takefocus=NO).pack(side=RIGHT, anchor=E, padx=10)
        # チェックボタンを設定
        Checkbutton(bottombar, text='最前面[Ctrl+T]', variable=self.is_topmost).pack(side=LEFT, padx=10)

        f1 = Frame(self, padding=15)
        f1.pack(fill=BOTH, expand=True)

        # ツリービューの設定
        self.treeview = Treeview(f1, height=10, show=[HEADINGS], columns=['#data', 'mark', 'index', 'text'])
        self.treeview.column('#data',width=0, stretch='no')
        self.treeview.column('mark', anchor=W, width=150, stretch='no')
        self.treeview.column('index', anchor=W, width=100, stretch='no')
        self.treeview.column('text', anchor=W, width=500)
        self.treeview.heading('mark', text='付箋', anchor=W, command=lambda: self.change_column_width('mark'))
        self.treeview.heading('index', text='位置', anchor=W, command=lambda: self.change_column_width('index'))
        self.treeview.heading('text', text='テキスト', anchor=W, command=lambda: self.change_column_width('text'))
        self.insert_recode()

        # スクロールバーを設定
        vbar = Scrollbar(f1, command=self.treeview.yview, orient=VERTICAL, bootstyle='rounded')
        vbar.pack(side=RIGHT, fill=Y)
        self.treeview.configure(yscrollcommand=vbar.set)
        hbar = Scrollbar(f1, command=self.treeview.xview, orient=HORIZONTAL, bootstyle='rounded')
        hbar.pack(side=BOTTOM, fill=X)
        self.treeview.configure(xscrollcommand=hbar.set)
        self.treeview.pack(fill=BOTH, expand=True)

        self.treeview.bind('<<TreeviewSelect>>', self.select)
        self.bind('<Escape>', self.close)
        self.bind('<F5>', self.reset)
        self.bind('<Control-p>', self.bookmark_setting)
        self.bind('<Control-t>', lambda _: self.is_topmost.set(not self.is_topmost.get()))

        # ウィンドウの設定
        self.treeview.focus_set()

    def insert_recode(self):
        if not app.bookmarks:
            self.treeview.insert('', END, values=('', '付箋なし', ''), tags='font')
            return
        TreeviewValues = namedtuple('TreeviewValues', ['place', 'mark', 'index', 'text'])
        for mark in app.bookmarks:
            mark_index = app.search(mark, app.use_regex_in_bookmarks)
            if not mark_index:
                self.treeview.insert('', END, values=('', mark, '', '該当なし'), tags='font')
                continue
            for i, place in enumerate(mark_index):
                index = f'{place[0]+1}列{place[1][0]+1}行'
                texts = [widget.get(float(place[1][0]+1), str(float(place[1][0]+1))+' lineend')
                            for widget in app.maintexts]
                texts = ' | '.join(texts)
                new_values = TreeviewValues(place, mark, index, texts) if i == 0 else TreeviewValues(place, '', index, texts)
                self.treeview.insert('', END, values=new_values)

    def reset(self, *_):
        self.treeview.delete(*self.treeview.get_children())
        self.insert_recode()

    def select(self, *_):
        selected_item = self.treeview.selection()

        if not selected_item:
            return
        place = re.findall('\d+', self.treeview.item(selected_item)['values'][0])
        if not place:
            return

        maintexts_index = int(place[0])
        line = float(place[1])+1

        app.align_the_lines(line, line, False)

        self.attributes('-topmost', True)
        app.maintexts[maintexts_index].focus()
        app.maintexts[maintexts_index].mark_set(INSERT, line)
        app.highlight()
        self.attributes('-topmost', self.is_topmost.get())
        self.treeview.focus_set()

    def change_column_width(self, widget):
        widget_widthes = {
            'mark': [150, 50],
            'index': [100, 50],
            'text': [500, 50],
        }
        clicked_widget_widthes = widget_widthes[widget]
        current_width = self.treeview.column(widget, 'width')

        if current_width == clicked_widget_widthes[0]:
            new_width = clicked_widget_widthes[1]
        elif current_width == clicked_widget_widthes[1]:
            new_width = clicked_widget_widthes[0]
            if widget == 'text':
                self.treeview.column(widget, stretch=True)
        else:
            new_width = clicked_widget_widthes[1]
            self.treeview.column(widget, stretch=False)

        self.treeview.column(widget, width=new_width)

    def on_is_topmost_changed(self, *_):
        if self.is_topmost.get():
            self.attributes('-topmost', True)
            self.menubar.entryconfig(0, label='✓最前面(T)', underline=5)
        else:
            self.attributes('-topmost', False)
            self.menubar.entryconfig(0, label='　最前面(T)', underline=4)

    def bookmark_setting(self, *_):
        window = Toplevel(master=self)
        # ウィンドウの作成
        window.title('SoroEditor - 付箋')
        window.geometry('600x600')
        menubar = Menu(window)
        menubar.add_command(label='閉じる(Q)', command=window.destroy, underline=4)
        window.config(menu=menubar)
        f1 = Frame(window, padding=10)
        f1.pack(fill=BOTH, expand=True)
        Label(f1, text='この画面で設定したテキストは【付箋】として、付箋ウィンドウで確認できます。\n'\
            '付箋が貼られた部分には付箋ウィンドウやショートカットキーで簡単にジャンプできます。\n'\
                'ブックマークはファイルごとに設定されます。\n'\
                    '下のテキストボックスの1行ごとが付箋になります。', padding=10).pack()
        # チェックボタンを設定
        def on_use_regex_changed(*_):
            app.use_regex_in_bookmarks = use_regex.get()
        use_regex = BooleanVar(value=app.use_regex_in_bookmarks)
        use_regex.trace_add('write', on_use_regex_changed)
        use_regex.set(False)
        Checkbutton(f1, text='正規表現を用いる[Ctrl+R]', variable=use_regex, padding=10).pack(side=BOTTOM, anchor=W)
        # テキストボックスを設定
        text = ScrolledText(f1, font=app.font)
        bookmarks = '\n'.join(app.bookmarks)
        text.insert(END, bookmarks)
        text.pack(fill=BOTH, expand=True)
        textwidget:ScrolledText|Text = text
        for widget in text.winfo_children():
            if widget.winfo_class() == 'Text':
                textwidget = widget

        def save(_):
            new_bookmarks = [s for s in text.get(1.0, END).splitlines() if s]
            if set(new_bookmarks) != set(app.bookmarks) or app.use_regex_in_bookmarks == use_regex.get():
                app.bookmarks = new_bookmarks
                app.make_menu_bookmarks()
                app.use_regex_in_bookmarks = use_regex.get()

        window.bind('<KeyRelease>', save)
        window.bind('<Button>', save)
        window.bind('<Escape>', lambda _: window.destroy())
        window.bind('<Control-r>', lambda _: use_regex.set(not use_regex.get()))

        window.grab_set()
        textwidget.focus()
        if self.is_topmost.get():
            window.attributes('-topmost', True)
        self.wait_window(window)
        self.lift()
        self.focus()


class Icons:
    '''Stores the path of the icon image'''
    def __init__(self):
        try:
            with open('./settings.yaml', mode='rt', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                theme = data['themename']
                if not theme or not theme in STANDARD_THEMES.keys():
                    theme = DEFAULT_THEME
                theme_type = STANDARD_THEMES[theme]['type']
            if theme_type == 'dark':
                icon_type = 'white'
            elif theme_type == 'light':
                icon_type = 'black'
        except:
            theme = 'litera'
            theme_type = 'light'
            icon_type = 'black'
        self.icon = PhotoImage(file='src/icon/icon.png')
        self.file_create = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/file_create.png'))
        self.file_open = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/file_open.png'))
        self.file_save = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/file_save.png'))
        self.file_save_as = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/file_save_as.png'))
        self.refresh = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/refresh.png'))
        self.undo = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/undo.png'))
        self.repeat = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/repeat.png'))
        self.bookmark = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/bookmark.png'))
        self.template = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/template.png'))
        self.search = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/search.png'))
        self.replace = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/replace.png'))
        self.settings = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/settings.png'))
        self.project_settings = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/project_settings.png'))
        self.export = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/export.png'))
        self.import_ = PhotoImage(file=self.__make_image_path(f'src/icon/{icon_type}/import.png'))

    def __make_image_path(self, path) -> str:
        return os.path.join(os.path.dirname(__file__), path)

if __name__ == '__main__':
    log_setting()
    log.info('===Start Application===')
    root = Window(title='SoroEditor', minsize=(800, 500))
    root.iconphoto(False, Icons().icon)
    app = Main(master=root)
    app.mainloop()
    log.info('===Close Application===')
