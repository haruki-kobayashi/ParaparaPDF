import os
import glob
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import io

# PyMuPDFとPILが利用可能かチェック
try:
    import fitz  # PyMuPDF
    from PIL import Image, ImageTk
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

class PDFViewerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ParaparaPDF")
        self.root.geometry("1310x752")  # 横A4サイズの読み込みに合わせたウィンドウサイズ
        self.root.configure(bg='SlateGray1')  # メインウィンドウの背景色

        self.pdf_files = []
        self.current_index = 0
        self.current_dir = os.getcwd()
        self.pdf_document = None
        self.current_pdf_path = None  # 現在開いているPDFのパス
        self.current_page = 0
        self.total_pages = 0
        self.zoom_level = 1.0

        # ページ範囲管理用変数
        self.allowed_pages = None  # None = 全ページ, リスト = 指定ページのみ
        self.filtered_pages = []   # 表示可能なページリスト

        # ドラッグ操作用変数
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.is_dragging = False

        # ファイル位置表示用変数
        self.file_position_var = tk.StringVar(value="")

        # ファイルごとの最後に見ていたページを記憶する辞書
        self.last_viewed_pages = {}

        # レンダリング品質設定
        self.quality_modes = [
            {"name": "Vector", "desc": "ベクターレンダリング（高品質）"},
            {"name": "Bitmap", "desc": "ビットマップレンダリング（高速）"}
        ]
        self.current_quality_index = 1  # デフォルトはBitmap

        self.setup_ui()

    def setup_ui(self):
        """UIをセットアップ"""
        # メイリオフォントを定義
        meiryo_frame_font = ("Meiryo", 10)
        meiryo_font = ("Meiryo", 9)

        # スタイルを設定
        style = ttk.Style()
        style.configure('White.TSeparator', background='white')

        # メインフレーム
        main_frame = tk.Frame(self.root, bg='SlateGray1', padx=5, pady=5)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))        # リサイズ可能なパネル（PanedWindow）
        self.paned_window = tk.PanedWindow(main_frame, orient=tk.HORIZONTAL, sashwidth=7, bg='SlateGray1')
        self.paned_window.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 左側パネル（ファイル操作）
        left_panel = tk.Frame(self.paned_window, width=320, bg='SlateGray1')

        # 検索フレーム
        search_frame = tk.LabelFrame(left_panel, text="PDFファイル検索", padx=5, pady=5, bg='SlateGray1', bd=1, relief='solid', font=meiryo_frame_font)
        search_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        # ワイルドカードのヘルプ
        dir_help_text = "* : 任意の文字列   ? : 任意の1文字"
        tk.Label(search_frame, text=dir_help_text, bg='SlateGray1', font=meiryo_font).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        # ディレクトリ選択
        tk.Label(search_frame, text="ディレクトリパターン:", bg='SlateGray1', font=meiryo_font).grid(row=1, column=0, sticky=tk.W)
        self.dir_var = tk.StringVar(value=self.current_dir)
        self.dir_entry = ttk.Entry(search_frame, textvariable=self.dir_var, width=25)
        self.dir_entry.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(2, 2))
        self.dir_entry.bind('<Return>', lambda e: self.search_pdfs())
        tk.Button(search_frame, text="参照", command=self.select_directory, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=2, column=1, padx=(5, 0), pady=(2, 2))

        # パターン入力
        tk.Label(search_frame, text="ファイル名パターン:", bg='SlateGray1', font=meiryo_font).grid(row=3, column=0, sticky=tk.W)
        self.pattern_var = tk.StringVar(value="*")
        pattern_frame = tk.Frame(search_frame, bg='SlateGray1')
        pattern_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(2, 0))
        self.pattern_entry = ttk.Entry(pattern_frame, textvariable=self.pattern_var, width=15)
        self.pattern_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.pattern_entry.bind('<Return>', lambda e: self.search_pdfs())
        tk.Button(pattern_frame, text="検索", command=self.search_pdfs, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=1, padx=(5, 0))
        pattern_frame.columnconfigure(0, weight=1)

        # ページ範囲指定
        tk.Label(search_frame, text="表示ページ範囲:", bg='SlateGray1', font=meiryo_font).grid(row=5, column=0, sticky=tk.W, pady=(5, 0))
        page_range_frame = tk.Frame(search_frame, bg='SlateGray1')
        page_range_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(2, 0))

        self.page_range_var = tk.StringVar(value="*")
        self.page_range_entry = ttk.Entry(page_range_frame, textvariable=self.page_range_var, width=15)
        self.page_range_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.page_range_entry.bind('<Return>', lambda e: (self.apply_page_range(), self.search_pdfs(), self.pdf_canvas.focus_set()))

        tk.Button(page_range_frame, text="適用", command=lambda: (self.apply_page_range(), self.search_pdfs()), bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=1, padx=(5, 0))
        page_range_frame.columnconfigure(0, weight=1)

        # ヘルプテキスト
        help_text = "例: *, 1-3, 5, -8, 10-, even, odd"
        tk.Label(search_frame, text=help_text, bg='SlateGray1', font=meiryo_font).grid(row=7, column=0, sticky=tk.W)

        # ファイルリストフレーム
        list_frame = tk.LabelFrame(left_panel, text="検索結果一覧", padx=5, pady=5, bg='SlateGray1', bd=1, relief='solid', font=meiryo_frame_font)
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 5))

        # ファイルリスト
        self.file_listbox = tk.Listbox(list_frame, height=15, width=40)
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        self.file_listbox.bind('<Button-1>', self.on_file_click)  # マウスクリック時
        self.file_listbox.bind('<Button-3>', self.on_file_right_click)  # 右クリック時

        # listboxのキーボード操作を無効化（PDFビューアでの操作を優先）
        self.file_listbox.bind('<Left>', self.handle_left_key)
        self.file_listbox.bind('<Right>', self.handle_right_key)
        self.file_listbox.bind('<Up>', self.handle_up_key)
        self.file_listbox.bind('<Down>', self.handle_down_key)
        self.file_listbox.bind('<Prior>', lambda e: self.prev_page())  # Page Up
        self.file_listbox.bind('<Next>', lambda e: self.next_page())   # Page Down

        # スクロールバー（ファイルリスト用）
        list_v_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        list_v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        list_h_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.file_listbox.xview)
        list_h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        self.file_listbox.configure(yscrollcommand=list_v_scrollbar.set, xscrollcommand=list_h_scrollbar.set)

        # 操作フレーム
        control_frame = tk.LabelFrame(left_panel, text="表示ファイル切替", padx=5, pady=5, bg='SlateGray1', bd=1, relief='solid', font=meiryo_frame_font)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))

        # ナビゲーションボタン
        nav_frame = tk.Frame(control_frame, bg='SlateGray1')
        nav_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

        tk.Button(nav_frame, text="◀ 前", command=self.prev_pdf, width=8, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=0, padx=(0, 0))

        # ファイル位置表示
        tk.Label(nav_frame, textvariable=self.file_position_var, bg='SlateGray1', font=meiryo_font, width=8, anchor='center').grid(row=0, column=1, padx=(0, 0))

        tk.Button(nav_frame, text="次 ▶", command=self.next_pdf, width=8, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=2, padx=(0, 0))
        nav_frame.columnconfigure(0, weight=1)
        nav_frame.columnconfigure(1, weight=0)  # 中央の表示は固定幅
        nav_frame.columnconfigure(2, weight=1)

        # 現在のファイル情報
        self.info_var = tk.StringVar(value="PDFファイルを検索してください")
        self.info_label = tk.Label(control_frame, textvariable=self.info_var, wraplength=250, height=2, bg='SlateGray1', font=meiryo_font)
        self.info_label.grid(row=1, column=0, pady=(5, 0))

        # 右側パネル（ビューア）
        right_panel = tk.Frame(self.paned_window, bg='SlateGray1')

        # PanedWindowに左右のパネルを追加
        self.paned_window.add(left_panel, minsize=250, width=280)  # 最小サイズ250px、初期幅280px
        self.paned_window.add(right_panel, minsize=400)  # 最小サイズ400px

        # ビューアフレーム
        pdf_frame = tk.LabelFrame(right_panel, text="ビューア", padx=5, pady=5, bg='SlateGray1', bd=1, relief='solid', font=meiryo_frame_font)
        pdf_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # PDF操作ツールバー
        toolbar_frame = tk.Frame(pdf_frame, bg='SlateGray1')
        toolbar_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        tk.Button(toolbar_frame, text="▲", command=self.prev_page, width=3, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=0)
        self.page_var = tk.StringVar(value="0 / 0")
        tk.Label(toolbar_frame, textvariable=self.page_var, bg='SlateGray1', font=meiryo_font).grid(row=0, column=1, padx=(5, 5))
        tk.Button(toolbar_frame, text="▼", command=self.next_page, width=3, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=2)

        ttk.Separator(toolbar_frame, orient=tk.VERTICAL, style='White.TSeparator').grid(row=0, column=3, sticky=(tk.N, tk.S), padx=(10, 10))

        tk.Button(toolbar_frame, text="−", command=self.zoom_out, width=3, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=4)
        self.zoom_var = tk.StringVar(value="100%")
        zoom_label = tk.Label(toolbar_frame, textvariable=self.zoom_var, width=6, anchor='center', bg='SlateGray1', font=meiryo_font)
        zoom_label.grid(row=0, column=5, padx=(5, 5))
        tk.Button(toolbar_frame, text="＋", command=self.zoom_in, width=3, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=6)

        tk.Button(toolbar_frame, text="⤢", command=self.fit_to_window, width=3, bg='white', font=meiryo_font, relief='solid', bd=1).grid(row=0, column=7, padx=(10, 0))

        ttk.Separator(toolbar_frame, orient=tk.VERTICAL, style='White.TSeparator').grid(row=0, column=8, sticky=(tk.N, tk.S), padx=(10, 10))

        # 品質モード切り替えボタン
        self.quality_toggle_button = tk.Button(toolbar_frame, text="⇄", command=self.toggle_quality_mode, width=3, bg='white', font=meiryo_font, relief='solid', bd=1)
        self.quality_toggle_button.grid(row=0, column=9)

        # 現在の品質モード表示
        self.quality_var = tk.StringVar(value=self.quality_modes[self.current_quality_index]["name"])
        quality_mode_label = tk.Label(toolbar_frame, textvariable=self.quality_var, width=6, bg='SlateGray1', font=meiryo_font)
        quality_mode_label.grid(row=0, column=10, padx=(5, 0))

        ttk.Separator(toolbar_frame, orient=tk.VERTICAL, style='White.TSeparator').grid(row=0, column=11, sticky=(tk.N, tk.S), padx=(10, 10))

        # 既定のアプリで開くボタン
        self.open_external_button = tk.Button(toolbar_frame, text="既定のアプリで開く", command=self.open_with_default_app, width=16, bg='white', font=meiryo_font, relief='solid', bd=1)
        self.open_external_button.grid(row=0, column=12)

        ttk.Separator(toolbar_frame, orient=tk.VERTICAL, style='White.TSeparator').grid(row=0, column=13, sticky=(tk.N, tk.S), padx=(10, 10))

        # ヘルプボタン
        self.help_button = tk.Button(toolbar_frame, text="ヘルプ", command=self.show_help, width=6, bg='white', font=meiryo_font, relief='solid', bd=1)
        self.help_button.grid(row=0, column=14)

        # PDFキャンバス（スクロール可能）
        canvas_frame = tk.Frame(pdf_frame, bg='SlateGray1')
        canvas_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.pdf_canvas = tk.Canvas(canvas_frame, bg="#FFFFFF")
        self.pdf_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # スクロールバー（PDF用）
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.pdf_canvas.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.pdf_canvas.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        self.pdf_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # マウスドラッグイベントをPDFキャンバスに追加
        self.pdf_canvas.bind("<Button-1>", self.start_drag)
        self.pdf_canvas.bind("<B1-Motion>", self.on_drag)
        self.pdf_canvas.bind("<ButtonRelease-1>", self.end_drag)

        # マウスホイールイベントをPDFキャンバスに追加（通常のスクロール/ページ遷移）
        self.pdf_canvas.bind("<Button-4>", self.on_mousewheel)  # Linux/Unix
        self.pdf_canvas.bind("<Button-5>", self.on_mousewheel)  # Linux/Unix
        self.pdf_canvas.bind("<MouseWheel>", self.on_mousewheel)  # Windows

        # グリッドの重み設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)  # PanedWindowが伸縮
        main_frame.rowconfigure(0, weight=1)

        # PanedWindowの重み設定
        self.paned_window.rowconfigure(0, weight=1)

        # 左側パネルの設定
        left_panel.rowconfigure(1, weight=1)
        left_panel.columnconfigure(0, weight=1)

        # list_frameは通常の表示設定
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)  # 縦方向のみ伸縮
        search_frame.columnconfigure(0, weight=1)
        control_frame.columnconfigure(0, weight=1)

        # 右側パネルの設定
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(0, weight=1)
        pdf_frame.columnconfigure(0, weight=1)
        pdf_frame.rowconfigure(1, weight=1)
        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)

        # キーボードショートカット（グローバル）
        self.root.bind_all('<Left>', self.handle_left_key)
        self.root.bind_all('<Right>', self.handle_right_key)
        self.root.bind_all('<Up>', self.handle_up_key)
        self.root.bind_all('<Down>', self.handle_down_key)
        self.root.bind_all('<Prior>', lambda e: self.prev_page())  # Page Up
        self.root.bind_all('<Next>', lambda e: self.next_page())   # Page Down
        self.root.bind_all('<Home>', self.handle_home_key)         # Home - 最初のページ
        self.root.bind_all('<End>', self.handle_end_key)           # End - 最後のページ
        self.root.bind_all('<Control-plus>', lambda e: self.zoom_in())
        self.root.bind_all('<Control-minus>', lambda e: self.zoom_out())
        self.root.bind_all('<Control-Key-0>', lambda e: self.reset_zoom())  # Ctrl+0でズームリセット
        self.root.bind_all('<Control-f>', lambda e: self.fit_to_window())  # Ctrl+Fで画面サイズに自動フィット

        # マウスホイールイベント（Ctrl+ホイールでズーム）
        self.root.bind_all('<Control-Button-4>', lambda e: (self.zoom_in(), "break")[1])  # Linux/Unix
        self.root.bind_all('<Control-Button-5>', lambda e: (self.zoom_out(), "break")[1])  # Linux/Unix
        self.root.bind_all('<Control-MouseWheel>', self.on_ctrl_mousewheel)  # Windows

        # マウスホイールイベント（Shift+ホイールでファイル切り替え）
        self.root.bind_all('<Shift-Button-4>', lambda e: (self.prev_pdf(), "break")[1])  # Linux/Unix
        self.root.bind_all('<Shift-Button-5>', lambda e: (self.next_pdf(), "break")[1])  # Linux/Unix
        self.root.bind_all('<Shift-MouseWheel>', self.on_shift_mousewheel)  # Windows

        # フォーカス設定
        self.root.focus_set()

        # ファイルリスト用コンテキストメニューを作成
        self.create_context_menu()

    def create_context_menu(self):
        """ファイルリスト用のコンテキストメニューを作成"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="開く", command=self.context_open_file)
        self.context_menu.add_command(label="既定のアプリで開く", command=self.context_open_with_default)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="エクスプローラーで表示", command=self.context_show_in_explorer)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="パスをコピー", command=self.context_copy_path)
        self.context_menu.add_command(label="ファイル名をコピー", command=self.context_copy_filename)

    def natural_sort_key(self, text):
        """ナチュラルソート用のキーを生成（数値部分を数値として扱う）"""
        def convert(part):
            if part.isdigit():
                return int(part)
            return part.lower()

        return [convert(c) for c in re.split('([0-9]+)', text)]

    def select_directory(self):
        """ディレクトリを選択"""
        directory = filedialog.askdirectory(initialdir=self.current_dir)
        if directory:
            self.dir_var.set(directory)
            self.current_dir = directory

    def find_pdf_files(self, directory, pattern):
        """指定されたディレクトリとその下層から、パターンにマッチするPDFファイルを検索"""
        pdf_files = []

        # パターンが拡張子を含まない場合は.pdfを追加
        if not pattern.lower().endswith('.pdf'):
            pattern += '.pdf'

        # カレントディレクトリとその下層を再帰的に検索
        search_pattern = os.path.join(directory, '**', pattern)
        pdf_files = glob.glob(search_pattern, recursive=True)

        # ナチュラルソート（数値を考慮したソート）
        pdf_files.sort(key=self.natural_sort_key)

        return pdf_files

    def search_pdfs(self):
        """PDFファイルを検索"""
        directory_pattern = self.dir_var.get()
        pattern = self.pattern_var.get()

        # ワイルドカードが含まれているかチェック
        if '*' in directory_pattern or '?' in directory_pattern:
            # ワイルドカードを使用したディレクトリ検索
            directories = glob.glob(directory_pattern)
            if not directories:
                messagebox.showerror("エラー", f"パターン'{directory_pattern}'にマッチするディレクトリが見つかりません")
                return

            # 見つかったディレクトリが実際に存在するかチェック
            directories = [d for d in directories if os.path.isdir(d)]
            if not directories:
                messagebox.showerror("エラー", f"パターン'{directory_pattern}'にマッチする有効なディレクトリが見つかりません")
                return

            # 複数ディレクトリからPDFファイルを収集
            all_pdf_files = []
            for directory in directories:
                pdf_files = self.find_pdf_files(directory, pattern)
                all_pdf_files.extend(pdf_files)

            # 重複を除去してソート
            self.pdf_files = list(set(all_pdf_files))
            self.pdf_files.sort(key=self.natural_sort_key)

            # 基準ディレクトリを最初に見つかったディレクトリに設定
            self.current_dir = directories[0] if directories else os.getcwd()

        else:
            # 単一ディレクトリの通常検索
            directory = directory_pattern
            if not os.path.exists(directory):
                messagebox.showerror("エラー", "指定されたディレクトリが存在しません")
                return

            self.pdf_files = self.find_pdf_files(directory, pattern)
            self.current_dir = directory

        self.current_index = 0

        # リストボックスを更新
        self.file_listbox.delete(0, tk.END)
        if self.pdf_files:
            for pdf_file in self.pdf_files:
                # ワイルドカード検索時は絶対パスを表示、通常検索時は相対パスを表示
                directory_pattern = self.dir_var.get()
                if '*' in directory_pattern or '?' in directory_pattern:
                    # ワイルドカード検索時は絶対パスを表示
                    display_path = pdf_file
                else:
                    # 通常検索時は相対パスを表示
                    display_path = os.path.relpath(pdf_file, self.current_dir)
                self.file_listbox.insert(tk.END, display_path)
            self.file_listbox.selection_set(0)
            self.update_info()
            # 検索後に最初のファイルを自動的に開く
            self.open_current_pdf()
            # PDFビューアにフォーカスを移す
            self.pdf_canvas.focus_set()
        else:
            self.info_var.set(f"パターン'{pattern}'にマッチするPDFファイルが見つかりませんでした")

    def apply_page_range(self):
        """ページ範囲を適用"""
        range_text = self.page_range_var.get().strip()

        if range_text.lower() in ['*', '全て', 'すべて', 'all', '']:
            self.allowed_pages = None
        else:
            try:
                self.allowed_pages = self.parse_page_range(range_text)
            except ValueError as e:
                messagebox.showerror("エラー", f"ページ範囲の形式が正しくありません: {e}")
                return

        # 現在のPDFファイルがある場合は再表示
        if self.pdf_document:
            self.update_filtered_pages()
            if self.filtered_pages:
                self.current_page = self.filtered_pages[0]  # フィルタ後の最初のページに移動
            self.display_page()

    def parse_page_range(self, range_text):
        """ページ範囲をパース"""
        pages = set()

        # '*'は全ページを意味
        if range_text.strip() == '*':
            return None

        parts = [p.strip() for p in range_text.split(',')]

        for part in parts:
            if not part:
                continue

            if '-' in part:
                # ハイフンの位置を確認
                if part.startswith('-'):
                    # "-10" 形式: 1からN
                    try:
                        end = int(part[1:])
                        if end < 1:
                            raise ValueError(f"無効なページ番号: {end}")
                        for i in range(1, end + 1):
                            pages.add(i)
                    except ValueError:
                        raise ValueError(f"無効な範囲形式: {part}")
                elif part.endswith('-'):
                    # "3-" 形式: Nから最後まで（特別なマーカーを使用）
                    try:
                        start = int(part[:-1])
                        if start < 1:
                            raise ValueError(f"無効なページ番号: {start}")
                        # 特別なマーカー値を追加（負の値を使用してopen-endedを示す）
                        pages.add(-start)  # 負の値でopen-endedを表現
                    except ValueError:
                        raise ValueError(f"無効な範囲形式: {part}")
                else:
                    # "1-5" 形式: 通常の範囲
                    try:
                        start, end = map(int, part.split('-'))
                        if start < 1 or end < 1 or start > end:
                            raise ValueError(f"無効な範囲: {start}-{end}")
                        for i in range(start, end + 1):
                            pages.add(i)
                    except ValueError:
                        raise ValueError(f"無効な範囲形式: {part}")
            else:
                # 単一ページまたは特殊キーワード
                if part.lower() in ['even', '偶数']:
                    # 偶数ページ（特別なマーカーを使用）
                    pages.add(-9998)  # 特別なマーカー値
                elif part.lower() in ['odd', '奇数']:
                    # 奇数ページ（特別なマーカーを使用）
                    pages.add(-9999)  # 特別なマーカー値
                else:
                    # 通常の単一ページ
                    try:
                        page = int(part)
                        if page < 1:
                            raise ValueError(f"無効なページ番号: {page}")
                        pages.add(page)
                    except ValueError:
                        raise ValueError(f"無効なページ番号: {part}")

        return sorted(list(pages))

    def update_filtered_pages(self):
        """表示可能なページリストを更新"""
        if not self.pdf_document:
            self.filtered_pages = []
            return

        if self.allowed_pages is None:
            # 全ページ表示
            self.filtered_pages = list(range(self.total_pages))
        else:
            # 指定ページのみ（1ベースから0ベースに変換し、範囲チェック）
            self.filtered_pages = []
            actual_pages = set()

            for page_marker in self.allowed_pages:
                if page_marker == -9999:  # 奇数ページ
                    for i in range(1, self.total_pages + 1, 2):  # 1, 3, 5, ...
                        actual_pages.add(i)
                elif page_marker == -9998:  # 偶数ページ
                    for i in range(2, self.total_pages + 1, 2):  # 2, 4, 6, ...
                        actual_pages.add(i)
                elif page_marker < 0 and page_marker > -9998:  # open-ended range (3-)
                    start = -page_marker
                    for i in range(start, self.total_pages + 1):
                        actual_pages.add(i)
                else:  # 通常のページ番号
                    if page_marker > 0:
                        actual_pages.add(page_marker)

            # 1ベースから0ベースに変換し、範囲チェック
            for page_num in sorted(actual_pages):
                zero_based = page_num - 1  # 1ベースから0ベースに変換
                if 0 <= zero_based < self.total_pages:
                    self.filtered_pages.append(zero_based)

    def on_file_select(self, event):
        """リストボックスでファイルが選択された時"""
        selection = self.file_listbox.curselection()
        if selection:
            self.current_index = selection[0]
            self.update_info()
            self.open_current_pdf()
            # PDFビューアにフォーカスを移す
            self.pdf_canvas.focus_set()

    def on_file_click(self, event):
        """リストボックスがクリックされた時"""
        # わずかな遅延をおいてフォーカスを移す（選択処理完了後）
        self.root.after(10, lambda: self.pdf_canvas.focus_set())

    def on_file_right_click(self, event):
        """リストボックスで右クリックされた時"""
        # 右クリックされた位置のアイテムを選択
        index = self.file_listbox.nearest(event.y)
        if 0 <= index < self.file_listbox.size():
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(index)
            self.current_index = index

            # コンテキストメニューを表示
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()

    def context_open_file(self):
        """コンテキストメニュー：ファイルを開く"""
        if self.pdf_files and 0 <= self.current_index < len(self.pdf_files):
            self.update_info()
            self.open_current_pdf()

    def context_open_with_default(self):
        """コンテキストメニュー：既定のアプリで開く"""
        self.open_with_default_app()

    def context_show_in_explorer(self):
        """コンテキストメニュー：エクスプローラーで表示"""
        if not self.pdf_files or self.current_index >= len(self.pdf_files):
            return

        try:
            file_path = self.pdf_files[self.current_index]
            # Windowsのエクスプローラーでファイルを選択状態で表示
            if os.name == 'nt':
                import subprocess
                subprocess.run(['explorer', '/select,', file_path])
            else:
                # Linux/macOSの場合はフォルダを開く
                folder_path = os.path.dirname(file_path)
                if os.name == 'posix' and os.uname().sysname == 'Darwin':
                    os.system(f'open "{folder_path}"')
                else:
                    os.system(f'xdg-open "{folder_path}"')
        except Exception as e:
            messagebox.showerror("エラー", f"エクスプローラーでの表示に失敗しました: {e}")

    def context_copy_path(self):
        """コンテキストメニュー：パスをコピー"""
        if not self.pdf_files or self.current_index >= len(self.pdf_files):
            return

        try:
            file_path = self.pdf_files[self.current_index]
            self.root.clipboard_clear()
            self.root.clipboard_append(file_path)
        except Exception as e:
            messagebox.showerror("エラー", f"パスのコピーに失敗しました: {e}")

    def context_copy_filename(self):
        """コンテキストメニュー：ファイル名をコピー"""
        if not self.pdf_files or self.current_index >= len(self.pdf_files):
            return

        try:
            file_path = self.pdf_files[self.current_index]
            filename = os.path.basename(file_path)
            self.root.clipboard_clear()
            self.root.clipboard_append(filename)
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル名のコピーに失敗しました: {e}")

    def load_pdf(self, pdf_path):
        """PDFファイルを読み込む"""
        if not PDF_SUPPORT:
            messagebox.showerror("エラー",
                               "PDFライブラリが利用できません。\n"
                               "pip install PyMuPDF Pillow でインストールしてください。")
            return False

        try:
            # 現在開いているファイルがある場合、そのファイルの最後に見ていたページを保存
            if self.pdf_document and hasattr(self, 'current_pdf_path'):
                self.last_viewed_pages[self.current_pdf_path] = self.current_page

            if self.pdf_document:
                self.pdf_document.close()

            self.pdf_document = fitz.open(pdf_path)
            self.current_pdf_path = pdf_path  # 現在のファイルパスを保存
            self.total_pages = len(self.pdf_document)
            # ページ範囲を更新
            self.update_filtered_pages()

            # 保存されたページがあれば復元、なければフィルタ後の最初のページに移動
            if pdf_path in self.last_viewed_pages:
                # 保存されたページがフィルタされたページ範囲内にあるかチェック
                saved_page = self.last_viewed_pages[pdf_path]
                if saved_page in self.filtered_pages:
                    self.current_page = saved_page
                else:
                    # フィルタ範囲外なら最初のページに
                    self.current_page = self.filtered_pages[0] if self.filtered_pages else 0
            else:
                # フィルタ後の最初のページに移動
                if self.filtered_pages:
                    self.current_page = self.filtered_pages[0]
                else:
                    self.current_page = 0

            # 最初のページでビューアに合わせてズームレベルを計算
            self.auto_fit_to_width()
            self.display_page()
            return True
        except Exception as e:
            messagebox.showerror("エラー", f"PDFファイルを読み込めませんでした: {e}")
            return False

    def auto_fit_to_width(self):
        """PDFページをビューアの幅に合わせて自動調整"""
        if not self.pdf_document or self.total_pages == 0:
            return

        try:
            # UIの更新を確実に実行
            self.root.update_idletasks()

            # キャンバスの実際のサイズを取得を遅延実行
            self.root.after(100, self._delayed_auto_fit)

        except Exception:
            self.zoom_level = 1.0

    def _delayed_auto_fit(self):
        """遅延実行でのフィット処理"""
        try:
            # キャンバスサイズを直接取得
            canvas_width = self.pdf_canvas.winfo_width()
            canvas_height = self.pdf_canvas.winfo_height()

            # キャンバスサイズが正常に取得できない場合のフォールバック
            if canvas_width <= 1 or canvas_height <= 1:
                # より積極的にサイズを取得
                self.root.update_idletasks()
                canvas_width = self.pdf_canvas.winfo_width()
                canvas_height = self.pdf_canvas.winfo_height()

                # それでもダメな場合は親コンテナから推定
                if canvas_width <= 1 or canvas_height <= 1:
                    total_width = self.root.winfo_width()
                    total_height = self.root.winfo_height()

                    # 現在の分割位置を考慮して計算
                    try:
                        sash_coord = self.paned_window.sash_coord(0)
                        if sash_coord:
                            left_width = sash_coord[0]
                        else:
                            left_width = 320  # デフォルト
                    except Exception:
                        left_width = 320

                    # より大きなサイズを使用（余裕を持たせる）
                    canvas_width = total_width - left_width - 20  # 最小限の余白のみ
                    canvas_height = total_height - 80  # ツールバー分のみ

            # 現在のページを取得
            page = self.pdf_document[self.current_page]
            rect = page.rect
            page_width = rect.width
            page_height = rect.height

            # 余白を最小限にして利用可能領域を最大化
            margin = 2  # 最小限の余白
            usable_width = canvas_width - margin
            usable_height = canvas_height - margin

            # ズーム比率を計算
            width_ratio = usable_width / page_width
            height_ratio = usable_height / page_height

            # ズームレベルを設定
            self.zoom_level = min(width_ratio, height_ratio) * 0.9965

            # ズームレベルの範囲制限
            self.zoom_level = max(0.1, min(self.zoom_level, 8.0))

            # ページを再描画
            self.display_page()

        except Exception:
            self.zoom_level = 1.0
            self.display_page()

    def display_page(self):
        """現在のページを表示"""
        if not PDF_SUPPORT or not self.pdf_document or self.current_page >= self.total_pages:
            return

        try:
            # ページを取得
            page = self.pdf_document[self.current_page]

            # ズームを適用したマトリックス
            mat = fitz.Matrix(self.zoom_level, self.zoom_level)

            # 現在の品質モードを取得
            current_mode = self.quality_modes[self.current_quality_index]
            is_vector_mode = current_mode["name"] == "Vector"

            if is_vector_mode:
                # 軽量ベクターレンダリング：高品質設定でビットマップ化
                try:
                    # ズームレベルに応じてベクタースケールを動的調整
                    # 高ズーム時はスケールを抑制してパフォーマンス向上
                    if self.zoom_level <= 2.0:
                        vector_scale = 1.5  # 200%まで：1.5倍
                    elif self.zoom_level <= 5.0:
                        vector_scale = 1.3  # 500%まで：1.3倍
                    else:
                        vector_scale = 1.1  # 800%まで：1.1倍（軽量化）

                    vector_matrix = fitz.Matrix(self.zoom_level * vector_scale, self.zoom_level * vector_scale)

                    pix = page.get_pixmap(
                        matrix=vector_matrix,
                        alpha=False,
                        annots=True,
                        clip=None,
                        colorspace=fitz.csRGB
                    )

                    # PIL Imageに変換
                    img_data = pix.tobytes("ppm")
                    pil_image = Image.open(io.BytesIO(img_data))

                    # 軽量ダウンサンプリング
                    target_size = (int(pil_image.width / vector_scale), int(pil_image.height / vector_scale))
                    pil_image = pil_image.resize(target_size, Image.Resampling.LANCZOS)

                    # Tkinter PhotoImageに変換
                    self.photo = ImageTk.PhotoImage(pil_image)

                    # キャンバスをクリア
                    self.pdf_canvas.delete("all")

                    # 画像を表示
                    self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)

                    # スクロール領域を設定
                    self.pdf_canvas.configure(scrollregion=self.pdf_canvas.bbox("all"))

                except Exception:
                    # フォールバック：標準ビットマップレンダリング
                    pix = page.get_pixmap(matrix=mat, alpha=False, annots=True, clip=None, colorspace=fitz.csRGB)
                    img_data = pix.tobytes("ppm")
                    pil_image = Image.open(io.BytesIO(img_data))
                    self.photo = ImageTk.PhotoImage(pil_image)
                    self.pdf_canvas.delete("all")
                    self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
                    self.pdf_canvas.configure(scrollregion=self.pdf_canvas.bbox("all"))
            else:
                # 標準ビットマップレンダリング（シンプル・高速）
                pix = page.get_pixmap(
                    matrix=mat,
                    alpha=False,
                    annots=True,
                    clip=None,
                    colorspace=fitz.csRGB
                )

                # PIL Imageに変換
                img_data = pix.tobytes("ppm")
                pil_image = Image.open(io.BytesIO(img_data))

                # Tkinter PhotoImageに変換
                self.photo = ImageTk.PhotoImage(pil_image)

                # キャンバスをクリア
                self.pdf_canvas.delete("all")

                # 画像を表示
                self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)

                # スクロール領域を設定
                self.pdf_canvas.configure(scrollregion=self.pdf_canvas.bbox("all"))

            # ドラッグスクロールのための初期設定（ベクターモードでないときのみ）
            if not is_vector_mode:
                self.pdf_canvas.scan_mark(0, 0)

            # ページ情報を更新
            if self.filtered_pages:
                current_index = self.get_current_filtered_index()
                display_page = self.current_page + 1  # 1ベース表示
                total_filtered = len(self.filtered_pages)
                self.page_var.set(f"{display_page} ({current_index + 1}/{total_filtered})")
            else:
                self.page_var.set("0 / 0")
            self.zoom_var.set(f"{int(self.zoom_level * 100)}%")

        except Exception as e:
            messagebox.showerror("エラー", f"ページの表示に失敗しました: {e}")

    def prev_page(self):
        """前のページに移動"""
        if not self.pdf_document or not self.filtered_pages:
            return

        current_index = self.get_current_filtered_index()
        if current_index > 0:
            self.current_page = self.filtered_pages[current_index - 1]
            self.display_page()

    def next_page(self):
        """次のページに移動"""
        if not self.pdf_document or not self.filtered_pages:
            return

        current_index = self.get_current_filtered_index()
        if current_index < len(self.filtered_pages) - 1:
            self.current_page = self.filtered_pages[current_index + 1]
            self.display_page()

    def get_current_filtered_index(self):
        """現在のページがフィルタされたページリスト内のどのインデックスかを取得"""
        try:
            return self.filtered_pages.index(self.current_page)
        except ValueError:
            # 現在のページがフィルタされたリストにない場合は最初のページに設定
            if self.filtered_pages:
                self.current_page = self.filtered_pages[0]
                return 0
            return 0

    def zoom_in(self):
        """ズームイン"""
        self.zoom_level = min(self.zoom_level * 1.2, 8.0)
        self.display_page()

    def zoom_out(self):
        """ズームアウト"""
        self.zoom_level = max(self.zoom_level / 1.2, 0.2)
        self.display_page()

    def reset_zoom(self):
        """ズームをリセット（100%に戻す）"""
        self.zoom_level = 1.0
        self.display_page()

    def on_ctrl_mousewheel(self, event):
        """Ctrl+マウスホイールでズーム（スクロールを防ぐ）"""
        if event.delta > 0:  # ホイールアップ
            self.zoom_in()
        else:  # ホイールダウン
            self.zoom_out()
        # returnでイベントの伝播を防ぎ、スクロールを無効化
        return "break"

    def on_shift_mousewheel(self, event):
        """Shift+マウスホイールでファイル切り替え（スクロールを防ぐ）"""
        if event.delta > 0:  # ホイールアップ
            self.prev_pdf()
        else:  # ホイールダウン
            self.next_pdf()
        # returnでイベントの伝播を防ぎ、スクロールを無効化
        return "break"

    def on_mousewheel(self, event):
        """マウスホイールイベント処理（スクロール/ページ遷移の判定）"""
        if not self.pdf_document:
            return

        # Ctrlキーが押されている場合はズーム処理（イベント伝播を停止）
        if event.state & 0x4:  # Ctrlキーが押されている
            return self.on_ctrl_mousewheel(event)

        # Shiftキーが押されている場合はファイル切り替え処理（イベント伝播を停止）
        if event.state & 0x1:  # Shiftキーが押されている
            return self.on_shift_mousewheel(event)

        # ページが画面内に収まっているかチェック
        if self.is_page_fits_in_view():
            # ページ全体が表示されている場合はページ遷移
            if hasattr(event, 'delta'):
                # Windows
                if event.delta > 0:
                    self.prev_page()
                else:
                    self.next_page()
            else:
                # Linux/Unix
                if event.num == 4:
                    self.prev_page()
                elif event.num == 5:
                    self.next_page()
        else:
            # ページが画面からはみ出している場合は通常のスクロール
            if hasattr(event, 'delta'):
                # Windows
                self.pdf_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                # Linux/Unix
                if event.num == 4:
                    self.pdf_canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    self.pdf_canvas.yview_scroll(1, "units")

    def is_page_fits_in_view(self):
        """ページが画面内に完全に収まっているかを判定"""
        if not self.pdf_document or not hasattr(self, 'photo'):
            return True

        # キャンバスのサイズを取得
        canvas_width = self.pdf_canvas.winfo_width()
        canvas_height = self.pdf_canvas.winfo_height()

        # 画像のサイズを取得
        image_width = self.photo.width()
        image_height = self.photo.height()

        # 画像がキャンバスより小さいか等しい場合は収まっている
        return image_width <= canvas_width and image_height <= canvas_height

    def start_drag(self, event):
        """ドラッグ開始"""
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self.is_dragging = True
        self.pdf_canvas.configure(cursor="fleur")  # カーソルを移動アイコンに変更
        self.pdf_canvas.scan_mark(event.x, event.y)

    def on_drag(self, event):
        """ドラッグ中"""
        if not self.is_dragging:
            return

        # scan_dragtoを使用してスムーズなドラッグスクロール
        self.pdf_canvas.scan_dragto(event.x, event.y, gain=1)

    def end_drag(self, event):
        """ドラッグ終了"""
        self.is_dragging = False
        self.pdf_canvas.configure(cursor="")  # カーソルを元に戻す

    def fit_to_window(self):
        """PDFページをウィンドウサイズに合わせる"""
        if self.pdf_document:
            self.auto_fit_to_width()
            self.display_page()

    def toggle_quality_mode(self):
        """レンダリング品質モードを切り替える（循環）"""
        self.current_quality_index = (self.current_quality_index + 1) % len(self.quality_modes)
        current_mode = self.quality_modes[self.current_quality_index]
        self.quality_var.set(current_mode["name"])

        # 現在のページを再描画
        if self.pdf_document:
            self.display_page()

    def open_with_default_app(self):
        """現在のPDFファイルを既定のアプリで開く"""
        if not self.pdf_files or self.current_index >= len(self.pdf_files):
            messagebox.showwarning("警告", "開くPDFファイルがありません")
            return

        try:
            current_file = self.pdf_files[self.current_index]
            current_path = os.path.join(self.current_dir, current_file)

            # Windowsの場合
            if os.name == 'nt':
                os.startfile(current_path)
            # macOSの場合
            elif os.name == 'posix' and os.uname().sysname == 'Darwin':
                os.system(f'open "{current_path}"')
            # Linuxの場合
            else:
                os.system(f'xdg-open "{current_path}"')

        except Exception as e:
            messagebox.showerror("エラー", f"外部アプリでの起動に失敗しました: {e}")

    def show_help(self):
        """ヘルプウィンドウを表示"""
        help_window = tk.Toplevel(self.root)
        help_window.title("ParaparaPDF - 使い方")
        help_window.geometry("750x600")
        help_window.configure(bg='SlateGray1')

        # スクロール可能なテキストエリア
        text_frame = tk.Frame(help_window, bg='SlateGray1')
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Meiryo", 10), bg='white')
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # ヘルプテキストの内容
        help_text = """ParaparaPDF - 使い方ガイド

■ 概要
ParaparaPDFは、特定のディレクトリ下にある複数のPDFファイルを効率的に閲覧するためのアプリケーションです。
ディレクトリ検索、ページ範囲指定、キーボードショートカットによる高速操作が可能です。

■ 基本操作

【ファイル検索】
1. ディレクトリパターン: 検索対象のディレクトリパスを指定
   ・ワイルドカード使用可能（*は任意の文字列、?は任意の1文字）
   ・例: C:/Documents/experiment, D:/PDF/202?

2. ファイル名パターン: 検索するPDFファイル名のパターンを指定
   ・デフォルトは「*」（すべてのPDFファイル）
   ・例: report.pdf, *manual*, test??
    Tips - 拡張子「.pdf」がない場合自動的に補完されます

3. 表示ページ範囲: 表示するページを限定可能
   ・*: 全ページ（デフォルト）
   ・1-3: 1～3ページ
   ・5: 5ページのみ
   ・-8: 8ページまで
   ・10-: 10ページ以降
   ・even: 偶数ページ
   ・odd: 奇数ページ
   ・複数指定: 1-3,5,10-

【ファイル操作】
・参照ボタン: ディレクトリをダイアログから選択
・検索ボタン: 指定条件でPDFファイルを検索
・適用ボタン: ページ範囲設定を現在のファイルに適用
Tips - 入力ボックスでEnterキーを押すことでも実行できます
    　- いずれかの操作を実行すると自動的に再検索されます

■ ビューア操作

【ページ操作】
・▲ / ▼: 前のページ / 次のページ
・マウスホイール: ページがウィンドウに収まっている場合はページ切り替え
　　　　　　　　  はみ出している場合はスクロール

【ズーム操作】
・− / ＋: ズームアウト / ズームイン
・⤢: 画面サイズにフィット
・Ctrl + マウスホイール: ズーム操作
・マウスドラッグ: PDFページをドラッグしてスクロール
Tips - ファイルを開いた際は自動的に画面サイズにフィットします

【表示品質】
・⇄: レンダリング品質モード切り替え
  - Bitmap: 高速ビットマップレンダリング
  - Vector: 高品質ベクターレンダリング

【ファイル切り替え】
・◀ 前 / 次 ▶ : 前のファイル / 次のファイル
・ファイルリストから直接選択
・右クリックメニューで各種操作

■ キーボードショートカット

【ファイル操作】
・左矢印キー (←): 前のファイルに切り替え
・右矢印キー (→): 次のファイルに切り替え

【ページ操作】
・上矢印キー (↑): 前のページ
・下矢印キー (↓): 次のページ
・Page Up: 前のページ
・Page Down: 次のページ
・Home: 最初のページにジャンプ
・End: 最後のページにジャンプ

【ズーム操作】
・Ctrl + Plus (+): ズームイン
・Ctrl + Minus (-): ズームアウト
・Ctrl + 0: ズームリセット（100%）
・Ctrl + F: 画面サイズにフィット

【マウス操作】
・マウスドラッグ: PDFページをドラッグしてスクロール
・Ctrl + マウスホイール: ズーム調整
・Shift + マウスホイール: ファイル切り替え

■ 右クリックメニュー（ファイルリスト）

・開く: 選択したファイルを内蔵ビューアで開く
・既定のアプリで開く: システムの既定PDFアプリで開く
・エクスプローラーで表示: ファイルの場所をエクスプローラーで開く
・パスをコピー: ファイルの完全パスをクリップボードにコピー
・ファイル名をコピー: ファイル名のみをクリップボードにコピー

■ 高度な機能

【ページ記憶機能】
・ファイルを切り替えた際、最後に見ていたページを自動記憶
・同じファイルに戻った時に前回見ていたページから再開可能

【ワイルドカード検索】
・ディレクトリ検索で複数パスを同時指定可能
・例: C:/Work/*/PDF で、Workフォルダ下の全サブディレクトリのPDFフォルダを検索

【レンダリング最適化】
・ファイルサイズや表示速度に応じて品質モードを選択
・起動時はBitmapモード
・大きなファイルはBitmapモード、精細な表示にはVectorモードを推奨

■ 操作のコツ

1. 大量のファイルを扱う場合:
   ・ファイル名パターンで絞り込んでから検索
   ・ページ範囲を限定して表示することで、必要な情報にすばやくアクセス

2. キーボード中心の操作:
   ・矢印キーでファイルとページをすばやく切り替え
   ・Homeキー/Endキーで先頭・末尾ページに素早く移動

3. 効率的な閲覧:
   ・⤢ボタンでウィンドウサイズに最適化
   ・マウスドラッグで詳細箇所を確認

■ 注意事項

・PDFファイルの読み込みにはPyMuPDFとPillowライブラリが必要です
　　pip install PyMuPDF Pillow
　でインストールしてください
・指定したディレクトリにフォルダやファイルが多すぎる場合、検索に時間がかかる場合があります
　ディレクトリ名やファイル名にワイルドカードを使用して絞り込むことを推奨します
・大きなPDFファイルは読み込みに時間がかかる場合があります
・ファイルパスに日本語が含まれる場合、一部環境で問題が生じる可能性があります

バージョン情報: ParaparaPDF v1.0
開発環境: Windows Server 2022 Standard, Python 3.13.1
"""

        text_widget.insert(tk.END, help_text)
        text_widget.configure(state=tk.DISABLED)  # 読み取り専用に設定

        # 閉じるボタン
        close_button = tk.Button(help_window, text="閉じる", command=help_window.destroy,
                               bg='white', font=("Meiryo", 10), relief='solid', bd=1)
        close_button.pack(pady=(0, 10))

        # ウィンドウを画面中央に配置
        help_window.update_idletasks()
        x = (help_window.winfo_screenwidth() // 2) - (help_window.winfo_width() // 2)
        y = (help_window.winfo_screenheight() // 2) - (help_window.winfo_height() // 2)
        help_window.geometry(f"+{x}+{y}")

    def update_info(self):
        """現在のファイル情報を更新"""
        if self.pdf_files:
            rel_path = os.path.relpath(self.pdf_files[self.current_index], self.dir_var.get())
            self.info_var.set(f"{rel_path}")

            # ファイル位置表示を更新
            self.file_position_var.set(f"{self.current_index + 1}/{len(self.pdf_files)}")
        else:
            self.info_var.set("PDFファイルがありません")
            # ファイルがない時の位置表示
            self.file_position_var.set("")

    def open_current_pdf(self):
        """現在選択されているPDFを読み込む"""
        if self.pdf_files and 0 <= self.current_index < len(self.pdf_files):
            self.load_pdf(self.pdf_files[self.current_index])

    def prev_pdf(self):
        """前のPDFに移動"""
        if self.pdf_files and self.current_index > 0:
            self.current_index -= 1
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(self.current_index)
            self.file_listbox.see(self.current_index)
            self.update_info()
            self.open_current_pdf()

    def next_pdf(self):
        """次のPDFに移動"""
        if self.pdf_files and self.current_index < len(self.pdf_files) - 1:
            self.current_index += 1
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(self.current_index)
            self.file_listbox.see(self.current_index)
            self.update_info()
            self.open_current_pdf()

    def handle_left_key(self, event):
        """左キーのハンドリング"""
        # 入力ボックス（Entry）にフォーカスがある場合はカーソル移動を許可
        if isinstance(event.widget, tk.Entry):
            return
        # すべての場合でファイル遷移
        self.prev_pdf()
        return "break"

    def handle_right_key(self, event):
        """右キーのハンドリング"""
        # 入力ボックス（Entry）にフォーカスがある場合はカーソル移動を許可
        if isinstance(event.widget, tk.Entry):
            return
        # すべての場合でファイル遷移
        self.next_pdf()
        return "break"

    def handle_up_key(self, event):
        """上キーのハンドリング"""
        # 入力ボックス（Entry）にフォーカスがある場合は標準動作を許可
        if isinstance(event.widget, tk.Entry):
            return
        # すべての場合でページ遷移
        self.prev_page()
        return "break"

    def handle_down_key(self, event):
        """下キーのハンドリング"""
        # 入力ボックス（Entry）にフォーカスがある場合は標準動作を許可
        if isinstance(event.widget, tk.Entry):
            return
        # すべての場合でページ遷移
        self.next_page()
        return "break"

    def handle_home_key(self, event):
        """Homeキーのハンドリング"""
        # 入力ボックス（Entry）にフォーカスがある場合は標準動作を許可
        if isinstance(event.widget, tk.Entry):
            return
        # 最初のページにジャンプ
        self.goto_first_page()
        return "break"

    def handle_end_key(self, event):
        """Endキーのハンドリング"""
        # 入力ボックス（Entry）にフォーカスがある場合は標準動作を許可
        if isinstance(event.widget, tk.Entry):
            return
        # 最後のページにジャンプ
        self.goto_last_page()
        return "break"

    def goto_first_page(self):
        """最初のページに移動"""
        if self.pdf_document and self.filtered_pages:
            self.current_page = self.filtered_pages[0]
            self.display_page()

    def goto_last_page(self):
        """最後のページに移動"""
        if self.pdf_document and self.filtered_pages:
            self.current_page = self.filtered_pages[-1]
            self.display_page()

    def run(self):
        """GUIアプリケーションを実行"""
        # ライブラリチェック
        if not PDF_SUPPORT:
            messagebox.showwarning("警告",
                                  "PyMuPDF (fitz) または Pillow (PIL) がインストールされていません。\n"
                                  "PDFの内蔵表示機能が利用できません。\n\n"
                                  "インストール方法:\n"
                                  "pip install PyMuPDF Pillow")

        self.root.mainloop()

def main():
    app = PDFViewerGUI()
    app.run()

if __name__ == "__main__":
    main()
