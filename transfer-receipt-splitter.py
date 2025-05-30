import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import os
import threading
from pathlib import Path
from dotenv import load_dotenv
import base64
import json
import logging
from datetime import datetime
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
try:
    from PyPDF2 import PdfReader, PdfWriter
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    import requests
    from google.cloud import vision
    from openai import OpenAI
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

class ZipExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Transfer Receipt Splitter - ZIP解凍&PDF分割ツール")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # ログの設定
        self.setup_logging()
        
        # .envファイルを読み込み
        load_dotenv()
        
        # デフォルトフォルダの設定
        self.default_folder = self.get_default_folder()
        
        # 選択されたフォルダパス
        self.folder_path = tk.StringVar()
        
        # 設定変数（.envファイルから読み込み、デフォルト値設定）
        self.extract_option = tk.IntVar(value=int(os.getenv('EXTRACT_OPTION', '1')))
        self.overwrite_var = tk.BooleanVar(value=os.getenv('OVERWRITE_FILES', 'True').lower() == 'true')
        self.split_pdf_var = tk.BooleanVar(value=os.getenv('SPLIT_PDF', 'True').lower() == 'true')
        self.ocr_rename_var = tk.BooleanVar(value=os.getenv('OCR_RENAME', 'True').lower() == 'true')
        
        # パフォーマンス設定
        self.max_workers = min(4, os.cpu_count() or 4)  # 並列処理数制限
        
        # API クライアントを事前初期化（使い回し）
        self.vision_client = None
        self.openai_client = None
        self.init_api_clients()
        
        # GUI要素の作成
        self.create_widgets()
        
        # ウィンドウを画面中央に配置（GUI作成後に実行）
        self.root.after(10, self.center_window)
        
        # デフォルトフォルダを設定して自動検索
        saved_folder = os.getenv('LAST_FOLDER')
        if saved_folder and Path(saved_folder).exists():
            self.folder_path.set(saved_folder)
        elif self.default_folder:
            self.folder_path.set(str(self.default_folder))
        
        if self.folder_path.get():
            self.scan_zip_files()
        
        # 設定変更時のコールバック設定
        self.setup_setting_callbacks()
    
    def init_api_clients(self):
        """APIクライアントを事前初期化"""
        try:
            if OCR_AVAILABLE:
                # Google Cloud Vision クライアント
                self.vision_client = vision.ImageAnnotatorClient()
                
                # OpenAI クライアント
                api_key = os.getenv('OPENAI_API_KEY')
                if api_key:
                    self.openai_client = OpenAI(api_key=api_key)
                    
            self.logger.info("APIクライアント初期化完了")
        except Exception as e:
            self.logger.warning(f"APIクライアント初期化警告: {e}")
    
    def setup_logging(self):
        """ログ設定を初期化"""
        # 実行中のPythonファイルと同じディレクトリにログファイルを作成
        script_dir = Path(__file__).parent if '__file__' in globals() else Path.cwd()
        log_file = script_dir / "transfer-receipt-splitter.log"
        
        # 前回のログファイルを削除
        if log_file.exists():
            log_file.unlink()
        
        # ログ設定
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()  # コンソールにも出力
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("=" * 50)
        self.logger.info("Transfer Receipt Splitter 開始")
        self.logger.info(f"ログファイル: {log_file}")
        self.logger.info("=" * 50)
    
    def center_window(self):
        """ウィンドウを画面中央に配置"""
        # ウィンドウの描画を完了させる
        self.root.update_idletasks()
        
        # ウィンドウのサイズを取得
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        
        # サイズが正しく取得できない場合は設定値を使用
        if window_width <= 1:
            window_width = 600
        if window_height <= 1:
            window_height = 500
        
        # 画面のサイズを取得
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 中央座標を計算
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # y座標を少し上に調整（タスクバーを考慮）
        y = max(0, y - 50)
        
        # ウィンドウの位置を設定
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # デバッグ情報（必要に応じて削除）
        self.logger.info(f"ウィンドウサイズ: {window_width}x{window_height}")
        self.logger.info(f"画面サイズ: {screen_width}x{screen_height}")
        self.logger.info(f"配置位置: {x}, {y}")
        print(f"ウィンドウサイズ: {window_width}x{window_height}")
        print(f"画面サイズ: {screen_width}x{screen_height}")
        print(f"配置位置: {x}, {y}")
    
    def setup_setting_callbacks(self):
        """設定変更時のコールバックを設定"""
        self.extract_option.trace('w', self.save_settings)
        self.overwrite_var.trace('w', self.save_settings)
        self.split_pdf_var.trace('w', self.save_settings)
        self.ocr_rename_var.trace('w', self.save_settings)
        self.folder_path.trace('w', self.save_folder_setting)
    
    def get_default_folder(self):
        """デフォルトフォルダを取得"""
        # .envファイルからDEFAULT_FOLDERを取得
        env_folder = os.getenv('DEFAULT_FOLDER')
        if env_folder:
            folder_path = Path(env_folder).expanduser()
            if folder_path.exists():
                return folder_path
        
        # .envに設定がない場合、ユーザーのダウンロードフォルダを使用
        download_paths = [
            Path.home() / "Downloads",  # Windows/Mac/Linux共通
            Path.home() / "ダウンロード",  # 日本語Windows
            Path.home() / "다운로드",     # 韓国語
            Path.home() / "下载",        # 中国語簡体字
            Path.home() / "下載",        # 中国語繁体字
        ]
        
        for path in download_paths:
            if path.exists():
                return path
        
        # どれも見つからない場合はホームフォルダ
        return Path.home()
        
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # フォルダ選択部分
        folder_frame = ttk.LabelFrame(main_frame, text="フォルダ選択", padding="10")
        folder_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # フォルダパス表示
        ttk.Label(folder_frame, text="選択フォルダ:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        path_frame = ttk.Frame(folder_frame)
        path_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        path_frame.columnconfigure(0, weight=1)
        
        self.path_entry = ttk.Entry(path_frame, textvariable=self.folder_path, state="readonly", width=50)
        self.path_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(path_frame, text="参照", command=self.select_folder, width=10).grid(row=0, column=1)
        
        # オプション部分
        options_frame = ttk.LabelFrame(main_frame, text="解凍オプション", padding="10")
        options_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 解凍先オプション
        ttk.Radiobutton(options_frame, text="各ZIPファイルごとに個別フォルダを作成", 
                       variable=self.extract_option, value=1).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(options_frame, text="選択フォルダ内に直接解凍", 
                       variable=self.extract_option, value=2).grid(row=1, column=0, sticky=tk.W)
        
        # 上書きオプション
        ttk.Checkbutton(options_frame, text="既存ファイルを上書きする", 
                       variable=self.overwrite_var).grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
        # PDF分割オプション
        pdf_frame = ttk.Frame(options_frame)
        pdf_frame.grid(row=3, column=0, sticky=tk.W, pady=(5, 0))
        
        if PDF_AVAILABLE:
            ttk.Checkbutton(pdf_frame, text="PDFファイルを1ページずつ分割する", 
                           variable=self.split_pdf_var).grid(row=0, column=0, sticky=tk.W)
        else:
            ttk.Label(pdf_frame, text="⚠️ PDF分割機能を使用するには 'pip install PyPDF2' が必要です", 
                     foreground="orange").grid(row=0, column=0, sticky=tk.W)
        
        # OCR・AI自動リネーム機能
        ocr_frame = ttk.Frame(options_frame)
        ocr_frame.grid(row=4, column=0, sticky=tk.W, pady=(5, 0))
        
        if OCR_AVAILABLE:
            ttk.Checkbutton(ocr_frame, text="OCR&AI自動リネーム機能を使用する", 
                           variable=self.ocr_rename_var).grid(row=0, column=0, sticky=tk.W)
            ttk.Label(ocr_frame, text="   (Google Cloud Vision + OpenAI APIが必要)", 
                     foreground="gray", font=("", 8)).grid(row=1, column=0, sticky=tk.W)
        else:
            ttk.Label(ocr_frame, text="⚠️ OCR機能を使用するには追加ライブラリが必要です", 
                     foreground="orange").grid(row=0, column=0, sticky=tk.W)
            ttk.Label(ocr_frame, text="   pip install pdf2image google-cloud-vision openai requests", 
                     foreground="gray", font=("", 8)).grid(row=1, column=0, sticky=tk.W)
        
        # ZIPファイル一覧
        list_frame = ttk.LabelFrame(main_frame, text="見つかったZIPファイル", padding="10")
        list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # リストボックスとスクロールバー
        list_container = ttk.Frame(list_frame)
        list_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_container.columnconfigure(0, weight=1)
        list_container.rowconfigure(0, weight=1)
        
        self.zip_listbox = tk.Listbox(list_container, height=8)
        self.zip_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.zip_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.zip_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 進捗表示
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_var = tk.StringVar(value="準備完了")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # ボタン部分
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))
        
        ttk.Button(button_frame, text="再検索", command=self.scan_zip_files, 
                  width=15).grid(row=0, column=0, padx=(0, 10))
        
        self.extract_button = ttk.Button(button_frame, text="解凍開始", command=self.start_extraction, 
                                        width=15, state="disabled")
        self.extract_button.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(button_frame, text="終了", command=self.root.quit, 
                  width=10).grid(row=0, column=2)
        
        # グリッドの重み設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        folder_frame.columnconfigure(0, weight=1)
        
    def select_folder(self):
        """フォルダ選択ダイアログを表示"""
        # 初期ディレクトリを現在のフォルダパスまたはデフォルトフォルダに設定
        initial_dir = self.folder_path.get() or str(self.default_folder)
        
        folder = filedialog.askdirectory(
            title="ZIPファイルがあるフォルダを選択してください",
            initialdir=initial_dir
        )
        if folder:
            self.folder_path.set(folder)
            self.zip_listbox.delete(0, tk.END)
            self.progress_var.set("フォルダが選択されました。ZIPファイルを検索中...")
            self.extract_button.config(state="disabled")
            # 自動的にZIPファイル検索を実行
            self.scan_zip_files()
    
    def scan_zip_files(self):
        """選択されたフォルダ内のZIPファイルをスキャン"""
        if not self.folder_path.get():
            messagebox.showwarning("警告", "フォルダを選択してください。")
            return
        
        folder = Path(self.folder_path.get())
        if not folder.exists():
            messagebox.showerror("エラー", "選択されたフォルダが存在しません。")
            return
        
        # ZIPファイルを検索
        zip_files = list(folder.glob("*.zip"))
        
        # リストボックスを更新
        self.zip_listbox.delete(0, tk.END)
        
        if not zip_files:
            self.progress_var.set("ZIPファイルが見つかりませんでした。")
            self.extract_button.config(state="disabled")
            return
        
        for zip_file in zip_files:
            self.zip_listbox.insert(tk.END, zip_file.name)
        
        self.progress_var.set(f"{len(zip_files)}個のZIPファイルが見つかりました。解凍準備完了。")
        self.extract_button.config(state="normal")
    
    def start_extraction(self):
        """解凍処理を開始（別スレッドで実行）"""
        if self.zip_listbox.size() == 0:
            messagebox.showwarning("警告", "解凍するZIPファイルがありません。")
            return
        
        # ボタンを無効化
        self.extract_button.config(state="disabled")
        
        # 別スレッドで解凍処理を実行
        threading.Thread(target=self.extract_files, daemon=True).start()
    
    def extract_files(self):
        """ZIPファイルを解凍（並列処理対応）"""
        folder = Path(self.folder_path.get())
        zip_files = list(folder.glob("*.zip"))
        total_files = len(zip_files)
        
        # プログレスバーの設定
        self.progress_bar.config(maximum=total_files, value=0)
        
        success_count = 0
        error_files = []
        
        start_time = time.time()
        
        for i, zip_file in enumerate(zip_files):
            try:
                # 進捗表示の更新
                self.progress_var.set(f"解凍中: {zip_file.name} ({i+1}/{total_files})")
                self.root.update_idletasks()
                
                self.logger.info(f"ZIP解凍開始: {zip_file.name} ({i+1}/{total_files})")
                
                # 解凍先の決定
                if self.extract_option.get() == 1:
                    # 各ZIPファイルごとに個別フォルダを作成
                    extract_path = folder / zip_file.stem
                    
                    # 同名フォルダが存在する場合は削除
                    if extract_path.exists():
                        shutil.rmtree(extract_path)
                    
                    extract_path.mkdir(exist_ok=True)
                else:
                    # 選択フォルダ内に直接解凍
                    extract_path = folder
                
                # 前回の作業ファイルを削除（PDF分割機能が有効な場合）
                if self.split_pdf_var.get() and PDF_AVAILABLE:
                    self.cleanup_previous_files(extract_path)
                
                # ZIPファイルを解凍
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    if self.overwrite_var.get():
                        # 上書きする場合
                        zip_ref.extractall(extract_path)
                    else:
                        # 上書きしない場合は既存ファイルをチェック
                        for member in zip_ref.namelist():
                            target_path = extract_path / member
                            if not target_path.exists():
                                zip_ref.extract(member, extract_path)
                
                # PDF分割処理
                if self.split_pdf_var.get() and PDF_AVAILABLE:
                    self.progress_var.set(f"PDF分割中: {zip_file.name} ({i+1}/{total_files})")
                    self.root.update_idletasks()
                    split_files = self.split_pdfs_in_folder_optimized(extract_path)
                    
                    # OCR & AI自動リネーム処理
                    if self.ocr_rename_var.get() and OCR_AVAILABLE and split_files:
                        self.progress_var.set(f"OCR&リネーム中: {zip_file.name} ({i+1}/{total_files})")
                        self.root.update_idletasks()
                        self.process_ocr_and_rename(split_files)
                
                success_count += 1
                
            except Exception as e:
                error_msg = f"{zip_file.name}: {str(e)}"
                error_files.append(error_msg)
                self.logger.error(f"ZIP処理エラー: {error_msg}")
            
            # プログレスバーの更新
            self.progress_bar.config(value=i+1)
            self.root.update_idletasks()
        
        # 処理時間ログ
        elapsed_time = time.time() - start_time
        self.logger.info(f"全ZIP処理完了: {elapsed_time:.1f}秒")
        
        # 完了メッセージ
        if error_files:
            error_msg = "\n".join(error_files)
            messagebox.showwarning("警告", 
                                 f"解凍完了: {success_count}/{total_files}\n\n"
                                 f"エラーが発生したファイル:\n{error_msg}")
        else:
            features = []
            if self.split_pdf_var.get() and PDF_AVAILABLE:
                features.append("PDF分割")
            if self.ocr_rename_var.get() and OCR_AVAILABLE:
                features.append("OCR&AI自動リネーム")
            
            feature_text = "と" + "・".join(features) if features else ""
            messagebox.showinfo("完了", 
                              f"すべてのZIPファイル({total_files}個)の解凍{feature_text}が完了しました。\n"
                              f"処理時間: {elapsed_time:.1f}秒")
        
        # UI状態をリセット
        self.progress_var.set("解凍完了")
        self.extract_button.config(state="normal")
        self.progress_bar.config(value=0)
    
    def split_pdfs_in_folder_optimized(self, folder_path):
        """最適化されたPDF分割処理"""
        pdf_files = list(folder_path.glob("*.pdf"))
        split_files = []
        
        # 並列処理でPDF分割を高速化
        with ThreadPoolExecutor(max_workers=min(2, self.max_workers)) as executor:
            future_to_pdf = {}
            
            for pdf_file in pdf_files:
                # 既に分割されたファイルはスキップ
                if "_page_" not in pdf_file.stem:
                    future = executor.submit(self.split_single_pdf_optimized, pdf_file)
                    future_to_pdf[future] = pdf_file
            
            for future in as_completed(future_to_pdf):
                try:
                    result = future.result()
                    if result:
                        split_files.extend(result)
                except Exception as e:
                    pdf_file = future_to_pdf[future]
                    self.logger.error(f"PDF分割エラー ({pdf_file.name}): {e}")
        
        return split_files
    
    def split_single_pdf_optimized(self, pdf_file):
        """最適化された単一PDF分割"""
        split_files = []
        
        try:
            # メモリ効率を改善
            with open(pdf_file, 'rb') as file:
                reader = PdfReader(file)
                total_pages = len(reader.pages)
                
                # バッチ処理で高速化
                for page_num in range(total_pages):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[page_num])
                    
                    # 出力ファイル名を生成
                    page_filename = f"{pdf_file.stem}_page_{page_num + 1:03d}.pdf"
                    output_path = pdf_file.parent / page_filename
                    
                    # メモリ効率的な書き込み
                    with open(output_path, 'wb') as output_file:
                        writer.write(output_file)
                    
                    split_files.append(output_path)
            
            # 元のPDFファイルを削除
            pdf_file.unlink()
            
            return split_files
            
        except Exception as e:
            raise Exception(f"PDF分割処理失敗: {str(e)}")
    
    def cleanup_previous_files(self, folder_path):
        """高速化された前回ファイル削除"""
        try:
            # glob使用で高速化
            pattern = "*_page_*.pdf"
            deleted_files = list(folder_path.glob(pattern))
            
            for pdf_file in deleted_files:
                pdf_file.unlink()
            
            if deleted_files:
                self.logger.info(f"前回ファイル削除: {len(deleted_files)}個")
            
        except Exception as e:
            self.logger.error(f"前回ファイル削除エラー: {e}")
    
    def process_ocr_and_rename(self, pdf_files):
        """OCRとAI自動リネーム処理（並列処理版）"""
        self.logger.info(f"OCR&リネーム処理開始: {len(pdf_files)}個のPDFファイル (並列処理: {self.max_workers})")
        
        start_time = time.time()
        
        # 並列処理で高速化
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # タスクを投入
            future_to_pdf = {
                executor.submit(self.process_single_pdf, i, pdf_file, len(pdf_files)): pdf_file 
                for i, pdf_file in enumerate(pdf_files)
            }
            
            # 結果を順次処理
            for future in as_completed(future_to_pdf):
                pdf_file = future_to_pdf[future]
                try:
                    future.result()
                except Exception as e:
                    self.logger.error(f"並列処理エラー ({pdf_file.name}): {e}")
        
        elapsed_time = time.time() - start_time
        self.logger.info(f"OCR&リネーム処理完了 (処理時間: {elapsed_time:.1f}秒)")
    
    def process_single_pdf(self, index, pdf_file, total_files):
        """単一PDFファイルの処理"""
        try:
            self.logger.info(f"[{index+1}/{total_files}] 処理開始: {pdf_file.name}")
            
            # PDF → JPEG変換
            start_convert = time.time()
            jpeg_file = self.pdf_to_jpeg_optimized(pdf_file)
            convert_time = time.time() - start_convert
            self.logger.info(f"JPEG変換完了: {convert_time:.1f}秒")
            
            # OCR処理
            start_ocr = time.time()
            ocr_text = self.perform_ocr_optimized(jpeg_file)
            ocr_time = time.time() - start_ocr
            self.logger.info(f"OCR処理完了: {ocr_time:.1f}秒")
            
            # 文書種別を判定
            doc_type = self.determine_document_type(ocr_text)
            self.logger.info(f"文書種別判定: {doc_type}")
            
            # OCR結果の詳細分析（払込取扱票の場合のみ、ログレベル調整）
            if doc_type == "払込取扱票" and self.logger.level <= logging.INFO:
                self.analyze_payment_slip_ocr_brief(ocr_text)
            
            # OpenAI APIでファイル名生成
            start_ai = time.time()
            new_name = self.generate_filename_with_ai_optimized(ocr_text, doc_type)
            ai_time = time.time() - start_ai
            self.logger.info(f"AI ファイル名生成完了: {ai_time:.1f}秒")
            
            # ファイルリネーム
            if new_name:
                old_name = pdf_file.name
                self.rename_pdf_file(pdf_file, new_name)
                self.logger.info(f"リネーム完了: {old_name} → {new_name}.pdf")
            else:
                self.logger.warning(f"ファイル名生成失敗: {pdf_file.name}")
            
            # 一時的なJPEGファイルを削除
            if jpeg_file.exists():
                jpeg_file.unlink()
            
            total_time = time.time() - (start_convert + convert_time + ocr_time + ai_time)
            self.logger.info(f"[{index+1}/{total_files}] 処理完了 (合計: {total_time:.1f}秒)")
                
        except Exception as e:
            self.logger.error(f"PDF処理エラー ({pdf_file.name}): {e}")
    
    def pdf_to_jpeg_optimized(self, pdf_file):
        """最適化されたPDF→JPEG変換"""
        try:
            # DPIを上げて精度を向上（150→200）
            images = convert_from_path(
                pdf_file, 
                dpi=200,  # 150→200に変更
                first_page=1, 
                last_page=1,
                fmt='jpeg'  # フォーマット指定で高速化
            )
            
            if images:
                jpeg_file = pdf_file.parent / f"{pdf_file.stem}_temp.jpg"
                images[0].save(jpeg_file, 'JPEG', quality=95, optimize=True)  # 品質を95に変更
                return jpeg_file
            else:
                raise Exception("PDF変換失敗: 画像が生成されませんでした")
                
        except Exception as e:
            raise Exception(f"PDF→JPEG変換エラー: {str(e)}")
    
    def perform_ocr_optimized(self, image_file):
        """最適化されたOCR処理"""
        try:
            # 事前初期化されたクライアントを使用
            if not self.vision_client:
                raise Exception("Vision APIクライアントが初期化されていません")
            
            # 画像ファイルを読み込み
            with open(image_file, 'rb') as image_file_obj:
                content = image_file_obj.read()
            
            image = vision.Image(content=content)
            
            # OCR実行（高速化オプション）
            response = self.vision_client.text_detection(
                image=image,
                # image_context=vision.ImageContext(
                #     language_hints=["ja"]  # 日本語ヒントで精度向上
                # )
            )
            
            if response.text_annotations:
                ocr_result = response.text_annotations[0].description
                self.logger.info(f"OCR成功: {len(ocr_result)}文字のテキストを抽出")
                # OCR結果をログに出力
                self.logger.info("OCR結果:")
                self.logger.info("-" * 50)
                self.logger.info(ocr_result)
                self.logger.info("-" * 50)
                return ocr_result
            else:
                self.logger.warning(f"OCR結果なし: {image_file}")
                return ""
                
        except Exception as e:
            self.logger.error(f"OCR処理エラー: {e}")
            raise Exception(f"OCR処理エラー: {str(e)}")
    
    def analyze_payment_slip_ocr_brief(self, ocr_text):
        """払込取扱票のOCR結果を簡潔に分析（パフォーマンス重視）"""
        lines = ocr_text.split('\n')
        
        # 重要な情報のみ抽出
        important_info = []
        for i, line in enumerate(lines):
            if any(keyword in line for keyword in ['加入者名', 'おなまえ', '円']):
                important_info.append(f"[{i}]: {line}")
        
        if important_info:
            self.logger.info(f"重要情報: {'; '.join(important_info[:3])}")  # 最大3行まで
    
    def generate_filename_with_ai_optimized(self, ocr_text, doc_type):
        """最適化されたAIファイル名生成"""
        try:
            if not self.openai_client:
                self.logger.error("OpenAI APIクライアントが初期化されていません")
                return None
            
            # 文書種別に応じた処理分岐
            if doc_type == "振替受払通知票":
                return self.generate_transfer_notification_filename_optimized(ocr_text)
            elif doc_type == "振替受入明細書":
                return self.generate_transfer_detail_filename_optimized(ocr_text)
            elif doc_type == "払込取扱票":
                return self.generate_payment_slip_filename_optimized(ocr_text)
            else:
                return self.generate_general_filename_optimized(ocr_text)
            
        except Exception as e:
            self.logger.error(f"AI ファイル名生成エラー: {e}")
            return None
    
    def generate_payment_slip_filename_optimized(self, ocr_text):
        """最適化された払込取扱票ファイル名生成"""
        # OCRテキストを短縮して処理速度向上
        text_preview = ocr_text[:800] if len(ocr_text) > 800 else ocr_text
        
        prompt = f"""
以下は払込取扱票のOCRテキストです。簡潔にファイル名を生成してください：

OCRテキスト:
{text_preview}

出力フォーマット:
YYYY年MM月DD日振込 [個人名]より払込取扱票 金額[金額]円

注意:
- 日付不明時は「20XX年XX月XX日」
- 個人名（石井孝志、倉本猛、西田啓など）を使用（善法寺は除外）
- 上記フォーマット以外は出力禁止

ファイル名:
"""
        
        response = self.openai_client.chat.completions.create(
            model="gpt-3.5-turbo",  # gpt-4より高速
            messages=[
                {"role": "system", "content": "払込取扱票のファイル名を生成。加入者名と依頼人名を区別し、個人名のみ使用。"},
                {"role": "user", "content": prompt}
            ],
            max_tokens=80,  # トークン数削減
            temperature=0.0
        )
        
        filename = response.choices[0].message.content.strip()
        return self.sanitize_filename(filename)
    
    def generate_transfer_notification_filename_optimized(self, ocr_text):
        """最適化された振替受払通知票ファイル名生成"""
        text_preview = ocr_text[:500] if len(ocr_text) > 500 else ocr_text
        
        prompt = f"""
振替受払通知票のファイル名生成:

{text_preview}

フォーマット: YYYY年MM月DD日振込 振替受払通知票 金額[金額]円
日付不明時: 20XX年XX月XX日振込 振替受払通知票 金額[金額]円

ファイル名:
"""
        
        response = self.openai_client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=60,
            temperature=0.0
        )
        
        return self.sanitize_filename(response.choices[0].message.content.strip())
    
    def generate_transfer_detail_filename_optimized(self, ocr_text):
        """最適化された振替受入明細書ファイル名生成"""
        text_preview = ocr_text[:500] if len(ocr_text) > 500 else ocr_text
        
        prompt = f"""
振替受入明細書のファイル名生成:

{text_preview}

フォーマット: YYYY年MM月DD日振込 [カタカナ依頼人名]より振替受入明細書 金額[金額]円
日付不明時: 20XX年XX月XX日振込 [カタカナ依頼人名]より振替受入明細書 金額[金額]円

ファイル名:
"""
        
        response = self.openai_client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=70,
            temperature=0.0
        )
        
        return self.sanitize_filename(response.choices[0].message.content.strip())
    
    def generate_general_filename_optimized(self, ocr_text):
        """最適化された一般文書ファイル名生成"""
        text_preview = ocr_text[:300] if len(ocr_text) > 300 else ocr_text
        
        prompt = f"""
文書内容から適切なファイル名を生成:

{text_preview}

要件: 50文字以内、日本語可、内容を表す名前

ファイル名:
"""
        
        response = self.openai_client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=50,
            temperature=0.3
        )
        
        return self.sanitize_filename(response.choices[0].message.content.strip())

    
    def determine_document_type(self, ocr_text):
        """OCR結果から文書種別を判定"""
        self.logger.info(f"文書種別判定開始: OCR文字数={len(ocr_text)}")
        
        # 振替受払通知票の判定（条件を緩和）
        transfer_notification_keywords = [
            "振替受払通知票", "振替受払", "振替受払通知",  # OCR誤認識対応
            "貯金事務センター", "貯金事務",  # 部分一致
            "振替受入", "振替受払", "振替受入通知"  # 類似表現
        ]
        
        found_keywords = []
        for keyword in transfer_notification_keywords:
            if keyword in ocr_text:
                found_keywords.append(keyword)
        
        if found_keywords:
            self.logger.info(f"振替受払通知票の判定キーワード発見: {found_keywords}")
            self.logger.info("判定結果: 振替受払通知票")
            return "振替受払通知票"
        
        # 振替受入明細書の判定
        if "振替受入明細書" in ocr_text and "貯金事務センター" in ocr_text:
            self.logger.info("判定結果: 振替受入明細書")
            return "振替受入明細書"
        
        # 払込取扱票の判定（複数の条件でチェック）
        payment_slip_keywords = [
            "払込取扱票", "払达取扱票", "払込取扱",  # OCR誤認識対応
            "口座記号番号はお間違え", "口座記号・番号はお間違え",  # 部分一致
            "番号はお間違えないよう", "番号はお間違えのないよう",
            "加入者名", "ご依頼人", "通信欄"
        ]
        
        found_keywords = []
        for keyword in payment_slip_keywords:
            if keyword in ocr_text:
                found_keywords.append(keyword)
        
        if found_keywords:
            self.logger.info(f"払込取扱票の判定キーワード発見: {found_keywords}")
            self.logger.info("判定結果: 払込取扱票")
            return "払込取扱票"
        
        # どれにも該当しない場合
        self.logger.info("判定結果: 一般文書")
        return "一般文書"
    
    def sanitize_filename(self, filename):
        """ファイル名のサニタイズ処理"""
        # ファイルシステムで使用できない文字を除去（アンダースコアは除外）
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        original_filename = filename
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        if original_filename != filename:
            self.logger.info(f"無効文字を置換: {original_filename} → {filename}")
        
        final_filename = filename[:50]  # 50文字制限
        if len(filename) > 50:
            self.logger.info(f"ファイル名を50文字に短縮: {filename} → {final_filename}")
        
        return final_filename
    
    def rename_pdf_file(self, pdf_file, new_name):
        """PDFファイルをリネーム"""
        try:
            if new_name:
                new_path = pdf_file.parent / f"{new_name}.pdf"
                original_new_path = new_path
                
                # ファイル名の重複を避ける
                counter = 1
                while new_path.exists():
                    new_path = pdf_file.parent / f"{new_name}_{counter}.pdf"
                    counter += 1
                
                if original_new_path != new_path:
                    self.logger.info(f"ファイル名重複回避: {original_new_path.name} → {new_path.name}")
                
                old_path = pdf_file
                pdf_file.rename(new_path)
                self.logger.info(f"ファイルリネーム成功: {old_path.name} → {new_path.name}")
                
        except Exception as e:
            self.logger.error(f"ファイルリネームエラー ({pdf_file.name}): {e}")
            print(f"ファイルリネームエラー: {e}")
    
    def save_settings(self, *args):
        """設定を.envファイルに保存"""
        try:
            self.update_env_file({
                'EXTRACT_OPTION': str(self.extract_option.get()),
                'OVERWRITE_FILES': str(self.overwrite_var.get()),
                'SPLIT_PDF': str(self.split_pdf_var.get()),
                'OCR_RENAME': str(self.ocr_rename_var.get())
            })
        except Exception as e:
            print(f"設定保存エラー: {e}")
    
    def save_folder_setting(self, *args):
        """フォルダ設定を.envファイルに保存"""
        try:
            if self.folder_path.get():
                self.update_env_file({'LAST_FOLDER': self.folder_path.get()})
        except Exception as e:
            print(f"フォルダ設定保存エラー: {e}")
    
    def update_env_file(self, new_settings):
        """環境変数ファイルを更新"""
        env_file = Path('.env')
        
        # 既存の設定を読み込み
        existing_settings = {}
        if env_file.exists():
            with open(env_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, value = line.split('=', 1)
                        existing_settings[key] = value
        
        # 新しい設定をマージ
        existing_settings.update(new_settings)
        
        # .envファイルに書き込み
        with open(env_file, 'w', encoding='utf-8') as f:
            f.write("# Transfer Receipt Splitter 設定ファイル\n")
            f.write("# 自動生成されたファイルです\n\n")
            
            # デフォルトフォルダ設定
            if 'DEFAULT_FOLDER' in existing_settings:
                f.write(f"DEFAULT_FOLDER={existing_settings['DEFAULT_FOLDER']}\n")
            
            # 最後に使用したフォルダ
            if 'LAST_FOLDER' in existing_settings:
                f.write(f"LAST_FOLDER={existing_settings['LAST_FOLDER']}\n")
            
            f.write("\n# UI設定\n")
            
            # UI設定
            ui_settings = ['EXTRACT_OPTION', 'OVERWRITE_FILES', 'SPLIT_PDF', 'OCR_RENAME']
            for key in ui_settings:
                if key in existing_settings:
                    f.write(f"{key}={existing_settings[key]}\n")
            
            f.write("\n# API設定\n")
            
            # API設定
            api_settings = ['OPENAI_API_KEY', 'GOOGLE_APPLICATION_CREDENTIALS']
            for key in api_settings:
                if key in existing_settings:
                    f.write(f"{key}={existing_settings[key]}\n")

def main():
    root = tk.Tk()
    app = ZipExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
                    