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
        """ZIPファイルを解凍"""
        folder = Path(self.folder_path.get())
        zip_files = list(folder.glob("*.zip"))
        total_files = len(zip_files)
        
        # プログレスバーの設定
        self.progress_bar.config(maximum=total_files, value=0)
        
        success_count = 0
        error_files = []
        
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
                        self.logger.info(f"既存フォルダを削除: {extract_path}")
                        shutil.rmtree(extract_path)
                        self.logger.info(f"既存フォルダ削除完了: {extract_path}")
                    
                    extract_path.mkdir(exist_ok=True)
                    self.logger.info(f"解凍先フォルダ作成: {extract_path}")
                else:
                    # 選択フォルダ内に直接解凍
                    extract_path = folder
                    self.logger.info(f"直接解凍モード: {extract_path}")
                
                # 前回の作業ファイルを削除（PDF分割機能が有効な場合）
                if self.split_pdf_var.get() and PDF_AVAILABLE:
                    self.cleanup_previous_files(extract_path)
                
                # ZIPファイルを解凍
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    if self.overwrite_var.get():
                        # 上書きする場合
                        zip_ref.extractall(extract_path)
                        self.logger.info(f"ZIP解凍完了（上書きモード）: {len(zip_ref.namelist())}個のファイル")
                    else:
                        # 上書きしない場合は既存ファイルをチェック
                        extracted_count = 0
                        skipped_count = 0
                        for member in zip_ref.namelist():
                            target_path = extract_path / member
                            if not target_path.exists():
                                zip_ref.extract(member, extract_path)
                                extracted_count += 1
                            else:
                                skipped_count += 1
                        self.logger.info(f"ZIP解凍完了（スキップモード）: {extracted_count}個抽出, {skipped_count}個スキップ")
                
                # PDF分割処理
                if self.split_pdf_var.get() and PDF_AVAILABLE:
                    self.progress_var.set(f"PDF分割中: {zip_file.name} ({i+1}/{total_files})")
                    self.root.update_idletasks()
                    self.logger.info(f"PDF分割処理開始: {zip_file.name}")
                    split_files = self.split_pdfs_in_folder(extract_path)
                    self.logger.info(f"PDF分割処理完了: {len(split_files)}個のファイルを生成")
                    
                    # OCR & AI自動リネーム処理
                    if self.ocr_rename_var.get() and OCR_AVAILABLE and split_files:
                        self.progress_var.set(f"OCR&リネーム中: {zip_file.name} ({i+1}/{total_files})")
                        self.root.update_idletasks()
                        self.process_ocr_and_rename(split_files)
                
                success_count += 1
                self.logger.info(f"ZIP処理完了: {zip_file.name}")
                
            except Exception as e:
                error_msg = f"{zip_file.name}: {str(e)}"
                error_files.append(error_msg)
                self.logger.error(f"ZIP処理エラー: {error_msg}")
            
            # プログレスバーの更新
            self.progress_bar.config(value=i+1)
            self.root.update_idletasks()
        
        # 完了メッセージ
        if error_files:
            error_msg = "\n".join(error_files)
            self.logger.warning(f"処理完了（一部エラー）: 成功 {success_count}/{total_files}")
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
            self.logger.info(f"全処理完了: {total_files}個のZIPファイル{feature_text}")
            messagebox.showinfo("完了", f"すべてのZIPファイル({total_files}個)の解凍{feature_text}が完了しました。")
        
        # UI状態をリセット
        self.progress_var.set("解凍完了")
        self.extract_button.config(state="normal")
        self.progress_bar.config(value=0)
    
    def cleanup_previous_files(self, folder_path):
        """前回の作業ファイルを削除"""
        try:
            self.logger.info(f"前回作業ファイルのクリーンアップ開始: {folder_path}")
            
            # 分割されたPDFファイル（_page_数字.pdf）を削除
            deleted_count = 0
            for pdf_file in folder_path.glob("*_page_*.pdf"):
                pdf_file.unlink()
                deleted_count += 1
                self.logger.info(f"前回ファイル削除: {pdf_file.name}")
            
            self.logger.info(f"前回作業ファイルクリーンアップ完了: {deleted_count}個のファイルを削除")
            
        except Exception as e:
            self.logger.error(f"前回ファイル削除エラー: {e}")
            print(f"前回ファイル削除エラー: {e}")
    
    def split_pdfs_in_folder(self, folder_path):
        """フォルダ内のPDFファイルを1ページずつ分割"""
        pdf_files = list(folder_path.glob("*.pdf"))
        split_files = []
        
        for pdf_file in pdf_files:
            # 既に分割されたファイルはスキップ
            if "_page_" in pdf_file.stem:
                continue
                
            try:
                split_result = self.split_single_pdf(pdf_file)
                split_files.extend(split_result)
            except Exception as e:
                print(f"PDF分割エラー ({pdf_file.name}): {e}")
        
        return split_files
    
    def split_single_pdf(self, pdf_file):
        """単一のPDFファイルを1ページずつ分割"""
        split_files = []
        
        try:
            # PDFを読み込み
            with open(pdf_file, 'rb') as file:
                reader = PdfReader(file)
                total_pages = len(reader.pages)
                
                # 各ページを個別ファイルとして保存
                for page_num in range(total_pages):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[page_num])
                    
                    # 出力ファイル名を生成（ゼロパディング）
                    page_filename = f"{pdf_file.stem}_page_{page_num + 1:03d}.pdf"
                    output_path = pdf_file.parent / page_filename
                    
                    # ページを保存
                    with open(output_path, 'wb') as output_file:
                        writer.write(output_file)
                    
                    split_files.append(output_path)
            
            # 元のPDFファイルを削除
            pdf_file.unlink()
            
            return split_files
            
        except Exception as e:
            raise Exception(f"PDF分割処理失敗: {str(e)}")
    
    def process_ocr_and_rename(self, pdf_files):
        """OCRとAI自動リネーム処理"""
        self.logger.info(f"OCR&リネーム処理開始: {len(pdf_files)}個のPDFファイル")
        
        for i, pdf_file in enumerate(pdf_files):
            try:
                self.logger.info(f"[{i+1}/{len(pdf_files)}] 処理開始: {pdf_file.name}")
                
                # PDF → JPEG変換
                self.logger.info(f"PDF→JPEG変換中: {pdf_file.name}")
                jpeg_file = self.pdf_to_jpeg(pdf_file)
                self.logger.info(f"JPEG変換完了: {jpeg_file.name}")
                
                # Google Cloud Vision OCR
                self.logger.info(f"OCR処理中: {pdf_file.name}")
                ocr_text = self.perform_ocr(jpeg_file)
                
                # OCR結果をログに記録（長すぎる場合は省略）
                ocr_preview = ocr_text[:200] + "..." if len(ocr_text) > 200 else ocr_text
                self.logger.info(f"OCR結果 ({len(ocr_text)}文字): {ocr_preview}")
                
                # OpenAI APIでファイル名生成
                self.logger.info(f"AI ファイル名生成中: {pdf_file.name}")
                new_name = self.generate_filename_with_ai(ocr_text)
                self.logger.info(f"生成されたファイル名: {new_name}")
                
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
                    self.logger.info(f"一時ファイル削除: {jpeg_file.name}")
                
                self.logger.info(f"[{i+1}/{len(pdf_files)}] 処理完了: {pdf_file.name}")
                    
            except Exception as e:
                self.logger.error(f"OCR&リネームエラー ({pdf_file.name}): {e}")
                print(f"OCR&リネームエラー ({pdf_file.name}): {e}")
        
        self.logger.info("OCR&リネーム処理完了")
    
    def pdf_to_jpeg(self, pdf_file):
        """PDFをJPEGに変換"""
        try:
            self.logger.info(f"PDF→JPEG変換開始: {pdf_file.name}")
            
            # PDF2Imageを使用してPDFをJPEGに変換
            images = convert_from_path(pdf_file, dpi=200, first_page=1, last_page=1)
            
            if images:
                jpeg_file = pdf_file.parent / f"{pdf_file.stem}_temp.jpg"
                images[0].save(jpeg_file, 'JPEG', quality=95)
                self.logger.info(f"JPEG変換成功: {jpeg_file.name} (DPI: 200, Quality: 95)")
                return jpeg_file
            else:
                raise Exception("PDF変換失敗: 画像が生成されませんでした")
                
        except Exception as e:
            self.logger.error(f"PDF→JPEG変換エラー ({pdf_file.name}): {e}")
            raise Exception(f"PDF→JPEG変換エラー: {str(e)}")
    
    def perform_ocr(self, image_file):
        """Google Cloud Vision OCRを実行"""
        try:
            self.logger.info(f"OCR処理開始: {image_file.name}")
            
            # Google Cloud Vision クライアント初期化
            client = vision.ImageAnnotatorClient()
            
            # 画像ファイルを読み込み
            with open(image_file, 'rb') as image_file_obj:
                content = image_file_obj.read()
            
            image = vision.Image(content=content)
            
            # OCR実行
            response = client.text_detection(image=image)
            texts = response.text_annotations
            
            if texts:
                ocr_result = texts[0].description
                self.logger.info(f"OCR成功: {len(ocr_result)}文字のテキストを抽出")
                return ocr_result
            else:
                self.logger.warning(f"OCR結果なし: {image_file.name}")
                return ""
                
        except Exception as e:
            self.logger.error(f"OCR処理エラー ({image_file.name}): {e}")
            raise Exception(f"OCR処理エラー: {str(e)}")
    
    def generate_filename_with_ai(self, ocr_text):
        """OpenAI APIを使用してファイル名を生成"""
        try:
            self.logger.info(f"AI ファイル名生成開始 (入力: {len(ocr_text)}文字)")
            
            # OpenAI APIキーを環境変数から取得
            api_key = os.getenv('OPENAI_API_KEY')
            if not api_key:
                self.logger.error("OPENAI_API_KEYが設定されていません")
                raise Exception("OPENAI_API_KEYが設定されていません")
            
            client = OpenAI(api_key=api_key)
            
            prompt = f"""
以下のOCRテキストを分析して、適切なファイル名を生成してください。

OCRテキスト:
{ocr_text}

要件:
- ファイル名は50文字以内
- 日本語可
- ファイルシステムで使用できない文字は使用しない（/, \\, :, *, ?, ", <, >, |）
- 内容を表す簡潔で分かりやすい名前
- 日付が含まれている場合は含める
- 会社名や取引先名が分かる場合は含める

ファイル名のみを回答してください（拡張子は不要）:
"""
            
            self.logger.info(f"OpenAI API呼び出し開始 (model: gpt-3.5-turbo)")
            
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "あなたは文書の内容から適切なファイル名を生成するアシスタントです。"},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=100,
                temperature=0.3
            )
            
            filename = response.choices[0].message.content.strip()
            self.logger.info(f"AI生成ファイル名: {filename}")
            
            # ファイルシステムで使用できない文字を除去
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
            
        except Exception as e:
            self.logger.error(f"AI ファイル名生成エラー: {e}")
            print(f"AI ファイル名生成エラー: {e}")
            return None
    
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