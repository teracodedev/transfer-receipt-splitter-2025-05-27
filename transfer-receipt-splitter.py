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
import sys
try:
    from PyPDF2 import PdfReader, PdfWriter
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

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
        
        # パフォーマンス設定
        self.max_workers = min(4, os.cpu_count() or 4)  # 並列処理数制限
        
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
        try:
            # 実行中のPythonファイルと同じディレクトリにログファイルを作成
            script_dir = Path(__file__).parent if '__file__' in globals() else Path.cwd()
            log_file = script_dir / "transfer-receipt-splitter.log"
            
            # 前回のログファイルを削除
            if log_file.exists():
                try:
                    log_file.unlink()
                except Exception as e:
                    print(f"ログファイル削除エラー: {e}")
            
            # ログ設定
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_file, encoding='utf-8', mode='w'),
                    logging.StreamHandler()  # コンソールにも出力
                ]
            )
            
            self.logger = logging.getLogger(__name__)
            self.logger.info("=" * 50)
            self.logger.info("Transfer Receipt Splitter 開始")
            self.logger.info(f"ログファイル: {log_file}")
            self.logger.info(f"Python バージョン: {sys.version}")
            self.logger.info(f"作業ディレクトリ: {Path.cwd()}")
            self.logger.info("=" * 50)
            
        except Exception as e:
            print(f"ログ設定エラー: {e}")
            # フォールバック: 基本的なログ設定
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s'
            )
            self.logger = logging.getLogger(__name__)
            self.logger.error(f"ログ設定失敗: {e}")
    
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
        self.extract_option.trace_add('write', self.save_settings)
        self.overwrite_var.trace_add('write', self.save_settings)
        self.split_pdf_var.trace_add('write', self.save_settings)
        self.folder_path.trace_add('write', self.save_folder_setting)
    
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
        try:
            folder = Path(self.folder_path.get())
            zip_files = list(folder.glob("*.zip"))
            total_files = len(zip_files)
            
            if total_files == 0:
                self.safe_update_ui(lambda: self.progress_var.set("ZIPファイルが見つかりませんでした。"))
                self.safe_update_ui(lambda: self.extract_button.config(state="normal"))
                return
            
            # プログレスバーの設定
            self.safe_update_ui(lambda: self.progress_bar.config(maximum=total_files, value=0))
            
            success_count = 0
            error_files = []
            
            start_time = time.time()
            
            for i, zip_file in enumerate(zip_files):
                try:
                    # 進捗表示の更新
                    self.safe_update_ui(lambda f=zip_file, idx=i, total=total_files: 
                                      self.progress_var.set(f"解凍中: {f.name} ({idx+1}/{total})"))
                    
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
                        self.safe_update_ui(lambda f=zip_file, idx=i, total=total_files: 
                                          self.progress_var.set(f"PDF分割中: {f.name} ({idx+1}/{total})"))
                        self.split_pdfs_in_folder_optimized(extract_path)
                    
                    success_count += 1
                    
                except Exception as e:
                    error_msg = f"{zip_file.name}: {str(e)}"
                    error_files.append(error_msg)
                    self.logger.error(f"ZIP処理エラー: {error_msg}")
                    print(f"エラー詳細: {e}")  # コンソールにも出力
                
                # プログレスバーの更新
                self.safe_update_ui(lambda idx=i: self.progress_bar.config(value=idx+1))
            
            # 処理時間ログ
            elapsed_time = time.time() - start_time
            self.logger.info(f"全ZIP処理完了: {elapsed_time:.1f}秒")
            
            # 完了メッセージ
            if error_files:
                error_msg = "\n".join(error_files)
                self.safe_update_ui(lambda: messagebox.showwarning("警告", 
                                     f"解凍完了: {success_count}/{total_files}\n\n"
                                     f"エラーが発生したファイル:\n{error_msg}"))
            else:
                features = []
                if self.split_pdf_var.get() and PDF_AVAILABLE:
                    features.append("PDF分割")
                
                feature_text = "と" + "・".join(features) if features else ""
                self.safe_update_ui(lambda: messagebox.showinfo("完了", 
                                  f"すべてのZIPファイル({total_files}個)の解凍{feature_text}が完了しました。\n"
                                  f"処理時間: {elapsed_time:.1f}秒"))
            
            # UI状態をリセット
            self.safe_update_ui(lambda: self.progress_var.set("解凍完了"))
            self.safe_update_ui(lambda: self.extract_button.config(state="normal"))
            self.safe_update_ui(lambda: self.progress_bar.config(value=0))
            
        except Exception as e:
            error_msg = f"解凍処理で予期しないエラーが発生しました: {str(e)}"
            self.logger.error(error_msg)
            print(f"致命的エラー: {e}")  # コンソールにも出力
            self.safe_update_ui(lambda: messagebox.showerror("エラー", error_msg))
            self.safe_update_ui(lambda: self.extract_button.config(state="normal"))
            self.safe_update_ui(lambda: self.progress_var.set("エラーが発生しました"))
    
    def safe_update_ui(self, update_func):
        """スレッドセーフなUI更新"""
        try:
            if threading.current_thread() is threading.main_thread():
                update_func()
            else:
                self.root.after(0, update_func)
        except Exception as e:
            print(f"UI更新エラー: {e}")
            self.logger.error(f"UI更新エラー: {e}")
    
    def split_pdfs_in_folder_optimized(self, folder_path):
        """最適化されたPDF分割処理"""
        try:
            pdf_files = list(folder_path.glob("*.pdf"))
            split_files = []
            
            if not pdf_files:
                self.logger.info("PDF分割: PDFファイルが見つかりませんでした")
                return split_files
            
            self.logger.info(f"PDF分割開始: {len(pdf_files)}個のPDFファイル")
            
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
                        error_msg = f"PDF分割エラー ({pdf_file.name}): {e}"
                        self.logger.error(error_msg)
                        print(f"PDF分割エラー: {error_msg}")  # コンソールにも出力
            
            self.logger.info(f"PDF分割完了: {len(split_files)}個のファイルに分割")
            return split_files
            
        except Exception as e:
            error_msg = f"PDF分割処理で予期しないエラー: {e}"
            self.logger.error(error_msg)
            print(f"PDF分割エラー: {error_msg}")  # コンソールにも出力
            return []
    
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
    
    def save_settings(self, *args):
        """設定を.envファイルに保存"""
        try:
            self.update_env_file({
                'EXTRACT_OPTION': str(self.extract_option.get()),
                'OVERWRITE_FILES': str(self.overwrite_var.get()),
                'SPLIT_PDF': str(self.split_pdf_var.get())
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
            ui_settings = ['EXTRACT_OPTION', 'OVERWRITE_FILES', 'SPLIT_PDF']
            for key in ui_settings:
                if key in existing_settings:
                    f.write(f"{key}={existing_settings[key]}\n")

def main():
    root = tk.Tk()
    app = ZipExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
                    