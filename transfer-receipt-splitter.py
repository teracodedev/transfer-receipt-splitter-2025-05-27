import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import os
import threading
from pathlib import Path
from dotenv import load_dotenv

class ZipExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ZIP ファイル一括解凍ツール")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # .envファイルを読み込み
        load_dotenv()
        
        # デフォルトフォルダの設定
        self.default_folder = self.get_default_folder()
        
        # 選択されたフォルダパス
        self.folder_path = tk.StringVar()
        
        # GUI要素の作成
        self.create_widgets()
        
        # デフォルトフォルダを設定して自動検索
        if self.default_folder:
            self.folder_path.set(str(self.default_folder))
            self.scan_zip_files()
    
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
        self.extract_option = tk.IntVar(value=1)
        ttk.Radiobutton(options_frame, text="各ZIPファイルごとに個別フォルダを作成", 
                       variable=self.extract_option, value=1).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(options_frame, text="選択フォルダ内に直接解凍", 
                       variable=self.extract_option, value=2).grid(row=1, column=0, sticky=tk.W)
        
        # 上書きオプション
        self.overwrite_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="既存ファイルを上書きする", 
                       variable=self.overwrite_var).grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        
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
                
                # 解凍先の決定
                if self.extract_option.get() == 1:
                    # 各ZIPファイルごとに個別フォルダを作成
                    extract_path = folder / zip_file.stem
                    extract_path.mkdir(exist_ok=True)
                else:
                    # 選択フォルダ内に直接解凍
                    extract_path = folder
                
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
                
                success_count += 1
                
            except Exception as e:
                error_files.append(f"{zip_file.name}: {str(e)}")
            
            # プログレスバーの更新
            self.progress_bar.config(value=i+1)
            self.root.update_idletasks()
        
        # 完了メッセージ
        if error_files:
            error_msg = "\n".join(error_files)
            messagebox.showwarning("警告", 
                                 f"解凍完了: {success_count}/{total_files}\n\n"
                                 f"エラーが発生したファイル:\n{error_msg}")
        else:
            messagebox.showinfo("完了", f"すべてのZIPファイル({total_files}個)の解凍が完了しました。")
        
        # UI状態をリセット
        self.progress_var.set("解凍完了")
        self.extract_button.config(state="normal")
        self.progress_bar.config(value=0)

def main():
    root = tk.Tk()
    app = ZipExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()