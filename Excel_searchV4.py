import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
from datetime import datetime
import threading

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

class DataExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("貸付コード抽出アプリ")
        self.db_df = None
        self.code_df = None
        self.output_dir = os.getcwd()
        self.dup_dir = os.getcwd()
        self.column_vars = []
        self.build_ui()

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        scroll_frame = ScrollableFrame(main_frame)
        scroll_frame.pack(fill="both", expand=True)
        frame = scroll_frame.scrollable_frame

        ttk.Button(frame, text="① DBファイルを開く", command=self.load_db).grid(row=0, column=0, sticky="w")
        ttk.Button(frame, text="② 検索ファイルを開く", command=self.load_code).grid(row=1, column=0, sticky="w")

        ttk.Button(frame, text="③ 抽出保存フォルダを選ぶ", command=self.choose_output_dir).grid(row=2, column=0, sticky="w")
        self.output_dir_label = ttk.Label(frame, text=self.output_dir, foreground="blue", wraplength=400)
        self.output_dir_label.grid(row=2, column=1, sticky="w")

        ttk.Button(frame, text="④ 重複保存フォルダを選ぶ", command=self.choose_dup_dir).grid(row=3, column=0, sticky="w")
        self.dup_dir_label = ttk.Label(frame, text=self.dup_dir, foreground="blue", wraplength=400)
        self.dup_dir_label.grid(row=3, column=1, sticky="w")

        ttk.Label(frame, text="⑤ DB側の『貸付コード』列:").grid(row=4, column=0, sticky="w")
        self.code_col_combo = ttk.Combobox(frame, state="readonly", width=40)
        self.code_col_combo.grid(row=4, column=1)

        for i in range(40):
            ttk.Label(frame, text=f"抽出列{i+1}:").grid(row=5+i, column=0, sticky="e")
            cb = ttk.Combobox(frame, state="readonly", width=40)
            cb.grid(row=5+i, column=1)
            self.column_vars.append(cb)

        ttk.Button(frame, text="⑥ 抽出開始", command=self.extract_data).grid(row=45, column=0, columnspan=2, pady=10)

    def show_loading_dialog(self, message="読み込み中..."):
        self.loading_window = tk.Toplevel(self.root)
        self.loading_window.title("処理中")
        self.loading_window.geometry("300x100")
        self.loading_window.transient(self.root)
        self.loading_window.grab_set()
        ttk.Label(self.loading_window, text=message).pack(expand=True, pady=30)
        self.loading_window.update()

    def close_loading_dialog(self):
        if hasattr(self, 'loading_window'):
            self.loading_window.destroy()
            del self.loading_window

    def select_sheets(self, sheet_names):
        top = tk.Toplevel(self.root)
        top.title("シートを選択")
        top.geometry("300x300")
        top.grab_set()
        listbox = tk.Listbox(top, selectmode=tk.MULTIPLE)
        for name in sheet_names:
            listbox.insert(tk.END, name)
        listbox.pack(fill=tk.BOTH, expand=True)

        selected = []

        def confirm():
            indices = listbox.curselection()
            for i in indices:
                selected.append(sheet_names[i])
            top.destroy()

        ttk.Button(top, text="OK", command=confirm).pack(pady=5)
        self.root.wait_window(top)
        return selected

    def load_excel_with_selection(self, path):
        xl = pd.ExcelFile(path)
        if len(xl.sheet_names) > 1:
            selected_sheets = self.select_sheets(xl.sheet_names)
        else:
            selected_sheets = xl.sheet_names

        df_list = [xl.parse(sheet_name, dtype=str) for sheet_name in selected_sheets]
        return pd.concat(df_list, ignore_index=True)

    def load_db_thread(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.root.after(0, lambda: self.show_loading_dialog("DBファイル読み込み中..."))

        try:
            df = self.load_excel_with_selection(path)
        except Exception as e:
            self.root.after(0, lambda: [self.close_loading_dialog(), messagebox.showerror("エラー", f"読み込み失敗：{e}")])
            return

        def update_ui():
            self.close_loading_dialog()
            self.db_df = df
            col_names = list(self.db_df.columns)
            self.code_col_combo['values'] = col_names
            for cb in self.column_vars:
                cb['values'] = col_names
            messagebox.showinfo("完了", "DBファイルを読み込みました。")

        self.root.after(0, update_ui)

    def load_code_thread(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.root.after(0, lambda: self.show_loading_dialog("検索ファイル読み込み中..."))

        try:
            df = self.load_excel_with_selection(path)
        except Exception as e:
            self.root.after(0, lambda: [self.close_loading_dialog(), messagebox.showerror("エラー", f"読み込み失敗：{e}")])
            return

        def update_ui():
            self.close_loading_dialog()
            self.code_df = df
            messagebox.showinfo("完了", "検索用ファイルを読み込みました。")

        self.root.after(0, update_ui)

    def load_db(self):
        threading.Thread(target=self.load_db_thread, daemon=True).start()

    def load_code(self):
        threading.Thread(target=self.load_code_thread, daemon=True).start()

    def extract_data_thread(self):
        self.root.after(0, lambda: self.show_loading_dialog("抽出中..."))

        if self.db_df is None or self.code_df is None:
            self.root.after(0, lambda: [self.close_loading_dialog(), messagebox.showerror("エラー", "両方のファイルを読み込んでください。")])
            return

        code_col = self.code_col_combo.get()
        if not code_col:
            self.root.after(0, lambda: [self.close_loading_dialog(), messagebox.showerror("エラー", "貸付コード列を選択してください。")])
            return

        code_list = self.code_df.iloc[:, 0].astype(str).tolist()
        db_df = self.db_df.copy()
        db_df[code_col] = db_df[code_col].astype(str)

        selected_columns = [code_col]
        for cb in self.column_vars:
            val = cb.get()
            if val and val not in selected_columns:
                selected_columns.append(val)

        extracted = db_df[db_df[code_col].isin(code_list)]
        result = extracted[selected_columns]
        date_str = datetime.now().strftime("%Y%m%d")
        result_path = os.path.join(self.output_dir, f"抽出結果_{date_str}.xlsx")
        result.to_excel(result_path, index=False)

        dup_df = result[result.duplicated(subset=[code_col], keep=False)]
        dup_message = ""
        if not dup_df.empty:
            dup_path = os.path.join(self.dup_dir, f"重複結果_{date_str}.xlsx")
            dup_df.to_excel(dup_path, index=False)
            dup_message = f"\n重複データ {len(dup_df)} 件は:\n{dup_path} に保存しました。"

        self.root.after(0, lambda: [
            self.close_loading_dialog(),
            messagebox.showinfo("完了", f"{len(result)} 件 抽出しました。\n抽出保存先:\n{result_path}{dup_message}")
        ])

    def extract_data(self):
        threading.Thread(target=self.extract_data_thread, daemon=True).start()

    def choose_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir = path
            self.output_dir_label.config(text=path)

    def choose_dup_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.dup_dir = path
            self.dup_dir_label.config(text=path)

if __name__ == "__main__":
    root = tk.Tk()
    app = DataExtractorApp(root)
    root.mainloop()
