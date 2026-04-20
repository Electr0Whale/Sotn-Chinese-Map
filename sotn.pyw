import tkinter as tk
from tkinter import messagebox, ttk
import csv
from pathlib import Path
import ctypes
import os

# ==================== 强制隐藏命令行黑窗（Windows 专属）====================
if os.name == 'nt':
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

class SotNTerminologyApp:
    def __init__(self):
        self.csv_file = "my_sotn_terms.csv"
        self.terminology = {}
        self.sources = []
        
        if not self.load_csv():
            return
        
        self.root = tk.Tk()
        self.root.title("恶魔城月下夜想曲 术语查询工具")
        self.root.geometry("580x560")
        self.root.minsize(580, 560)
        self.root.resizable(True, True)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(4, weight=1)
        
        tk.Label(self.root, text="SotN 术语查询 (英/日 → 中文)", 
                font=("微软雅黑", 16, "bold")).grid(row=0, column=0, pady=12, sticky="ew")
        
        tk.Label(self.root, text="输入原文（支持 Tab 补全）:", 
                font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", padx=20)
        
        self.entry = tk.Entry(self.root, width=65, font=("Consolas", 11))
        self.entry.grid(row=2, column=0, pady=8, padx=20, ipady=4, sticky="ew")
        self.entry.bind("<KeyRelease>", self.update_suggestions)
        self.entry.bind("<Tab>", self.complete_with_tab)
        self.entry.bind("<Return>", self.confirm_selection)
        
        tk.Label(self.root, text="候选匹配（↑↓ 切换 / 鼠标悬停查看翻译）:", 
                font=("微软雅黑", 10)).grid(row=3, column=0, sticky="w", padx=20)
        
        self.suggestion_list = tk.Listbox(self.root, height=10, font=("微软雅黑", 11))
        self.suggestion_list.grid(row=4, column=0, pady=5, padx=20, sticky="nsew")
        
        self.suggestion_list.bind("<<ListboxSelect>>", self.on_list_select)
        self.suggestion_list.bind("<Motion>", self.on_mouse_hover)
        self.suggestion_list.bind("<Double-Button-1>", self.on_double_click)
        
        self.translation_label = tk.Label(self.root, text="翻译：", 
                                        font=("微软雅黑", 14, "bold"), 
                                        fg="#0066cc", wraplength=520, justify="left")
        self.translation_label.grid(row=5, column=0, pady=8, padx=20, sticky="w")
        
        self.current_translation = ""
        
        btn_frame = tk.Frame(self.root)
        btn_frame.grid(row=6, column=0, pady=10, sticky="ew")
        btn_frame.columnconfigure((0,1,2), weight=1)
        
        self.topmost_var = tk.BooleanVar(value=False)
        self.topmost_check = tk.Checkbutton(btn_frame, text="窗口置顶", 
                                          variable=self.topmost_var, 
                                          command=self.toggle_topmost,
                                          font=("微软雅黑", 10))
        self.topmost_check.grid(row=0, column=0, padx=10)
        
        self.copy_btn = tk.Button(btn_frame, text="复制翻译", 
                                 command=self.copy_translation,
                                 font=("微软雅黑", 10), width=12, bg="#4CAF50", fg="white")
        self.copy_btn.grid(row=0, column=1, padx=10)
        
        self.edit_btn = tk.Button(btn_frame, text="编辑术语库（添加新词）", 
                                 command=self.open_edit_window,
                                 font=("微软雅黑", 10), width=18, bg="#2196F3", fg="white")
        self.edit_btn.grid(row=0, column=2, padx=10)
        
        self.entry.focus_set()
        self.root.mainloop()
    
    def load_csv(self):
        path = Path(self.csv_file)
        if not path.exists():
            messagebox.showerror("文件未找到", f"未找到 {self.csv_file}！请和程序放在同一文件夹")
            return False
        try:
            self.terminology.clear()
            self.sources.clear()
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    source = row.get('source', '').strip()
                    target = row.get('target', '').strip()
                    if source and target and source not in self.terminology:
                        self.terminology[source] = target
                        self.sources.append(source)
            self.sources.sort()
            return True
        except Exception as e:
            messagebox.showerror("加载失败", f"读取 CSV 失败：{str(e)}")
            return False
    
    def refresh_data(self):
        if self.load_csv():
            self.update_suggestions()
    
    def toggle_topmost(self):
        self.root.attributes('-topmost', self.topmost_var.get())
    
    def copy_translation(self):
        if self.current_translation:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.current_translation)
            messagebox.showinfo("成功", "翻译已复制到剪贴板！")
        else:
            messagebox.showwarning("提示", "当前没有翻译可复制")
    
    def update_suggestions(self, event=None):
        current = self.entry.get().strip()
        self.suggestion_list.delete(0, tk.END)
        self.current_translation = ""
        self.translation_label.config(text="翻译：")
        if not current: return
        matches = [s for s in self.sources if s.lower().startswith(current.lower())]
        for match in matches[:30]:
            self.suggestion_list.insert(tk.END, match)
        if matches:
            self.suggestion_list.selection_set(0)
            self.show_translation_for_index(0)
    
    # 以下函数保持不变（省略以节省篇幅，但实际代码已包含）
    def on_mouse_hover(self, event): 
        index = self.suggestion_list.nearest(event.y)
        if 0 <= index < self.suggestion_list.size():
            self.suggestion_list.selection_clear(0, tk.END)
            self.suggestion_list.selection_set(index)
            self.show_translation_for_index(index)
    
    def on_list_select(self, event=None):
        selection = self.suggestion_list.curselection()
        if selection:
            idx = selection[0]
            source = self.suggestion_list.get(idx)
            self.entry.delete(0, tk.END)
            self.entry.insert(0, source)
            self.show_translation_for_index(idx)
    
    def on_double_click(self, event=None):
        self.on_list_select()
    
    def complete_with_tab(self, event=None):
        if self.suggestion_list.size() > 0:
            first = self.suggestion_list.get(0)
            self.entry.delete(0, tk.END)
            self.entry.insert(0, first)
            self.show_translation_for_index(0)
        return "break"
    
    def confirm_selection(self, event=None):
        current = self.entry.get().strip()
        if current in self.terminology:
            trans = self.terminology[current]
            self.current_translation = trans
            self.translation_label.config(text=f"翻译：{trans}")
    
    def show_translation_for_index(self, index):
        source = self.suggestion_list.get(index)
        if source in self.terminology:
            trans = self.terminology[source]
            self.current_translation = trans
            self.translation_label.config(text=f"翻译：{trans}")
        else:
            self.current_translation = ""
            self.translation_label.config(text="翻译：无匹配")
    
    def open_edit_window(self):
        edit_win = tk.Toplevel(self.root)
        edit_win.title("添加新术语")
        edit_win.geometry("420x220")
        edit_win.resizable(False, False)
        edit_win.grab_set()
        
        tk.Label(edit_win, text="原文（英文或日文）:", font=("微软雅黑", 10)).pack(anchor="w", padx=20, pady=(15,5))
        source_entry = tk.Entry(edit_win, width=50, font=("Consolas", 11))
        source_entry.pack(padx=20, ipady=4)
        
        tk.Label(edit_win, text="中文翻译:", font=("微软雅黑", 10)).pack(anchor="w", padx=20, pady=(10,5))
        target_entry = tk.Entry(edit_win, width=50, font=("微软雅黑", 11))
        target_entry.pack(padx=20, ipady=4)
        
        def add_term():
            source = source_entry.get().strip()
            target = target_entry.get().strip()
            if not source or not target:
                messagebox.showwarning("错误", "原文和翻译都不能为空！")
                return
            try:
                with open(self.csv_file, 'a', encoding='utf-8', newline='') as f:
                    csv.writer(f).writerow([source, target, 'zh-CN'])
                self.terminology[source] = target
                if source not in self.sources:
                    self.sources.append(source)
                    self.sources.sort()
                messagebox.showinfo("成功", f"已添加：\n{source} → {target}")
                self.refresh_data()
                edit_win.destroy()
            except Exception as e:
                messagebox.showerror("失败", f"保存失败：{str(e)}")
        
        tk.Button(edit_win, text="添加并保存", command=add_term, 
                 font=("微软雅黑", 11), bg="#4CAF50", fg="white", width=15).pack(pady=20)

if __name__ == "__main__":
    app = SotNTerminologyApp()