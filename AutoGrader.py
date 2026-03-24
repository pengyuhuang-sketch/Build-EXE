import customtkinter as ctk
from tkinter import filedialog, messagebox
from docx import Document
import easyocr
import os
import threading
from pdf2image import convert_from_path
import numpy as np
import sys

class AutoGraderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 考卷自動評分系統 v4.0")
        self.geometry("800x700")
        
        self.reader = None 
        self.answer_key = {}
        self.setup_ui()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        self.label = ctk.CTkLabel(self, text="自動閱卷系統 (穩定版)", font=("Microsoft JhengHei", 24, "bold"))
        self.label.pack(pady=20)

        self.btn_frame = ctk.CTkFrame(self)
        self.btn_frame.pack(pady=10, padx=20, fill="x")

        self.btn_ans = ctk.CTkButton(self.btn_frame, text="1. 載入 Word 解答檔案", command=self.load_word_answers)
        self.btn_ans.grid(row=0, column=0, padx=20, pady=20)

        self.btn_exam = ctk.CTkButton(self.btn_frame, text="2. 選擇 PDF 考卷評分", command=self.start_grading, fg_color="#2ecc71")
        self.btn_exam.grid(row=0, column=1, padx=20, pady=20)

        self.status_var = ctk.StringVar(value="狀態：請先載入解答")
        self.status_label = ctk.CTkLabel(self, textvariable=self.status_var, font=("Microsoft JhengHei", 14))
        self.status_label.pack()

        self.result_box = ctk.CTkTextbox(self, width=700, height=400, font=("Consolas", 14))
        self.result_box.pack(pady=20, padx=20)

    def log(self, text):
        self.result_box.insert("end", f"{text}\n")
        self.result_box.see("end")

    def load_word_answers(self):
        path = filedialog.askopenfilename(filetypes=[("Word 檔案", "*.docx")])
        if not path: return
        try:
            doc = Document(path)
            self.answer_key = {}
            temp_ans_list = []
            
            # 遍歷所有表格抓取答案字母
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text = cell.text.strip().upper()
                        # 只抓取單獨的 A, B, C 字樣
                        if text in ['A', 'B', 'C']:
                            temp_ans_list.append(text)
            
            for i, ans in enumerate(temp_ans_list):
                self.answer_key[i + 1] = ans
                
            if self.answer_key:
                self.status_var.set(f"成功！載入 {len(self.answer_key)} 題解答")
                messagebox.showinfo("解析成功", f"從 Word 提取了 {len(self.answer_key)} 個答案。")
            else:
                messagebox.showwarning("解析失敗", "Word 表格中找不到 A, B, C 格式的答案。")
        except Exception as e:
            messagebox.showerror("Word 錯誤", f"讀取失敗：{str(e)}")

    def get_poppler_path(self):
        """自動偵測 Poppler 位置"""
        # 獲取程式執行的路徑 (相對於 .exe 或 .py)
        if getattr(sys, 'frozen', False):
            base_dir = sys._MEIPASS # PyInstaller 臨時目錄
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        # 優先檢查程式目錄下的 poppler/bin
        path_options = [
            os.path.join(base_dir, "poppler", "bin"),
            os.path.join(base_dir, "poppler", "Library", "bin"),
            r"C:\ProgramData\chocolatey\lib\poppler\tools\bin" # 這是 GitHub Actions 常見路徑
        ]
        
        for p in path_options:
            if os.path.exists(p): return p
        return None

    def start_grading(self):
        if not self.answer_key:
            messagebox.showwarning("提示", "請先載入正確解答！")
            return
        path = filedialog.askopenfilename(filetypes=[("PDF 檔案", "*.pdf")])
        if path:
            self.result_box.delete("1.0", "end")
            threading.Thread(target=self.grading_process, args=(path,), daemon=True).start()

    def grading_process(self, pdf_path):
        try:
            p_path = self.get_poppler_path()
            if not p_path:
                self.log("[錯誤] 找不到 Poppler 路徑，請確保 poppler 資料夾與程式在同一目錄。")
                return

            self.status_var.set("正在載入 AI 模型並辨識中...")
            if self.reader is None:
                self.reader = easyocr.Reader(['ch_tra', 'en'])

            self.log("正在將 PDF 轉為圖片...")
            pages = convert_from_path(pdf_path, poppler_path=p_path)
            
            all_text = ""
            for i, page in enumerate(pages):
                self.log(f"分析第 {i+1} 頁...")
                img_np = np.array(page)
                res = self.reader.readtext(img_np, detail=0)
                all_text += "".join(res).upper()

            # 移除所有干擾字符
            clean_text = all_text.replace(" ", "").replace(".", "").replace(":", "").replace("-", "")

            correct = 0
            report = "--- 聽力測驗比對結果 ---\n"
            for q_num, ans in self.answer_key.items():
                pattern = f"{q_num}{ans}"
                if pattern in clean_text:
                    correct += 1
                    report += f"題號 {q_num}: ✅ 正確 ({ans})\n"
                else:
                    report += f"題號 {q_num}: ❌ 錯誤 (正確解答: {ans})\n"
            
            report += f"\n總分：{correct} / {len(self.answer_key)}"
            self.log(report)
            self.status_var.set("評分完成")
        except Exception as e:
            self.log(f"[執行出錯] {str(e)}")
            self.status_var.set("發生錯誤")

if __name__ == "__main__":
    app = AutoGraderApp()
    app.mainloop()
