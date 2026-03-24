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
        self.title("AI 考卷自動評分系統 v6.0")
        self.geometry("800x700")
        self.reader = None 
        self.answer_key = {}
        self.setup_ui()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        self.label = ctk.CTkLabel(self, text="自動閱卷系統 (路徑修正版)", font=("Microsoft JhengHei", 24, "bold"))
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
            
            # 解析 Word 表格：尋找含有 A, B, C 的聽力答案區
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text = cell.text.strip().upper()
                        # 排除掉題號或說明，只抓單一字母答案
                        if text in ['A', 'B', 'C']:
                            temp_ans_list.append(text)
            
            # 建立題號索引
            for i, ans in enumerate(temp_ans_list):
                self.answer_key[i + 1] = ans
                
            if self.answer_key:
                self.status_var.set(f"成功！載入 {len(self.answer_key)} 題聽力解答")
                messagebox.showinfo("解析成功", f"從 Word 提取了 {len(self.answer_key)} 個答案。")
            else:
                messagebox.showwarning("解析失敗", "Word 表格中找不到 A, B, C 格式的答案。")
        except Exception as e:
            messagebox.showerror("Word 錯誤", f"讀取失敗：{str(e)}")

    def get_poppler_path(self):
        """核心修正：針對使用者提供的路徑進行偵測"""
        # 使用者提供的精確路徑
        user_specified_path = r"C:\Users\ROC\Downloads\AutoGrader_v4\poppler\Library\bin"
        
        # 備用路徑 (相對於程式目錄)
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        path_options = [
            user_specified_path,
            os.path.join(base_dir, "poppler", "Library", "bin"),
            os.path.join(base_dir, "poppler", "bin")
        ]
        
        for p in path_options:
            if os.path.exists(os.path.join(p, "pdftoppm.exe")):
                return p
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
                self.log("[錯誤] 找不到 Poppler 執行檔！")
                self.log(f"請確認路徑是否存在: C:\\Users\\ROC\\Downloads\\AutoGrader_v4\\poppler\\Library\\bin")
                return

            self.status_var.set("正在初始化 AI 模型...")
            if self.reader is None:
                # 下載模型需連網，之後可離線
                self.reader = easyocr.Reader(['ch_tra', 'en'])

            self.status_var.set("正在辨識考卷 (這可能需要 1-2 分鐘)...")
            self.log("正在將 PDF 轉為圖片...")
            
            # 使用修正後的路徑
            pages = convert_from_path(pdf_path, poppler_path=p_path)
            
            all_text = ""
            for i, page in enumerate(pages):
                self.log(f"正在分析第 {i+1} 頁考卷內容...")
                img_np = np.array(page)
                res = self.reader.readtext(img_np, detail=0)
                all_text += "".join(res).upper()

            # 移除所有干擾字符
            clean_text = all_text.replace(" ", "").replace(".", "").replace(":", "").replace("-", "")

            correct = 0
            report = "\n--- 聽力測驗比對結果 ---\n"
            for q_num, ans in self.answer_key.items():
                pattern = f"{q_num}{ans}"
                # 如果 OCR 辨識字串中包含 "1A" 則判定對
                if pattern in clean_text:
                    correct += 1
                    report += f"題號 {q_num}: ✅ 正確 ({ans})\n"
                else:
                    report += f"題號 {q_num}: ❌ 錯誤 (正確解答: {ans})\n"
            
            report += f"\n====================\n總分：答對 {correct} 題 / 共 {len(self.answer_key)} 題"
            self.log(report)
            self.status_var.set("閱卷完成！")
        except Exception as e:
            self.log(f"[執行出錯] {str(e)}")
            self.status_var.set("發生錯誤")

if __name__ == "__main__":
    app = AutoGraderApp()
    app.mainloop()
