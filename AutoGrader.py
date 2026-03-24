import customtkinter as ctk
from tkinter import filedialog, messagebox
from docx import Document
import easyocr
import os
import threading
from pdf2image import convert_from_path
import numpy as np
import sys

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AutoGraderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 考卷評分系統 v3.0")
        self.geometry("800x700")
        
        self.reader = None 
        self.answer_key = {}
        # 注意：請確保此路徑正確，或將 poppler 放在程式目錄下
        self.poppler_path = r"C:\poppler\bin" 
        self.setup_ui()

    def setup_ui(self):
        self.label = ctk.CTkLabel(self, text="自動閱卷系統 (修正版)", font=("Microsoft JhengHei", 24, "bold"))
        self.label.pack(pady=20)

        self.btn_frame = ctk.CTkFrame(self)
        self.btn_frame.pack(pady=10, padx=20, fill="x")

        self.btn_ans = ctk.CTkButton(self.btn_frame, text="1. 載入 Word 解答", command=self.load_word_answers)
        self.btn_ans.grid(row=0, column=0, padx=20, pady=20)

        self.btn_exam = ctk.CTkButton(self.btn_frame, text="2. 選擇 PDF 考卷評分", command=self.start_grading, fg_color="#2ecc71")
        self.btn_exam.grid(row=0, column=1, padx=20, pady=20)

        self.status_var = ctk.StringVar(value="狀態：等待載入解答")
        ctk.CTkLabel(self, textvariable=self.status_var).pack()

        self.result_box = ctk.CTkTextbox(self, width=700, height=400, font=("Consolas", 14))
        self.result_box.pack(pady=20, padx=20)

    def load_word_answers(self):
        path = filedialog.askopenfilename(filetypes=[("Word 檔案", "*.docx")])
        if not path: return
        
        try:
            doc = Document(path)
            self.answer_key = {}
            temp_answers = []
            
            # 專門尋找包含聽力答案 (A, B, C) 的表格
            for table in doc.tables:
                for row in table.rows:
                    cells = [c.text.strip().upper() for c in row.cells]
                    # 如果這列包含 A, B 或 C，就認定它是答案列
                    for val in cells:
                        if val in ['A', 'B', 'C']:
                            temp_answers.append(val)
            
            # 將抓到的答案編號 (1, 2, 3...)
            for i, ans in enumerate(temp_answers):
                self.answer_key[i + 1] = ans
            
            if self.answer_key:
                self.status_var.set(f"成功！已載入 {len(self.answer_key)} 題聽力解答")
                messagebox.showinfo("解析成功", f"偵測到 {len(self.answer_key)} 個聽力答案。")
            else:
                raise ValueError("在 Word 表格中找不到 A, B, C 格式的答案。")

        except Exception as e:
            messagebox.showerror("Word 解析失敗", f"錯誤原因：{str(e)}")

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
            self.status_var.set("正在執行 OCR 辨識 (可能需要 1-2 分鐘)...")
            if self.reader is None:
                self.reader = easyocr.Reader(['ch_tra', 'en'])

            pages = convert_from_path(pdf_path, poppler_path=self.poppler_path)
            full_text = ""
            for page in pages:
                img_np = np.array(page)
                res = self.reader.readtext(img_np, detail=0)
                full_text += "".join(res).upper()

            # 清理文字以利比對
            clean_text = full_text.replace(" ", "").replace(".", "").replace(":", "")

            correct = 0
            report = "--- 評分報告 ---\n"
            for q_num, ans in self.answer_key.items():
                # 比對模式：例如 "1A", "2B"
                pattern = f"{q_num}{ans}"
                if pattern in clean_text:
                    correct += 1
                    report += f"第 {q_num} 題: ✅ 正確 ({ans})\n"
                else:
                    report += f"第 {q_num} 題: ❌ 錯誤 (正確解答: {ans})\n"
            
            report += f"\n總結：答對 {correct} 題 / 總計 {len(self.answer_key)} 題"
            self.result_box.insert("end", report)
            self.status_var.set("評分完成")
        except Exception as e:
            self.result_box.insert("end", f"\n[錯誤] 執行失敗: {str(e)}")
            self.status_var.set("執行發生錯誤")

if __name__ == "__main__":
    app = AutoGraderApp()
    app.mainloop()
