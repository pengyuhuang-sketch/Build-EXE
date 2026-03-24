import customtkinter as ctk
from tkinter import filedialog, messagebox
from docx import Document
import easyocr
import os
import threading
from pdf2image import convert_from_path # 需要 poppler
import numpy as np
from PIL import Image

class AutoGraderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 考卷評分系統 (Word/PDF 版)")
        self.geometry("700x600")
        self.reader = None 
        self.answer_data = [] # 改用 List 儲存字典
        self.setup_ui()

    def setup_ui(self):
        ctk.CTkLabel(self, text="自動閱卷工具 v2.0", font=("Microsoft JhengHei", 24, "bold")).pack(pady=20)
        
        self.btn_ans = ctk.CTkButton(self, text="1. 載入正確解答 (Word)", command=self.load_word_answers)
        self.btn_ans.pack(pady=10)

        self.btn_exam = ctk.CTkButton(self, text="2. 選擇學生考卷 (PDF)", command=self.start_processing_thread, fg_color="green")
        self.btn_exam.pack(pady=10)

        self.result_box = ctk.CTkTextbox(self, width=600, height=350)
        self.result_box.pack(pady=20, padx=20)

    def load_word_answers(self):
        path = filedialog.askopenfilename(filetypes=[("Word 檔案", "*.docx")])
        if path:
            try:
                doc = Document(path)
                self.answer_data = []
                # 假設解答在 Word 的第一個表格中
                table = doc.tables[0]
                for i, row in enumerate(table.rows):
                    if i == 0: continue # 跳過標題列
                    cells = [cell.text.strip() for cell in row.cells]
                    # 格式：大題, 小題, 答案, 分數
                    self.answer_data.append({
                        "id": f"{cells[0]}-{cells[1]}",
                        "ans": cells[2],
                        "score": int(cells[3])
                    })
                messagebox.showinfo("成功", f"已載入 {len(self.answer_data)} 題解答")
            except Exception as e:
                messagebox.showerror("錯誤", f"Word 解析失敗: {e}")

    def process_pdf(self, pdf_path):
        try:
            self.result_box.delete("1.0", "end")
            if self.reader is None:
                self.reader = easyocr.Reader(['ch_tra', 'en'])

            # 將 PDF 每一頁轉為圖片進行 OCR
            self.result_box.insert("end", "正在轉換 PDF 並辨識文字...\n")
            pages = convert_from_path(pdf_path)
            full_text = ""
            
            for i, page in enumerate(pages):
                self.result_box.insert("end", f"正在辨識第 {i+1} 頁...\n")
                # 將 PIL Image 轉為 numpy array 給 easyocr
                img_array = np.array(page)
                results = self.reader.readtext(img_array, detail=0)
                full_text += "".join(results)

            # 比對邏輯
            total = 0
            report = "\n--- 評分結果 ---\n"
            for item in self.answer_data:
                if item["ans"] in full_text:
                    total += item["score"]
                    report += f"✅ {item['id']}: 正確 (+{item['score']})\n"
                else:
                    report += f"❌ {item['id']}: 錯誤 (解答: {item['ans']})\n"
            
            report += f"\n總分：{total} 分"
            self.result_box.insert("end", report)
        except Exception as e:
            messagebox.showerror("錯誤", f"處理失敗: {e}")

    def start_processing_thread(self):
        path = filedialog.askopenfilename(filetypes=[("PDF 檔案", "*.pdf")])
        if path: threading.Thread(target=self.process_pdf, args=(path,), daemon=True).start()

if __name__ == "__main__":
    app = AutoGraderApp()
    app.mainloop()
