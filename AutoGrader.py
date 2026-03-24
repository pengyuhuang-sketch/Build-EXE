import sys

def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)
    
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import easyocr
import os
import threading

# 設定介面風格
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AutoGraderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 考卷自動評分系統 v1.0")
        self.geometry("700x550")
        
        self.reader = None 
        self.answer_data = None
        self.ans_path = ""
        self.setup_ui()

    def setup_ui(self):
        # 標題
        self.label_title = ctk.CTkLabel(self, text="自動閱卷與評分工具", font=("Microsoft JhengHei", 26, "bold"))
        self.label_title.pack(pady=20)

        # 按鈕區
        self.frame_btns = ctk.CTkFrame(self)
        self.frame_btns.pack(pady=10, padx=20, fill="x")

        self.btn_load_ans = ctk.CTkButton(self.frame_btns, text="1. 載入正確答案 (Excel/CSV)", command=self.load_answers)
        self.btn_load_ans.grid(row=0, column=0, padx=20, pady=20)

        self.btn_load_img = ctk.CTkButton(self.frame_btns, text="2. 辨識照片並評分", command=self.start_processing_thread, fg_color="#2ecc71", hover_color="#27ae60")
        self.btn_load_img.grid(row=0, column=1, padx=20, pady=20)

        # 狀態顯示
        self.status_label = ctk.CTkLabel(self, text="狀態：等待操作...", font=("Microsoft JhengHei", 12), text_color="gray")
        self.status_label.pack()

        # 結果輸出區
        self.result_box = ctk.CTkTextbox(self, width=600, height=300, font=("Consolas", 14))
        self.result_box.pack(pady=20, padx=20)

    def log(self, message):
        self.result_box.insert("end", f"{message}\n")
        self.result_box.see("end")

    def load_answers(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV 檔案", "*.xlsx *.csv")])
        if path:
            try:
                if path.endswith('.csv'):
                    self.answer_data = pd.read_csv(path)
                else:
                    self.answer_data = pd.read_excel(path)
                self.ans_path = path
                self.status_label.configure(text=f"狀態：已載入答案檔 {os.path.basename(path)}", text_color="#3498db")
                messagebox.showinfo("成功", "標準答案載入成功！")
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取失敗: {e}")

    def start_processing_thread(self):
        # 使用執行緒避免 UI 凍結
        if not self.ans_path:
            messagebox.showwarning("提示", "請先載入正確答案檔！")
            return
        
        img_path = filedialog.askopenfilename(filetypes=[("圖片檔案", "*.jpg *.png *.jpeg")])
        if img_path:
            threading.Thread(target=self.process_exam, args=(img_path,), daemon=True).start()

    def process_exam(self, img_path):
        try:
            self.result_box.delete("1.0", "end")
            self.status_label.configure(text="狀態：正在辨識中，請稍候...", text_color="#e67e22")
            
            if self.reader is None:
                self.log("正在載入 AI 模型 (首次啟動需約 30 秒)...")
                self.reader = easyocr.Reader(['ch_tra', 'en']) 

            self.log(f"分析檔案：{os.path.basename(img_path)}")
            
            # 辨識文字
            ocr_results = self.reader.readtext(img_path, detail=0)
            full_text = "".join(ocr_results)

            total_score = 0
            report = "\n--- 比對報告 ---\n"

            for index, row in self.answer_data.iterrows():
                q_id = f"{row['大題']}-{row['小題']}"
                ans = str(row['答案'])
                pts = row['分數']
                
                if ans in full_text:
                    total_score += pts
                    report += f"✅ {q_id}: 正確 (+{pts})\n"
                else:
                    report += f"❌ {q_id}: 錯誤 (正確答案: {ans})\n"

            report += f"\n====================\n最終得分統計：{total_score} 分"
            self.log(report)
            self.status_label.configure(text="狀態：比對完成！", text_color="#2ecc71")

        except Exception as e:
            self.log(f"發生錯誤: {e}")
            self.status_label.configure(text="狀態：辨識出錯", text_color="#e74c3c")

if __name__ == "__main__":
    app = AutoGraderApp()
    app.mainloop()
