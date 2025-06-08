import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import openpyxl
except ImportError:
    openpyxl = None

class ExcelApp:
    def __init__(self, root):
        self.root = root
        root.title("Excel TC Sorgu")

        self.filepath = None
        self.workbook = None

        self.load_button = tk.Button(root, text="Excel Yükle", command=self.load_excel)
        self.load_button.pack(pady=5)

        self.tc_label = tk.Label(root, text="TC Kimlik Numarası:")
        self.tc_label.pack()
        self.tc_entry = tk.Entry(root)
        self.tc_entry.pack(pady=5)

        self.search_button = tk.Button(root, text="Sorgula", command=self.search_tc)
        self.search_button.pack(pady=5)

        self.result_tc = tk.Label(root, text="")
        self.result_tc.pack()
        self.result_ad = tk.Label(root, text="")
        self.result_ad.pack()
        self.result_kan = tk.Label(root, text="")
        self.result_kan.pack()

    def load_excel(self):
        if openpyxl is None:
            messagebox.showerror("Eksik Paket", "openpyxl yüklü değil")
            return
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.workbook = openpyxl.load_workbook(path)
            self.filepath = path
            messagebox.showinfo("Başarılı", "Excel yüklendi")
        except Exception as e:
            messagebox.showerror("Hata", str(e))

    def search_tc(self):
        if openpyxl is None:
            messagebox.showerror("Eksik Paket", "openpyxl yüklü değil")
            return
        if not self.workbook:
            messagebox.showwarning("Uyarı", "Önce bir Excel dosyası yükleyin")
            return
        tc = self.tc_entry.get().strip()
        if not tc:
            messagebox.showwarning("Uyarı", "TC kimlik numarası girin")
            return
        sheet1 = self.workbook.worksheets[0]
        record = None
        for row in sheet1.iter_rows(min_row=2, values_only=True):
            tc_value, adsoyad, kan = row[:3]
            if str(tc_value) == tc:
                record = (tc_value, adsoyad, kan)
                break
        if not record:
            messagebox.showinfo("Sonuç Yok", "TC bulunamadı")
            return
        tc_value, adsoyad, kan = record
        # Sheet2 güncelle
        if len(self.workbook.worksheets) > 1:
            sheet2 = self.workbook.worksheets[1]
            sheet2['A2'] = tc_value
            sheet2['B2'] = adsoyad
            sheet2['C2'] = kan
            try:
                self.workbook.save(self.filepath)
            except Exception:
                pass
        self.result_tc.config(text=f"TC: {tc_value}")
        self.result_ad.config(text=f"Ad Soyad: {adsoyad}")
        self.result_kan.config(text=f"Kan Grubu: {kan}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
