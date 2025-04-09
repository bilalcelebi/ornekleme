import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import pandas as pd

class OrneklemOlusturucu:
    def __init__(self, root):
        self.root = root
        self.root.title('Örneklem Oluşturucu Uygulaması')
        self.root.geometry('800x600')

        self.dosya_yolu = None
        self.df = None
        self.orneklem_df = None

        self.create_widgets()

    def create_widgets(self):
        # Ana çerçeve
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Adım 1: Dosya Seçimi
        step1_frame = tk.Frame(main_frame)
        step1_frame.pack(fill=tk.X, pady=5)

        self.lbl_dosya = tk.Label(step1_frame, text='Excel dosyası seçin:', anchor='w')
        self.lbl_dosya.pack(side=tk.LEFT, padx=5)

        self.btn_dosya_sec = tk.Button(step1_frame, text='Dosya Seç', command=self.dosya_sec)
        self.btn_dosya_sec.pack(side=tk.RIGHT, padx=5)

        # Adım 2: Sayfa Seçimi
        step2_frame = tk.Frame(main_frame)
        step2_frame.pack(fill=tk.X, pady=5)

        self.lbl_sayfa = tk.Label(step2_frame, text='Sayfa seçin:', anchor='w')
        self.lbl_sayfa.pack(side=tk.LEFT, padx=5)

        self.cmb_sayfa_sec = ttk.Combobox(step2_frame, state='disabled')
        self.cmb_sayfa_sec.pack(side=tk.RIGHT, padx=5)
        self.cmb_sayfa_sec.bind("<<ComboboxSelected>>", self.sayfa_sec)

        # Adım 3: Örneklem Yüzdesi ve Oluşturma
        step3_frame = tk.Frame(main_frame)
        step3_frame.pack(fill=tk.X, pady=5)

        self.lbl_orneklem = tk.Label(step3_frame, text='Örneklem yüzdesi (%):', anchor='w')
        self.lbl_orneklem.pack(side=tk.LEFT, padx=5)

        self.ent_orneklem_yuzdesi = tk.Entry(step3_frame, state='disabled')
        self.ent_orneklem_yuzdesi.pack(side=tk.LEFT, padx=5)

        self.btn_orneklem_sec = tk.Button(step3_frame, text='Örneklem Oluştur', state=tk.DISABLED, command=self.orneklem_sec)
        self.btn_orneklem_sec.pack(side=tk.RIGHT, padx=5)

        # Sonuçlar
        self.result_frame = tk.Frame(main_frame)
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.result_listbox = tk.Listbox(self.result_frame, selectmode=tk.SINGLE, height=10, width=80)
        self.result_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.h_scrollbar = tk.Scrollbar(self.result_frame, orient=tk.HORIZONTAL, command=self.result_listbox.xview)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.result_listbox.config(xscrollcommand=self.h_scrollbar.set)

        self.v_scrollbar = tk.Scrollbar(self.result_frame, orient=tk.VERTICAL, command=self.result_listbox.yview)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.result_listbox.config(yscrollcommand=self.v_scrollbar.set)

    def dosya_sec(self):
        """Kullanıcıya dosya seçtirmek için dosya gezgini penceresi açar."""
        self.dosya_yolu = filedialog.askopenfilename(
            filetypes=[('Excel Dosyası', '*.xlsx;*.xls')],
            title='Excel Dosyasını Seçin'
        )
        if self.dosya_yolu:
            self.lbl_dosya.config(text=f'Seçilen Dosya: {self.dosya_yolu}')
            self.cmb_sayfa_sec.config(state='readonly')
            self.sayfa_listesini_yukle()

    def sayfa_listesini_yukle(self):
        """Excel dosyasındaki sayfaları combobox'a yükler."""
        xls = pd.ExcelFile(self.dosya_yolu)
        sayfalar = xls.sheet_names
        self.cmb_sayfa_sec['values'] = sayfalar
        self.cmb_sayfa_sec.current(0)

    def sayfa_sec(self, event):
        """Seçilen sayfayı yükler."""
        sayfa_adi = self.cmb_sayfa_sec.get()
        self.df = pd.read_excel(self.dosya_yolu, sheet_name=sayfa_adi)
        self.lbl_sayfa.config(text=f'Seçilen Sayfa: {sayfa_adi}')
        self.ent_orneklem_yuzdesi.config(state=tk.NORMAL)
        self.btn_orneklem_sec.config(state=tk.NORMAL)

    def orneklem_sec(self):
        """Örneklem yüzdesini alır, örneklem oluşturur ve sonuçları gösterir."""
        try:
            orneklem_yuzdesi = float(self.ent_orneklem_yuzdesi.get())
            if not (0 <= orneklem_yuzdesi <= 100):
                raise ValueError
        except ValueError:
            messagebox.showerror('Hata', 'Lütfen 0 ile 100 arasında bir değer girin.')
            return

        # Örneklem oluşturulur
        self.orneklem_df = self.df.sample(frac=orneklem_yuzdesi / 100, random_state=42)

        # Örneklem verileri 'Orneklem' adlı yeni bir sayfada kaydedilir
        with pd.ExcelWriter(self.dosya_yolu, engine='openpyxl', mode='a') as writer:
            self.orneklem_df.to_excel(writer, sheet_name='Orneklem', index=False)

        # Benzersizlik analizi
        benzersizlik_sonuclari = []
        for kolon in self.df.columns:
            orjinal_benzersiz_sayisi = self.df[kolon].nunique()
            orneklem_benzersiz_sayisi = self.orneklem_df[kolon].nunique()
            benzersizlik_orani = orneklem_benzersiz_sayisi / orjinal_benzersiz_sayisi if orjinal_benzersiz_sayisi > 0 else 0
            benzersizlik_sonuclari.append(f"'{kolon}': Örneklem: {orneklem_benzersiz_sayisi} / Orijinal: {orjinal_benzersiz_sayisi} / Oran: {benzersizlik_orani:.2f}")
            self.result_listbox.delete(0, tk.END)

        for sonuc in benzersizlik_sonuclari:
            self.result_listbox.insert(tk.END, sonuc)

        # Kullanıcıya bilgilendirme mesajı göster
        messagebox.showinfo('Bilgi', 'Örneklem başarıyla oluşturuldu ve benzersizlik analizi tamamlandı.')

if __name__ == '__main__':
    root = tk.Tk()
    app = OrneklemOlusturucu(root)
    root.mainloop()