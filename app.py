import os
import shutil
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import comtypes.client

def get_file_size_in_kb(file_path):
    
    return os.path.getsize(file_path) / 1000

def convert_word_to_pdf(input_file_path):
   
    word = comtypes.client.CreateObject("Word.Application")
    doc = word.Documents.Open(input_file_path)
    pdf_output_file = os.path.splitext(input_file_path)[0] + ".pdf"
    doc.SaveAs(pdf_output_file, FileFormat=17)
    doc.Close()
    word.Quit()
    return pdf_output_file

def compress_pdf(input_file_path, power=0):
    
    quality = {
        0: "/default",
        1: "/screen",
        2: "/ebook",
        3: "/printer",
        4: "/prepress",
    }

    # Temel kontrol işlemleri
   
    if not os.path.isfile(input_file_path):
        messagebox.showerror("Hata", f"Girdiğiniz PDF dosyası için geçerli bir yol değil: {input_file_path}")
        return

    
    if input_file_path.split('.')[-1].lower() != 'pdf':
        messagebox.showerror("Hata", f"Girdiğiniz dosya bir PDF değil: {input_file_path}")
        return

    initial_size_kb = get_file_size_in_kb(input_file_path)

    input_file_name = os.path.basename(input_file_path)
    output_file = os.path.join(os.path.expanduser("~"), "Desktop", f"{os.path.splitext(input_file_name)[0]}_Compressed.pdf")

    gs = get_ghostscript_path()
    subprocess.call(
        [
            gs,
            "-dSAFER",  # Güvenli modu etkinleştir
            "-dQUIET",
            "-dBATCH",
            "-dNOPAUSE",
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS={}".format(quality[power]),
            "-sOutputFile={}".format(output_file),
            input_file_path,
        ]
    )

    final_size_kb = get_file_size_in_kb(output_file)
    ratio = 1 - (final_size_kb / initial_size_kb)
    messagebox.showinfo("Sıkıştırma Tamamlandı",
                        f"Giriş dosyası boyutu: {initial_size_kb:.2f} KB\n"
                        f"Sıkıştırıldıktan sonraki boyutu: {final_size_kb:.2f} KB\n"
                        f"Sıkıştırma oranı: {ratio:.0%}")

def get_ghostscript_path():
    gs_names = ["gs", "gswin32", "gswin64"]
    for name in gs_names:
        if shutil.which(name):
            return shutil.which(name)
    raise FileNotFoundError(f"GhostScript yürütülebilir dosyası bulunamadı ({'/'.join(gs_names)})")

def browse_pdf_file():
    file_path = filedialog.askopenfilename(
        title="PDF Dosyası Seçin", filetypes=[("PDF Dosyaları", "*.pdf")]
    )
    input_file_entry.delete(0, tk.END)
    input_file_entry.insert(tk.END, file_path)

def browse_word_file():
    file_path = filedialog.askopenfilename(
        title="Word Dosyası Seçin", filetypes=[("Word Dosyaları", "*.docx")]
    )
    input_word_entry.delete(0, tk.END)
    input_word_entry.insert(tk.END, file_path)

def compress_button_callback():
    input_file = input_file_entry.get()
    compress_level = compress_level_var.get()
    
    
    if input_file.lower().endswith('.pdf'):
        compress_pdf(input_file, power=compress_level)
    else:
        messagebox.showerror("Hata", "Lütfen geçerli bir PDF dosyası seçin.")

def convert_button_callback():
    input_word_file = input_word_entry.get()
    
    if input_word_file.lower().endswith('.docx'):
        pdf_file = convert_word_to_pdf(input_word_file)
        messagebox.showinfo("Dönüştürme Tamamlandı",
                            f"Word belgesi PDF'e dönüştürüldü:\n{pdf_file}")
    else:
        messagebox.showerror("Hata", "Lütfen geçerli bir Word belgesi seçin.")
   
if __name__ == "__main__":
    root = tk.Tk()
    root.title("PDF Sıkıştırıcı ve Word'den PDF'e Dönüştürücü - Powered By Furkan Atasert")
    root.geometry("600x400")  # Pencere boyutu

    compress_level_var = tk.IntVar()
    compress_level_var.set(2)

    root.configure(bg="lightgray")

    title_label = tk.Label(root, text="PDF Sıkıştırıcı ve Word'den PDF'e Dönüştürücü \n Furkan ATASERT",
                           font=("Helvetica", 10, "bold"))
    title_label.grid(row=0, column=0, columnspan=5, padx=10, pady=10)

    input_file_label = tk.Label(root, text="PDF Dosyası Seçin:")
    input_file_label.grid(row=1, column=0, padx=10, pady=5, sticky="e")

    input_file_entry = tk.Entry(root, width=50)
    input_file_entry.grid(row=1, column=1, padx=10, pady=5, columnspan=3)

    browse_input_button = tk.Button(root, text="Gözat", command=browse_pdf_file, bg="#007ACC", fg="white")
    browse_input_button.grid(row=1, column=4, padx=10, pady=10, sticky="w")

    compress_level_label = tk.Label(root, text="Sıkıştırma Seviyesi:")
    compress_level_label.grid(row=2, column=0, padx=10, pady=5, sticky="e")

    compress_level_menu = tk.OptionMenu(root, compress_level_var, 1, 2, 3, 4)
    compress_level_menu.grid(row=2, column=1, padx=10, pady=5, columnspan=3)

    compress_button = tk.Button(root, text="PDF'i Sıkıştır", command=compress_button_callback, bg="#4CAF50", fg="white")
    compress_button.grid(row=3, column=0, columnspan=5, padx=15, pady=15)

    input_word_label = tk.Label(root, text="Word Dosyası Seçin:")
    input_word_label.grid(row=4, column=0, padx=10, pady=5, sticky="e")

    input_word_entry = tk.Entry(root, width=50)
    input_word_entry.grid(row=4, column=1, padx=10, pady=5, columnspan=3)

    browse_word_button = tk.Button(root, text="Gözat", command=browse_word_file, bg="#007ACC", fg="white")
    browse_word_button.grid(row=4, column=4, padx=5, pady=5, sticky="w")

    convert_button = tk.Button(root, text="Word'den PDF'e Dönüştür", command=convert_button_callback, bg="#FF9800", fg="white")
    convert_button.grid(row=5, column=0, columnspan=5, padx=15, pady=15)

    quality_info_label = tk.Label(root, text="Sıkıştırma Seviyesi Açıklamaları:\n\n" +
                                            "1: Ekranlar için uygun, düşük kaliteli ve yüksek sıkıştırma. (En Yüksek Sıkıştırma Derecesi!!) \n" +
                                            "2: Dijital kitaplar için uygun, daha yüksek kalite ve orta düzeyde sıkıştırma.\n" +
                                            "3: Yazdırma için uygun, daha yüksek kalite ve orta düzeyde sıkıştırma.\n" +
                                            "4: Baskı için uygun, yüksek kalite ve düşük düzeyde sıkıştırma.\n" ,
                                            bg="#f0f0f0")  # Arka plan rengi

    quality_info_label.grid(row=6, column=0, columnspan=5, padx=10, pady=5)
    
    for i in range(7):  
        root.grid_rowconfigure(i, weight=1) 
    for i in range(5): 
        root.grid_columnconfigure(i, weight=1)  

    root.mainloop()
