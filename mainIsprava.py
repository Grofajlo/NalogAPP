import sqlite3 as sq
import ttkbootstrap as ttkb
from tkinter import END, messagebox
from datetime import datetime
import docx as doc
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn 
from docx.oxml import OxmlElement 
from docx.shared import RGBColor
from docxtpl import DocxTemplate

# GLAVNI PROGRAM ZA GENERISANJE ISPRAVA

dokument = Document()

root =ttkb.Window(themename="morph")
root.title("ISPRAVA O KONTROLISANJU")
root.grid_rowconfigure(4, weight=1)
root.grid_columnconfigure(0, weight=1)
root.iconbitmap(r"IMAGES\LOGO_TP.ico")

godina_label = ttkb.Label(root, text="Evidencija za godinu", width=30)
godina_label.grid(row=0, column=0, sticky="w", padx=10)

br_isprave_label = ttkb.Label(root, text="broj isprave", width=30)
br_isprave_label.grid(row=0, column=1, sticky="w", padx=10)

godina_entry = ttkb.Entry(root, width=15)
godina_entry.insert(0, "2025")
godina_entry.grid(row=1, column=0, padx=10, pady=10, sticky="w")

oblasti = "PPA H DP GP Ex DetP"
godina = godina_entry.get()

br_isprave_entry = ttkb.Entry(root, width=15)
br_isprave_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")
br_isprave_entry.focus()

datum_label = ttkb.Label(root, text="Datum isprave", width=30)
datum_label.grid(row=0, column=2, sticky="w", padx=10)

datum_entry = ttkb.Entry(root, width=15)
datum_entry.grid(row=1, column=2, padx=10, pady=10, sticky="w")

oblast_label = ttkb.Label(root, text="Datum isprave", width=30)
oblast_label.grid(row=0, column=3, sticky="w", padx=10)

oblast_combo = ttkb.Combobox(root, width=15, values = oblasti)
oblast_combo.grid(row=1, column=3, padx=10, pady=10, sticky="w")

frame_klijent = ttkb.LabelFrame(root, text="PODACI O KLIJENTU")
frame_klijent.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=20)

klijent_label = ttkb.Label(frame_klijent, text="Naziv klijenta:", width=20)
klijent_label.grid(row=0, column=0, sticky="w")

klijent_entry = ttkb.Entry(frame_klijent, width=60)
klijent_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

mesto_label = ttkb.Label(frame_klijent, text="Mesto:", width=20)
mesto_label.grid(row=1, column=0, sticky="w")

mesto_entry = ttkb.Entry(frame_klijent, width=60)
mesto_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

adresa_label = ttkb.Label(frame_klijent, text="Adresa:", width=20)
adresa_label.grid(row=2, column=0, sticky="w")

adresa_entry = ttkb.Entry(frame_klijent, width=60)
adresa_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

br_ugovora_label = ttkb.Label(frame_klijent, text="Broj ugovora/ponude:", width=20)
br_ugovora_label.grid(row=3, column=0, sticky="w")

br_ugovora_entry = ttkb.Entry(frame_klijent, width=60)
br_ugovora_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

def uvezi_podatke_o_klijentu():
    global klijent, mesto, adresa, invoice, rok, br_isprave
    con = sq.connect("Evidencija.db")
    c = con.cursor()
    br_isprave = br_isprave_entry.get()
    godina = godina_entry.get()
    
    # Execute the query
    c.execute(f"SELECT klijent, mesto, adresa, invoice, rok FROM Evidencija{godina} WHERE broj_dokumenta = ?", (br_isprave,))
    result = c.fetchone()
    
    if result:
        klijent, mesto, adresa, invoice, rok = result
        
        klijent_entry.delete(0, END)
        mesto_entry.delete(0, END)
        adresa_entry.delete(0, END)
        br_ugovora_entry.delete(0, END)
        datum_entry.delete(0, END)
        
        klijent_entry.insert(0, klijent)
        mesto_entry.insert(0, mesto)
        adresa_entry.insert(0, adresa)
        br_ugovora_entry.insert(0, invoice)
        datum_entry.insert(0, rok)
    else:
        messagebox.showinfo("OBAVEŠTENJE", f"Ne postoji nalog sa brojem {br_isprave}!!!")
    con.close()

def uvezi_prethodnu_ispravu():
    pass

def isprava_maker():
    global klijent, mesto, adresa, invoice, rok, br_isprave
    forma = DocxTemplate('isprava_test.docx')
    klijent = klijent_entry.get()
    mesto = mesto_entry.get()
    adresa = adresa_entry.get()
    invoice = br_ugovora_entry.get()
    rok = datum_entry.get()
    br_isprave = br_isprave_entry.get()
    
    context = {
        # kada bude proširena baza osoblja sa licencama dodati i ta polja da se odmah generišu
        "klijent": klijent,
        "mesto": mesto,
        "adresa": adresa,
        "invoice": invoice,
        "rok": rok,
        "br_dokumenta": br_isprave
    }
    forma.render(context)    

    forma.save(r"Nove_isprave/"f"{klijent}_{rok}.docx")
    messagebox.showinfo("OBAVEŠTENJE", F"Isprava broj {br_isprave}, za {klijent}, je uspešno sačuvana!")
    

kreiraj_ispravu_btn = ttkb.Button(root, text="KREIRAJ\nISPRAVU", width=10, command=isprava_maker)
kreiraj_ispravu_btn.grid(row=2, column=0, padx=10, pady=10, sticky="w")
klijent_data = ttkb.Button(root, text="PODACI\nO KLIJENTU", width=10, command=uvezi_podatke_o_klijentu)
klijent_data.grid(row=2, column=1, padx=10, pady=10, sticky="w")
uvezi_ispravu_btn = ttkb.Button(root, text="UVEZI PRETHODNU\nISPRAVU", width=15)
uvezi_ispravu_btn.grid(row=2, column=2, padx=10, pady=10, sticky="w")



root.mainloop()