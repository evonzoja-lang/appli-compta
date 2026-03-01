import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import datetime
import pandas as pd
import csv
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from tkcalendar import DateEntry

DB_FILE = "comptabilite.db"

# -----------------------
# INITIALISATION DB
# -----------------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        type TEXT NOT NULL,
        montant REAL NOT NULL,
        description TEXT,
        categorie TEXT,
        compte TEXT,
        mode_paiement TEXT,
        date TEXT
    )
    """)
    conn.commit()
    conn.close()

init_db()

CATEGORIES = ["Ventes", "Achats", "Formation","Salaires", "Cotisations", "Reparation", "Materiels", "Location","Autres"]
COMPTES = ["Caisse", "Banque","Mobile Money","Autres"]
MODES_PAIEMENT = ["Cash", "Chèque", "Lumicash", "Ecocash", "Bancobu Enoti", "Ihera", "Cashtel", "Gasape Cash", "Akaravyo", "autres"]
TYPES_TRANSACTION = ["Entrée", "Sortie"]
MODIFIER_ID = None

# -----------------------
# BASE DE DONNÉES
# -----------------------
def lire_transactions():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, type, montant, description, categorie, compte, mode_paiement, date
        FROM transactions
        ORDER BY date DESC
    """)
    rows = cursor.fetchall()
    conn.close()
    return rows

# -----------------------
# TRI COLONNE
# -----------------------
def trier_colonne(col, reverse):
    data = [(tree.set(k, col), k) for k in tree.get_children("")]
    try:
        data.sort(key=lambda t: float(t[0]), reverse=reverse)
    except:
        data.sort(reverse=reverse)
    for index, (val, k) in enumerate(data):
        tree.move(k, "", index)
    tree.heading(col, command=lambda: trier_colonne(col, not reverse))

# -----------------------
# TABLEAU + SOLDES CUMULATIFS
# -----------------------
def mise_a_jour_tableau(transactions=None):
    if transactions is None:
        transactions = lire_transactions()
    for row in tree.get_children():
        tree.delete(row)
    total_entrees = total_sorties = solde_cumul = 0
    for t in transactions:
        id_, type_, montant, description, categorie, compte, mode_paiement, date = t
        montant_val = float(montant)
        debit = credit = 0
        if type_ == "Entrée":
            credit = montant_val
            total_entrees += montant_val
            solde_cumul += montant_val
        else:
            debit = montant_val
            total_sorties += montant_val
            solde_cumul -= montant_val
        tree.insert("", "end", values=(id_, date, description, categorie,
                                       type_, credit, debit, solde_cumul, compte, mode_paiement))
    total_label.config(
        text=f"Total Entrées: {total_entrees:.2f} F   Total Sorties: {total_sorties:.2f} F   Solde: {solde_cumul:.2f} F"
    )
    mise_a_jour_resume(transactions)

# -----------------------
# AJOUT / MODIFICATION / SUPPRESSION
# -----------------------
def effacer_champs():
    global MODIFIER_ID
    montant_entry.delete(0, tk.END)
    description_entry.delete(0, tk.END)
    date_entry.set_date(datetime.date.today())
    MODIFIER_ID = None

def ajouter_transaction():
    global MODIFIER_ID
    t_type = type_var.get()
    montant_str = montant_entry.get().strip()
    if not montant_str:
        messagebox.showerror("Erreur", "Montant obligatoire")
        return
    try:
        montant = float(montant_str)
    except:
        messagebox.showerror("Erreur", "Montant invalide")
        return
    description = description_entry.get()
    categorie = categorie_var.get()
    compte = compte_var.get()
    mode_paiement = mode_var.get()
    date_input = date_entry.get_date().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    if MODIFIER_ID:
        cursor.execute("""
            UPDATE transactions
            SET type=?, montant=?, description=?, categorie=?, compte=?, mode_paiement=?, date=?
            WHERE id=?
        """, (t_type, montant, description, categorie, compte, mode_paiement, date_input, MODIFIER_ID))
        MODIFIER_ID = None
    else:
        cursor.execute("""
            INSERT INTO transactions (type, montant, description, categorie, compte, mode_paiement, date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (t_type, montant, description, categorie, compte, mode_paiement, date_input))
    conn.commit()
    conn.close()
    effacer_champs()
    mise_a_jour_tableau()

def supprimer_transaction():
    selected = tree.selection()
    if not selected:
        return
    if not messagebox.askyesno("Confirmation", "Supprimer cette transaction ?"):
        return
    transaction_id = tree.item(selected[0])["values"][0]
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM transactions WHERE id=?", (transaction_id,))
    conn.commit()
    conn.close()
    mise_a_jour_tableau()
    effacer_champs()

def modifier_transaction():
    global MODIFIER_ID
    selected = tree.selection()
    if not selected:
        return
    transaction_id = tree.item(selected[0])["values"][0]
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM transactions WHERE id=?", (transaction_id,))
    t = cursor.fetchone()
    conn.close()
    if t:
        MODIFIER_ID = t[0]
        type_var.set(t[1])
        montant_entry.delete(0, tk.END)
        montant_entry.insert(0, t[2])
        description_entry.delete(0, tk.END)
        description_entry.insert(0, t[3])
        categorie_var.set(t[4])
        compte_var.set(t[5])
        mode_var.set(t[6])
        date_entry.set_date(datetime.datetime.strptime(t[7], "%Y-%m-%d").date())

# -----------------------
# RECHERCHE / FILTRE SQL
# -----------------------
def rechercher_transactions():
    conn = sqlite3.connect("comptabilite.db")
    cursor = conn.cursor()

    requete = "SELECT * FROM transactions WHERE 1=1"
    params = []

    # --- NORMALISATION ---
    categorie = filtre_categorie.get().strip()
    type_tx = filtre_type.get().strip()
    compte = filtre_compte.get().strip()
    mode = filtre_mode.get().strip()

    date_exacte = filtre_date.get().strip()
    date_debut = filtre_date_debut.get().strip()
    date_fin = filtre_date_fin.get().strip()

    # -----------------------
    # FILTRES TEXTE (indépendants)
    # -----------------------

    if categorie:
        requete += " AND LOWER(TRIM(categorie)) = LOWER(TRIM(?))"
        params.append(categorie)

    if type_tx:
        requete += " AND LOWER(TRIM(type)) = LOWER(TRIM(?))"
        params.append(type_tx)

    if compte:
        requete += " AND LOWER(TRIM(compte)) = LOWER(TRIM(?))"
        params.append(compte)

    if mode:
        requete += " AND LOWER(TRIM(mode_paiement)) = LOWER(TRIM(?))"
        params.append(mode)

    # -----------------------
    # FILTRES DATE (indépendants)
    # -----------------------

    try:
        # Date exacte seule
        if date_exacte:
            date_sql = datetime.datetime.strptime(date_exacte, "%d-%m-%Y").strftime("%Y-%m-%d")
            requete += " AND date = ?"
            params.append(date_sql)

        # Date début seule
        if date_debut:
            date_debut_sql = datetime.datetime.strptime(date_debut, "%d-%m-%Y").strftime("%Y-%m-%d")
            requete += " AND date >= ?"
            params.append(date_debut_sql)

        # Date fin seule
        if date_fin:
            date_fin_sql = datetime.datetime.strptime(date_fin, "%d-%m-%Y").strftime("%Y-%m-%d")
            requete += " AND date <= ?"
            params.append(date_fin_sql)

    except ValueError:
        messagebox.showerror("Erreur", "Format de date invalide.")
        conn.close()
        return

    # Tri chronologique propre
    requete += " ORDER BY date DESC"

    print("REQUETE:", requete)
    print("PARAMS:", params)

    cursor.execute(requete, params)
    transactions = cursor.fetchall()

    conn.close()

    mise_a_jour_tableau(transactions)

def tout_afficher():
    conn = sqlite3.connect("comptabilite.db")
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM transactions ORDER BY date DESC")
    transactions = cursor.fetchall()

    conn.close()

    mise_a_jour_tableau(transactions)

# -----------------------
# RAPPORTS
# -----------------------
def filtrer_transactions(par):
    today = datetime.date.today()
    transactions = lire_transactions()
    filtered = []
    for t in transactions:
        try:
            t_date = datetime.datetime.strptime(t[7], "%Y-%m-%d").date()
        except:
            continue
        if par=="jour" and t_date == today:
            filtered.append(t)
        elif par=="semaine" and (today - t_date).days < 7:
            filtered.append(t)
        elif par=="mois" and t_date.month == today.month and t_date.year == today.year:
            filtered.append(t)
        elif par=="annee" and t_date.year == today.year:
            filtered.append(t)
    return filtered

def afficher_rapport(transactions, titre):
    if not transactions:
        messagebox.showinfo(titre, "Aucune transaction")
        return

    fen = tk.Toplevel(root)
    fen.title(f"Rapport - {titre}")
    fen.geometry("950x550")

    cols = ("Date","Description","Catégorie","Type","Montant","Compte","Mode")
    tree_rap = ttk.Treeview(fen, columns=cols, show="headings")

    scroll = ttk.Scrollbar(fen, orient="vertical", command=tree_rap.yview)
    tree_rap.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")

    for c in cols:
        tree_rap.heading(c, text=c)
        tree_rap.column(c, width=120, anchor="center")

    tree_rap.pack(fill="both", expand=True, padx=10, pady=10)

    # Style couleurs
    tree_rap.tag_configure("entree", foreground="green")
    tree_rap.tag_configure("sortie", foreground="red")
    tree_rap.tag_configure("total", background="#E0E0E0", font=("Arial", 10, "bold"))

    total_ent = 0
    total_sort = 0

    for t in transactions:
        date = t[7]
        description = t[3]
        categorie = t[4]
        type_tx = t[1]
        montant = float(t[2])
        compte = t[5]
        mode = t[6]

        if type_tx.strip().lower() == "entrée":
            total_ent += montant
            tag = "entree"
        elif type_tx.strip().lower() == "sortie":
            total_sort += montant
            tag = "sortie"
        else:
            tag = ""

        tree_rap.insert("", "end",
                        values=(date, description, categorie, type_tx,
                                f"{montant:,.2f}", compte, mode),
                        tags=(tag,))

    solde = total_ent - total_sort

    # Ligne vide
    tree_rap.insert("", "end", values=("", "", "", "", "", "", ""))

    # Ligne totaux dans le tableau
    tree_rap.insert("", "end",
                    values=("","TOTAL ENTRÉES","", "", f"{total_ent:,.2f} F","", ""),
                    tags=("total",))

    tree_rap.insert("", "end",
                    values=("","TOTAL SORTIES","", "", f"{total_sort:,.2f} F","", ""),
                    tags=("total",))

    tree_rap.insert("", "end",
                    values=("","SOLDE","", "", f"{solde:,.2f} F","", ""),
                    tags=("total",))

    # Résumé en bas (optionnel mais plus visible)
    lbl_tot = tk.Label(
        fen,
        text=f"Total Entrées: {total_ent:,.2f} F   |   "
             f"Total Sorties: {total_sort:,.2f} F   |   "
             f"Solde: {solde:,.2f} F",
        font=("Arial",12,"bold")
    )
    lbl_tot.pack(pady=5)

    tk.Button(
        fen,
        text="Imprimer PDF",
        command=lambda: imprimer_pdf(transactions, f"Historique_{titre}"),
        font=("Arial",12)
    ).pack(pady=5)


# -----------------------
# EXPORTATIONS PDF EXCEL
# -----------------------


# ----------------------------
# PDF
# ----------------------------
    
def imprimer_pdf(transactions, titre):
    file_name = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialfile=f"{titre}.pdf"
    )
    if not file_name:
        return

    doc = SimpleDocTemplate(file_name, pagesize=A4)
    elements = []

    styles = getSampleStyleSheet()

    small_style = styles["Normal"]
    small_style.fontSize = 8
    small_style.leading = 10

    elements.append(Paragraph(titre, styles['Title']))
    elements.append(Paragraph("<br/><br/>", styles['Normal']))

    data_pdf = []

    header = ["Date","Description","Catégorie","Type","Montant","Compte","Mode"]
    data_pdf.append([Paragraph(h, styles["Heading6"]) for h in header])

    total_ent = 0
    total_sort = 0

    for t in transactions:
        type_tx = str(t[1]).strip()
        montant = float(t[2])

        if type_tx.lower() == "entrée":
            total_ent += montant
        elif type_tx.lower() == "sortie":
            total_sort += montant

        row = [
            Paragraph(str(t[7]), small_style),
            Paragraph(str(t[3]), small_style),
            Paragraph(str(t[4]), small_style),
            Paragraph(type_tx, small_style),
            Paragraph(f"{montant:,.2f}", small_style),
            Paragraph(str(t[5]), small_style),
            Paragraph(str(t[6]), small_style),
        ]
        data_pdf.append(row)

    table = Table(data_pdf, colWidths=[55,180,65,70,70,55,55])

    style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2196F3")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
        ('ALIGN',(4,1), (4,-1),'RIGHT'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('LEFTPADDING',(0,0),(-1,-1),4),
        ('RIGHTPADDING',(0,0),(-1,-1),4),
        ('TOPPADDING',(0,0),(-1,-1),3),
        ('BOTTOMPADDING',(0,0),(-1,-1),3),
    ])

    table.setStyle(style)
    elements.append(table)

    # ----------------------------
    # RÉSUMÉ EN DEHORS DU TABLEAU
    # ----------------------------

    solde = total_ent - total_sort

    elements.append(Paragraph("<br/><br/>", styles['Normal']))

    resume_style = styles["Normal"]
    resume_style.fontSize = 11
    resume_style.leading = 14

    elements.append(Paragraph(f"<b>Total Entrées :</b> {total_ent:,.2f} F", resume_style))
    elements.append(Paragraph(f"<b>Total Sorties :</b> {total_sort:,.2f} F", resume_style))

    if solde >= 0:
        elements.append(Paragraph(f"<b>Solde :</b> <font color='green'>{solde:,.2f} F</font>", resume_style))
    else:
        elements.append(Paragraph(f"<b>Solde :</b> <font color='red'>{solde:,.2f} F</font>", resume_style))

    doc.build(elements)

    messagebox.showinfo("Succès", f"PDF généré : {file_name}")
    
# ----------------------------
# EXCEL
# ----------------------------
def exporter_csv(transactions):
    file_name = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="Transactions.csv")
    if not file_name:
        return
    with open(file_name,"w",newline="",encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Date","Description","Catégorie","Type","Montant","Compte","Mode"])
        for t in transactions:
            writer.writerow([t[7], t[3], t[4], t[1], t[2], t[5], t[6]])
    messagebox.showinfo("Succès", f"Export CSV généré : {file_name}")

# -----------------------
# INTERFACE PRINCIPALE
# -----------------------
root = tk.Tk()
root.title("Application Comptabilité INT MSC")
root.geometry("1600x850")

# ---- FONTS ----
font_label = ("Arial", 12)
font_entry = ("Arial", 12)
font_button = ("Arial", 12, "bold")

# ---- SAISIE ----
frame_saisie = tk.LabelFrame(root, text="Gestion Transactions", padx=20, pady=20, font=("Arial",14,"bold"))
frame_saisie.pack(fill="x", padx=20, pady=10)

type_var = tk.StringVar(value=TYPES_TRANSACTION[0])
ttk.Combobox(frame_saisie, textvariable=type_var, values=TYPES_TRANSACTION, font=font_entry, width=15).grid(row=0,column=1, padx=5, pady=5)
tk.Label(frame_saisie,text="Type", font=font_label).grid(row=0,column=0, sticky="w", padx=5, pady=5)

montant_entry = tk.Entry(frame_saisie, font=font_entry, width=17)
montant_entry.grid(row=0,column=3, padx=5, pady=5)
tk.Label(frame_saisie,text="Montant", font=font_label).grid(row=0,column=2, sticky="w", padx=5, pady=5)

description_entry = tk.Entry(frame_saisie, font=font_entry, width=17)
description_entry.grid(row=0,column=5, padx=5, pady=5)
tk.Label(frame_saisie,text="Description", font=font_label).grid(row=0,column=4, sticky="w", padx=5, pady=5)

categorie_var = tk.StringVar(value=CATEGORIES[0])
ttk.Combobox(frame_saisie,textvariable=categorie_var, values=CATEGORIES, font=font_entry, width=15).grid(row=1,column=1, padx=5, pady=5)
tk.Label(frame_saisie,text="Catégorie", font=font_label).grid(row=1,column=0, sticky="w", padx=5, pady=5)

compte_var = tk.StringVar(value=COMPTES[0])
ttk.Combobox(frame_saisie,textvariable=compte_var, values=COMPTES, font=font_entry, width=15).grid(row=1,column=3, padx=5, pady=5)
tk.Label(frame_saisie,text="Compte", font=font_label).grid(row=1,column=2, sticky="w", padx=5, pady=5)

mode_var = tk.StringVar(value=MODES_PAIEMENT[0])

ttk.Combobox(frame_saisie,textvariable=mode_var,values=MODES_PAIEMENT,font=font_entry, width=15,state="readonly").grid(row=1, column=5, padx=5, pady=5)
tk.Label(frame_saisie,text="Mode",font=font_label).grid(row=1, column=4, sticky="w", padx=5, pady=5)

date_entry = DateEntry(frame_saisie, font=font_entry, width=15, date_pattern='dd-mm-yyyy')
date_entry.grid(row=2,column=1, padx=5, pady=10)
tk.Label(frame_saisie,text="Date", font=font_label).grid(row=2,column=0, sticky="w", padx=5, pady=10)

tk.Button(frame_saisie,text="Ajouter/Modifier", command=ajouter_transaction,bg="#4CAF50",fg="white", font=font_button, width=18).grid(row=2,column=3, padx=5, pady=10)
tk.Button(frame_saisie,text="Supprimer", command=supprimer_transaction,bg="#f44336",fg="white", font=font_button, width=15).grid(row=2,column=4, padx=5, pady=10)
tk.Button(frame_saisie,text="Modifier", command=modifier_transaction,bg="#FF9800",fg="white", font=font_button, width=15).grid(row=2,column=5, padx=5, pady=10)

# ---- FILTRE ET RAPPORTS ----
frame_top = tk.Frame(root)
frame_top.pack(fill="x", padx=20, pady=10)

frame_filtre = tk.LabelFrame(frame_top, text="Recherche / Filtrage", padx=10, pady=10, font=("Arial",12,"bold"))
frame_filtre.pack(side="left", padx=10, pady=5)

# --- CHAMPS EXISTANTS ---
filtre_categorie = tk.StringVar(value="")
ttk.Combobox(frame_filtre, textvariable=filtre_categorie, values=CATEGORIES, font=font_entry, width=12).grid(row=0,column=1, padx=5, pady=5)
tk.Label(frame_filtre,text="Catégorie:", font=font_label).grid(row=0,column=0, padx=5, pady=5)

tk.Label(frame_filtre, text="Date exacte:", font=font_label).grid(row=1, column=0, padx=5, pady=5)
filtre_date = DateEntry(frame_filtre, font=font_entry, width=12, date_pattern='dd-mm-yyyy')
filtre_date.grid(row=1, column=1, padx=5)

tk.Label(frame_filtre, text="Date début:", font=font_label).grid(row=2, column=0, padx=5, pady=5)
filtre_date_debut = DateEntry(frame_filtre, font=font_entry, width=12, date_pattern='dd-mm-yyyy')
filtre_date_debut.grid(row=2, column=1, padx=5)

tk.Label(frame_filtre, text="Date fin:", font=font_label).grid(row=3, column=0, padx=5, pady=5)
filtre_date_fin = DateEntry(frame_filtre, font=font_entry, width=12, date_pattern='dd-mm-yyyy')
filtre_date_fin.grid(row=3, column=1, padx=5, pady=5)

# --- NOUVEAUX CHAMPS (Type, Compte, Mode) ---
filtre_type = tk.StringVar(value="")
tk.Label(frame_filtre, text="Type:", font=font_label).grid(row=0, column=2, padx=5, pady=5)
ttk.Combobox(frame_filtre, textvariable=filtre_type, values=TYPES_TRANSACTION, font=font_entry, width=12).grid(row=0, column=3, padx=5, pady=5)

filtre_compte = tk.StringVar(value="")
tk.Label(frame_filtre, text="Compte:", font=font_label).grid(row=1, column=2, padx=5, pady=5)
ttk.Combobox(frame_filtre, textvariable=filtre_compte, values=COMPTES, font=font_entry, width=12).grid(row=1, column=3, padx=5, pady=5)

filtre_mode = tk.StringVar(value="")
tk.Label(frame_filtre, text="Mode:", font=font_label).grid(row=2, column=2, padx=5, pady=5)
ttk.Combobox(frame_filtre, textvariable=filtre_mode, values=MODES_PAIEMENT, font=font_entry, width=12).grid(row=2, column=3, padx=5, pady=5)


# --- BOUTONS ---
tk.Button(frame_filtre,text="Rechercher",command=rechercher_transactions,font=font_button,bg="#2196F3",fg="white").grid(row=4,column=0,columnspan=2,pady=10)
tk.Button(frame_filtre,text="Tout afficher",command=tout_afficher,font=font_button,bg="#607D8B",fg="white").grid(row=5,column=0,columnspan=2,pady=5)


# --- CARD RÉSUMÉ ---
resume_frame = tk.LabelFrame(frame_top, text="Résumé", padx=10, pady=10, font=("Arial",12,"bold"))
resume_frame.pack(side="left", padx=10, pady=5)

resume_label = tk.Label(resume_frame, text="", font=("Arial",12,"bold"), justify="left")
resume_label.pack(padx=5, pady=5)

def mise_a_jour_resume(transactions=None):
    if transactions is None:
        transactions = lire_transactions()
    total_entrees = sum(float(t[2]) for t in transactions if t[1]=="Entrée")
    total_sorties = sum(float(t[2]) for t in transactions if t[1]=="Sortie")
    solde = total_entrees - total_sorties
    resume_label.config(
        text=f"Total Entrées: {total_entrees:.2f} F\n"
             f"Total Sorties: {total_sorties:.2f} F\n"
             f"Solde: {solde:.2f} F"
    )
    # Met à jour les informations supplémentaires
    mise_a_jour_infos(transactions)
    
    # --- INFORMATIONS SUPPLÉMENTAIRES ENTRE RÉSUMÉ ET RAPPORTS ---
infos_frame = tk.LabelFrame(frame_top, text="Informations", padx=10, pady=10, font=("Arial",12,"bold"))
infos_frame.pack(side="left", padx=10, pady=5)

infos_label = tk.Label(infos_frame, text="", font=("Arial",12), justify="left")
infos_label.pack(padx=5, pady=5)

def mise_a_jour_infos(transactions=None):
    if transactions is None:
        transactions = lire_transactions()
    
    nb_transactions = len(transactions)
    derniere_trans = transactions[-1] if transactions else None
    dernier_detail = f"{derniere_trans[0]} | {derniere_trans[1]} | {derniere_trans[2]} F" if derniere_trans else "Aucune"
    total_entrees = sum(float(t[2]) for t in transactions if t[1]=="Entrée")
    total_sorties = sum(float(t[2]) for t in transactions if t[1]=="Sortie")
    moyenne_entrees = (total_entrees / nb_transactions) if nb_transactions else 0
    moyenne_sorties = (total_sorties / nb_transactions) if nb_transactions else 0
    
    infos_label.config(
        text=f"Nombre total de transactions: {nb_transactions}\n"
             f"Dernière transaction: {dernier_detail}\n"
             
    )

# --- RAPPORTS À DROITE ---
frame_rapports = tk.LabelFrame(frame_top, text="Rapports", padx=10, pady=10, font=("Arial",12,"bold"))
frame_rapports.pack(side="right", padx=10)
tk.Button(frame_rapports,text="Journalier",command=lambda: afficher_rapport(filtrer_transactions("jour"), "Journalier"),font=font_button,bg="#2196F3",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Hebdomadaire",command=lambda: afficher_rapport(filtrer_transactions("semaine"), "Hebdomadaire"),font=font_button,bg="#2196F3",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Mensuel",command=lambda: afficher_rapport(filtrer_transactions("mois"), "Mensuel"),font=font_button,bg="#2196F3",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Annuel",command=lambda: afficher_rapport(filtrer_transactions("annee"), "Annuel"),font=font_button,bg="#2196F3",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Complet",command=lambda: afficher_rapport(lire_transactions(), "Complet"),font=font_button,bg="#4CAF50",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Exporter CSV",command=lambda: exporter_csv(lire_transactions()),font=font_button,bg="#FF9800",fg="white").pack(padx=5,pady=5, fill="x")

# ---- TABLEAU ----
frame_table = tk.Frame(root)
frame_table.pack(fill="both", expand=True, padx=20, pady=10)
cols = ("ID","Date","Description","Catégorie","Type","Débit","Crédit","Solde","Compte","Mode")
tree = ttk.Treeview(frame_table, columns=cols, show="headings")
scroll = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scroll.set)
scroll.pack(side="right", fill="y")
for c in cols:
    tree.heading(c, text=c, command=lambda _c=c: trier_colonne(_c, False))
    tree.column(c, width=120)
tree.pack(fill="both", expand=True)
total_label = tk.Label(root, text="", font=("Arial",12,"bold"))
total_label.pack(pady=5)

# --- INITIALISATION ---
mise_a_jour_tableau()
root.mainloop()
