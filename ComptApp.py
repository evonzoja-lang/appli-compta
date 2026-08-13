import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import datetime
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph,Spacer, HRFlowable
from reportlab.lib.pagesizes import A4
from tkcalendar import DateEntry
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Alignment
from reportlab.lib import pagesizes
from matplotlib.gridspec import GridSpec
from reportlab.lib.enums import TA_LEFT
import os
import sys
import statistics

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_path()
DB_FILE = os.path.join(BASE_DIR, "comptabilite.db")
print("Base utilisée :", DB_FILE)

DOCUMENTS = os.path.join(os.path.expanduser("~"), "Documents")
APP_FOLDER = os.path.join(DOCUMENTS, "Comptabilite_Magasin_Sainte_Rita")
RAPPORTS_FOLDER = os.path.join(APP_FOLDER, "rapports")
os.makedirs(RAPPORTS_FOLDER, exist_ok=True)

def get_app_folder():
    documents = os.path.join(os.path.expanduser("~"), "Documents")
    app_folder = os.path.join(documents, "Comptabilite_Magasin_Sainte_Rita")
    if not os.path.exists(app_folder):
        os.makedirs(app_folder)
    return app_folder

APP_FOLDER = get_app_folder()

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

def trier_colonne(col, reverse):
    if col == "Solde":
        return
    data = [(tree.set(k, col), k) for k in tree.get_children("")]
    try:
        data.sort(key=lambda t: float(str(t[0]).replace(',', '').replace(' ', '')), reverse=reverse)
    except:
        data.sort(reverse=reverse)
    for index, (val, k) in enumerate(data):
        if tree.exists(k):
            tree.move(k, "", index)
    tree.heading(col, command=lambda: trier_colonne(col, not reverse))

# -----------------------
# TABLEAU + SOLDES CUMULATIFS CORRIGE SOLDE = DEBIT - CREDIT
# -----------------------
def mise_a_jour_tableau(transactions=None):
    if transactions is None:
        transactions = lire_transactions()
    for row in tree.get_children():
        tree.delete(row)

    # Pour le solde cumulé on doit calculer en ordre ASC
    def parse_date(t):
        try:
            return datetime.datetime.strptime(t[7], "%d-%m-%Y")
        except:
            return datetime.datetime.min

    transactions_asc = sorted(transactions, key=parse_date)
    solde_par_id = {}
    cumul = 0
    for t in transactions_asc:
        _, type_, montant, _, _, _, _, _ = t
        try: mv = float(montant)
        except: mv = 0
        if str(type_).strip().lower() == "entrée":
            cumul += mv # Débit +
        else:
            cumul -= mv # Crédit -
        solde_par_id[t[0]] = cumul

    # Affichage en DESC
    transactions_desc = sorted(transactions, key=parse_date, reverse=True)
    total_debit = 0
    total_credit = 0
    for t in transactions_desc:
        id_, type_, montant, description, categorie, compte, mode_paiement, date = t
        montant_val = float(montant)
        debit = 0
        credit = 0
        if type_ == "Entrée":
            debit = montant_val # CORRIGE : Entrée = Débit
            total_debit += montant_val
        else:
            credit = montant_val # CORRIGE : Sortie = Crédit
            total_credit += montant_val

        # CORRIGE : on insère bien debit puis credit puis solde = debit - credit cumulé
        tree.insert("", "end", values=(id_, date, description, categorie,
                                       type_, debit, credit, solde_par_id[id_], compte, mode_paiement))

    solde_final = total_debit - total_credit # CORRIGE : Solde = Débit - Crédit
    total_label.config(
        text=f"Tot Débit: {total_debit:.2f} F Tot Crédit: {total_credit:.2f} F Solde (Débit-Crédit): {solde_final:.2f} F"
    )
    mise_a_jour_resume(transactions)

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
    date_input = date_entry.get_date().strftime("%d-%m-%Y")
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
            VALUES (?,?,?,?,?,?,?)
        """, (t_type, montant, description, categorie, compte, mode_paiement, date_input))
    conn.commit()
    conn.close()
    effacer_champs()
    mise_a_jour_tableau()

def supprimer_transaction():
    selected = tree.selection()
    if not selected:
        return
    if not messagebox.askyesno("Confirmation", "Supprimer cette transaction?"):
        return
    transaction_id = tree.item(selected[0])["values"][0]
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM transactions WHERE id=?", (transaction_id,))
    conn.commit()
    conn.close()
    mise_a_jour_tableau()
    effacer_champs()

ADMIN_PASSWORD = "1234"

def demander_mot_de_passe():
    fenetre_mdp = tk.Toplevel()
    fenetre_mdp.title("Authentification requise")
    fenetre_mdp.geometry("300x150")
    fenetre_mdp.resizable(False, False)
    tk.Label(fenetre_mdp, text="Entrez le mot de passe :", font=("Arial", 11)).pack(pady=10)
    entry_mdp = tk.Entry(fenetre_mdp, show="*", width=25)
    entry_mdp.pack(pady=5)
    def verifier():
        if entry_mdp.get() == ADMIN_PASSWORD:
            fenetre_mdp.destroy()
            supprimer_transaction()
        else:
            messagebox.showerror("Erreur", "Mot de passe incorrect")
    tk.Button(fenetre_mdp, text="Valider", command=verifier, bg="#4CAF50", fg="white", width=12).pack(pady=10)

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
        date_entry.set_date(datetime.datetime.strptime(t[7], "%d-%m-%Y").date())

# -----------------------
# RECHERCHE / FILTRE SQL
# -----------------------
def rechercher_transactions():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    requete = "SELECT * FROM transactions"
    conditions = []
    params = []
    categorie = filtre_categorie.get().strip()
    type_tx = filtre_type.get().strip()
    compte = filtre_compte.get().strip()
    mode = filtre_mode.get().strip()
    try:
        date_debut = filtre_date_debut.get_date()
        date_fin = filtre_date_fin.get_date()
    except:
        date_debut = None
        date_fin = None

    if categorie:
        conditions.append("LOWER(TRIM(categorie)) = LOWER(TRIM(?))")
        params.append(categorie)
    if type_tx:
        conditions.append("LOWER(TRIM(type)) = LOWER(TRIM(?))")
        params.append(type_tx)
    if compte:
        conditions.append("LOWER(TRIM(compte)) = LOWER(TRIM(?))")
        params.append(compte)
    if mode:
        conditions.append("LOWER(TRIM(mode_paiement)) = LOWER(TRIM(?))")
        params.append(mode)

    if conditions:
        requete += " WHERE " + " AND ".join(conditions)
    requete += " ORDER BY date DESC"
    cursor.execute(requete, params)
    transactions = cursor.fetchall()
    conn.close()

    # Filtrage date en Python car format DD-MM-YYYY
    if date_debut and date_fin:
        transactions = [t for t in transactions if parse_date_filtre(t[7]) and date_debut <= parse_date_filtre(t[7]).date() <= date_fin]

    mise_a_jour_tableau(transactions)

def parse_date_filtre(s):
    try:
        return datetime.datetime.strptime(s, "%d-%m-%Y")
    except:
        return None

def tout_afficher():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM transactions ORDER BY date DESC")
    transactions = cursor.fetchall()
    conn.close()
    mise_a_jour_tableau(transactions)

def imprimer_selection():
    items = tree.get_children()
    if not items:
        messagebox.showwarning("Attention", "Aucune donnée à imprimer")
        return
    dossier = os.path.join(APP_FOLDER, "rapports")
    os.makedirs(dossier, exist_ok=True)
    now = datetime.datetime.now().strftime("%d%m%Y_%H%M%S")
    file_path = os.path.join(dossier, f"rapport_filtre_{now}.pdf")
    doc = SimpleDocTemplate(file_path, pagesize=pagesizes.A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    elements.append(Paragraph("<b>MAGASIN SAINTE RITA - RAPPORT FILTRE</b>", styles["Title"]))
    elements.append(Spacer(1, 10))
    date_rapport = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    elements.append(Paragraph(f"Date du rapport : {date_rapport}", styles["Normal"]))
    elements.append(Spacer(1, 20))
    data = []
    headers = tree["columns"][:-2]
    data.append(headers)
    total_debit = 0
    total_credit = 0
    for item in items:
        values = tree.item(item)["values"][:-2]
        data.append(values)
        try:
            debit = float(values[5]) # CORRIGE colonne Débit
            credit = float(values[6]) # CORRIGE colonne Crédit
            total_debit += debit
            total_credit += credit
        except:
            pass
    table = Table(data, repeatRows=1, hAlign='CENTER')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
        ('ALIGN',(0,0),(-1,-1),'LEFT'),
        ('FONTSIZE', (0,0), (-1,-1), 7),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('BOTTOMPADDING',(0,0),(-1,-1),2),
        ('TOPPADDING',(0,0),(-1,-1),2),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 10))
    solde = total_debit - total_credit # CORRIGE
    couleur_solde = "green" if solde >= 0 else "red"
    elements.append(Paragraph(f"<b><font color='green'>Total Débit : {total_debit:,.2f}</font></b>", styles["Heading5"]))
    elements.append(Paragraph(f"<b><font color='red'>Total Crédit : {total_credit:,.2f}</font></b>", styles["Heading5"]))
    elements.append(Paragraph(f"<b><font color='{couleur_solde}'>Solde (Débit-Crédit) : {solde:,.2f}</font></b>", styles["Heading5"]))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("Généré automatiquement par MSR", styles["Normal"]))
    doc.build(elements)
    messagebox.showinfo("Succès", "Rapport généré avec succès")
    try: os.startfile(file_path)
    except: pass

# -----------------------
# RAPPORTS
# -----------------------
def filtrer_transactions(par):
    today = datetime.date.today()
    transactions = lire_transactions()
    filtered = []
    for t in transactions:
        try:
            t_date = datetime.datetime.strptime(t[7], "%d-%m-%Y").date()
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
    tree_rap.tag_configure("entree", foreground="green")
    tree_rap.tag_configure("sortie", foreground="red")
    tree_rap.tag_configure("total", background="#E0E0E0", font=("Arial", 10, "bold"))
    total_debit = 0
    total_credit = 0
    for t in transactions:
        date = t[7]; description = t[3]; categorie = t[4]; type_tx = t[1]; montant = float(t[2]); compte = t[5]; mode = t[6]
        if type_tx.strip().lower() == "entrée":
            total_debit += montant
            tag = "entree"
        else:
            total_credit += montant
            tag = "sortie"
        tree_rap.insert("", "end", values=(date, description, categorie, type_tx, f"{montant:,.2f}", compte, mode), tags=(tag,))
    solde = total_debit - total_credit
    tree_rap.insert("", "end", values=("", "", "", "", "", "", ""))
    tree_rap.insert("", "end", values=("","TOTAL DEBIT","", "", f"{total_debit:,.2f} F","", ""), tags=("total",))
    tree_rap.insert("", "end", values=("","TOTAL CREDIT","", "", f"{total_credit:,.2f} F","", ""), tags=("total",))
    tree_rap.insert("", "end", values=("","SOLDE D-C","", "", f"{solde:,.2f} F","", ""), tags=("total",))
    tk.Label(fen, text=f"Total Débit: {total_debit:,.2f} F | Total Crédit: {total_credit:,.2f} F | Solde D-C: {solde:,.2f} F", font=("Arial",12,"bold")).pack(pady=5)
    tk.Button(fen, text="Imprimer PDF", command=lambda: imprimer_pdf(transactions, f"Historique_{titre}"), font=("Arial",12)).pack(pady=5)

#-----------------------
# RAPPORT PAR CATEGORIE
#-----------------------
def afficher_rapport_par_categorie(transactions):
    if not transactions:
        messagebox.showinfo("Rapport", "Aucune transaction")
        return
    rapport = defaultdict(lambda: defaultdict(list))
    for t in transactions:
        try:
            date_obj = datetime.datetime.strptime(t[7], "%d-%m-%Y")
        except:
            continue
        categorie = t[4]
        mois = date_obj.strftime("%m-%Y")
        rapport[categorie][mois].append(t)
    fen = tk.Toplevel(root)
    fen.title("Rapport détaillé par catégorie / mois")
    fen.geometry("1000x600")
    cols = ("Date", "Description", "Type", "Montant", "Compte", "Mode")
    tree_cat = ttk.Treeview(fen, columns=cols, show="tree headings")
    tree_cat.heading("#0", text="Catégorie / Mois")
    tree_cat.column("#0", width=200)
    for c in cols:
        tree_cat.heading(c, text=c)
        tree_cat.column(c, width=120, anchor="center")
    tree_cat.pack(fill="both", expand=True)
    tree_cat.tag_configure("categorie", background="#DDEEFF", font=("Arial", 10, "bold"))
    tree_cat.tag_configure("mois", background="#EEEEEE", font=("Arial", 9, "bold"))
    tree_cat.tag_configure("entree", foreground="green")
    tree_cat.tag_configure("sortie", foreground="red")
    tree_cat.tag_configure("total", background="#C8E6C9", font=("Arial", 10, "bold"))
    grand_total_debit = 0
    grand_total_credit = 0
    for cat in sorted(rapport.keys()):
        cat_id = tree_cat.insert("", "end", text=cat.upper(), values=("", "", "", "", "", ""), tags=("categorie",))
        total_cat_debit = 0
        total_cat_credit = 0
        for mois in sorted(rapport[cat].keys()):
            mois_id = tree_cat.insert(cat_id, "end", text=mois, values=("", "", "", "", "", ""), tags=("mois",))
            total_mois_debit = 0
            total_mois_credit = 0
            for t in rapport[cat][mois]:
                date = t[7]; desc = t[3]; type_tx = t[1]; montant = float(t[2]); compte = t[5]; mode = t[6]
                if type_tx.lower() == "entrée":
                    total_mois_debit += montant
                    tag = "entree"
                else:
                    total_mois_credit += montant
                    tag = "sortie"
                tree_cat.insert(mois_id, "end", text="", values=(date, desc, type_tx, f"{montant:,.2f}", compte, mode), tags=(tag,))
            solde_mois = total_mois_debit - total_mois_credit
            tree_cat.insert(mois_id, "end", text="TOTAL MOIS", values=("", "", "", f"{solde_mois:,.2f}", "", ""), tags=("total",))
            total_cat_debit += total_mois_debit
            total_cat_credit += total_mois_credit
        solde_cat = total_cat_debit - total_cat_credit
        tree_cat.insert(cat_id, "end", text="TOTAL " + cat.upper(), values=("", "", "", f"{solde_cat:,.2f}", "", ""), tags=("total",))
        grand_total_debit += total_cat_debit
        grand_total_credit += total_cat_credit
    solde_general = grand_total_debit - grand_total_credit
    tk.Label(fen, text=f"Total Débit: {grand_total_debit:,.2f} F | Total Crédit: {grand_total_credit:,.2f} F | Solde D-C: {solde_general:,.2f} F", font=("Arial", 12, "bold")).pack(pady=10)
    tk.Button(fen, text="Imprimer PDF (période)", command=imprimer_pdf_par_periode, bg="#2563EB", fg="white", font=("Arial", 11, "bold")).pack(pady=10)

def imprimer_pdf_par_periode():
    from tkinter import Toplevel, Label, Button
    from tkcalendar import DateEntry
    import datetime
    fen = Toplevel()
    fen.title("Choisir la période")
    fen.geometry("300x200")
    Label(fen, text="Date début").pack(pady=5)
    date_debut = DateEntry(fen, date_pattern='dd-mm-yyyy')
    date_debut.pack(pady=5)
    Label(fen, text="Date fin").pack(pady=5)
    date_fin = DateEntry(fen, date_pattern='dd-mm-yyyy')
    date_fin.pack(pady=5)
    def valider():
        d1 = date_debut.get_date()
        d2 = date_fin.get_date()
        transactions = lire_transactions()
        transactions_filtrees = []
        for t in transactions:
            try:
                d = datetime.datetime.strptime(str(t[7]), "%d-%m-%Y").date()
                if d1 <= d <= d2:
                    transactions_filtrees.append(t)
            except:
                continue
        if not transactions_filtrees:
            messagebox.showinfo("Info", "Aucune transaction pour cette période")
            return
        imprimer_pdf_par_categorie_detail(transactions_filtrees, f"Rapport financier du {d1.strftime('%d-%m-%Y')} au {d2.strftime('%d-%m-%Y')}")
        fen.destroy()
    Button(fen, text="Générer PDF", command=valider).pack(pady=15)

def filtrer_par_date(transactions, date_debut, date_fin):
    resultat = []
    for t in transactions:
        try:
            d = datetime.datetime.strptime(t[7], "%d-%m-%Y").date()
        except:
            continue
        if date_debut <= d <= date_fin:
            resultat.append(t)
    return resultat

#---------------
#DASHBOARD
#---------------
def afficher_dashboard():
    transactions = lire_transactions()
    if not transactions:
        messagebox.showinfo("Dashboard", "Aucune donnée disponible")
        return
    total_debit = 0
    total_credit = 0
    par_mois = defaultdict(lambda: {"entree": 0, "sortie": 0})
    par_categorie = defaultdict(lambda: {"entree": 0, "sortie": 0})
    montants = []
    for t in transactions:
        try:
            type_tx = t[1].strip().lower()
            montant = float(t[2])
            categorie = t[4]
            date_obj = datetime.datetime.strptime(t[7], "%d-%m-%Y")
            mois = date_obj.strftime("%m-%Y")
            montants.append(montant)
            if type_tx == "entrée":
                total_debit += montant
                par_mois[mois]["entree"] += montant
                par_categorie[categorie]["entree"] += montant
            elif type_tx == "sortie":
                total_credit += montant
                par_mois[mois]["sortie"] += montant
                par_categorie[categorie]["sortie"] += montant
        except:
            continue
    solde = total_debit - total_credit
    mois_sorted = sorted(par_mois.keys(), key=lambda x: datetime.datetime.strptime(x, "%m-%Y"))
    fen = tk.Toplevel(root)
    fen.title("Dashboard Financier Pro")
    fen.geometry("1350x850")
    fen.configure(bg="#F3F4F6")
    header = tk.Frame(fen, bg="#1E3A8A", height=60)
    header.pack(side="top", fill="x")
    tk.Label(header, text="TABLEAU DE BORD FINANCIER", bg="#1E3A8A", fg="white", font=("Arial", 15, "bold")).pack(pady=10)
    footer = tk.Frame(fen, bg="#E5E7EB", height=60)
    footer.pack(side="bottom", fill="x")
    content = tk.Frame(fen, bg="#F3F4F6")
    content.pack(side="top", fill="both", expand=True)
    fig = plt.Figure(figsize=(13, 7), dpi=100)
    gs = GridSpec(2, 2, figure=fig)
    ax1 = fig.add_subplot(gs[0, 0])
    def autopct_format(values):
        def my_format(pct):
            total = sum(values)
            val = int(round(pct * total / 100.0))
            return f"{pct:.1f}%\n({val:,.0f} F)"
        return my_format
    if total_debit+total_credit>0:
        ax1.pie([total_debit, total_credit], labels=["Débit", "Crédit"], autopct=autopct_format([total_debit, total_credit]), startangle=90, textprops={'fontsize': 9})
    ax1.set_title("Débit vs Crédit (Débit - Crédit)")
    ax2 = fig.add_subplot(gs[0, 1])
    entrees = [par_mois[m]["entree"] for m in mois_sorted]
    sorties = [par_mois[m]["sortie"] for m in mois_sorted]
    ax2.plot(mois_sorted, entrees, marker='o', label="Débit")
    ax2.plot(mois_sorted, sorties, marker='o', label="Crédit")
    ax2.legend(); ax2.grid(True, linestyle="--", alpha=0.5); ax2.tick_params(axis='x', rotation=45); ax2.set_title("Évolution Mensuelle")
    ax3 = fig.add_subplot(gs[1, 0])
    cats = list(par_categorie.keys()); vals = [par_categorie[c]["sortie"] for c in cats]
    bars = ax3.barh(cats, vals)
    for bar in bars:
        width = bar.get_width()
        ax3.text(width, bar.get_y() + bar.get_height()/2, f"{width:,.0f}", va='center', fontsize=9)
    ax3.set_title("Dépenses par catégorie")
    ax4 = fig.add_subplot(gs[1, 1]); ax4.axis("off")
    color = "green" if solde >= 0 else "red"
    ax4.text(0.1, 0.85, "INDICATEURS CLÉS", fontsize=13, fontweight="bold")
    ax4.text(0.1, 0.70, f"Solde D-C : {solde:,.0f} F", fontsize=12, color=color, fontweight="bold")
    ax4.text(0.1, 0.60, f"Transactions : {len(transactions)}", fontsize=10)
    fig.subplots_adjust(top=0.92, bottom=0.10, left=0.05, right=0.95, hspace=0.35, wspace=0.25)
    canvas = FigureCanvasTkAgg(fig, master=content); canvas.draw(); canvas.get_tk_widget().pack(fill="both", expand=True)
    def exporter_pdf():
        dossier = os.path.join(APP_FOLDER, "rapports"); os.makedirs(dossier, exist_ok=True)
        nom = datetime.datetime.now().strftime("dashboard_%d%m%Y_%H%M%S.pdf")
        chemin = os.path.join(dossier, nom)
        with PdfPages(chemin) as pdf:
            pdf.savefig(fig, bbox_inches='tight')
        messagebox.showinfo("Succès", f"PDF généré : {chemin}")
    def fermer(): fen.destroy()
    tk.Button(footer, text="📤 Exporter PDF", bg="#2563EB", fg="white", font=("Arial", 10, "bold"), command=exporter_pdf).pack(side="left", padx=20, pady=10)
    tk.Button(footer, text="❌ Fermer", bg="#DC2626", fg="white", font=("Arial", 10, "bold"), command=fermer).pack(side="right", padx=20, pady=10)

# EXPORTATIONS PDF EXCEL
def imprimer_pdf(transactions, titre):
    if not transactions:
        messagebox.showinfo("PDF", "Aucune transaction")
        return
    dossier_rapports = os.path.join(APP_FOLDER, "rapports")
    os.makedirs(dossier_rapports, exist_ok=True)
    date_str = datetime.datetime.now().strftime("%d%m%Y_%H%M%S")
    file_path = os.path.join(dossier_rapports, f"{titre}_{date_str}.pdf")
    doc = SimpleDocTemplate(file_path, pagesize=A4, topMargin=20, bottomMargin=40)
    elements = []; styles = getSampleStyleSheet()
    left_style = ParagraphStyle('left', parent=styles['Normal'], fontSize=7, leading=8, alignment=0)
    right_style = ParagraphStyle('right', parent=styles['Normal'], fontSize=7, leading=8, alignment=2)
    resume_style = ParagraphStyle('resume', parent=styles['Normal'], fontSize=9, leading=11, alignment=2)
    elements.append(Paragraph(titre, styles['Title'])); elements.append(Spacer(1, 5 * mm))
    data_pdf = []; header = ["Date", "Description", "Catégorie", "Type", "Montant", "Compte", "Mode"]
    data_pdf.append([Paragraph(h, styles["Heading6"]) for h in header])
    total_debit = total_credit = 0
    for t in transactions:
        type_tx = str(t[1]).strip(); montant = float(t[2])
        if type_tx.lower() == "entrée": total_debit += montant
        else: total_credit += montant
        data_pdf.append([Paragraph(str(t[7]), left_style), Paragraph(str(t[3]), left_style), Paragraph(str(t[4]), left_style), Paragraph(type_tx, left_style), Paragraph(f"{montant:,.2f}", right_style), Paragraph(str(t[5]), left_style), Paragraph(str(t[6]), left_style),])
    table = Table(data_pdf, colWidths=[55, 180, 65, 70, 70, 55, 55], repeatRows=1)
    table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2196F3")), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('GRID', (0, 0), (-1, -1), 0.3, colors.grey), ('ALIGN', (4, 1), (4, -1), 'RIGHT'),]))
    elements.append(table); elements.append(Spacer(1, 5 * mm))
    solde = total_debit - total_credit
    elements.append(Paragraph(f"<b>Total Débit :</b> <font color='green'>{total_debit:,.2f} F</font>", resume_style))
    elements.append(Paragraph(f"<b>Total Crédit :</b> <font color='red'>{total_credit:,.2f} F</font>", resume_style))
    couleur_solde = "green" if solde >= 0 else "red"
    elements.append(Paragraph(f"<b>Solde D-C :</b> <font color='{couleur_solde}'>{solde:,.2f} F</font>", resume_style))
    date_impression = datetime.datetime.today().strftime("%d-%m-%Y %H:%M")
    def add_page_number(canvas, doc):
        canvas.saveState(); canvas.setFont("Helvetica", 7); canvas.line(20 * mm, 18 * mm, A4[0] - 20 * mm, 18 * mm); canvas.drawString(20 * mm, 10 * mm, f"Date d'impression : {date_impression}"); canvas.drawRightString(A4[0] - 20 * mm, 10 * mm, f"Page {canvas.getPageNumber()}"); canvas.restoreState()
    doc.build(elements, onFirstPage=add_page_number, onLaterPages=add_page_number)
    messagebox.showinfo("Succès", f"PDF généré : {file_path}")

def exporter_excel(transactions):
    if not transactions:
        messagebox.showinfo("Excel", "Aucune transaction")
        return
    dossier_rapports = os.path.join(APP_FOLDER, "rapports")
    os.makedirs(dossier_rapports, exist_ok=True)
    date_str = datetime.datetime.now().strftime("%d%m%Y_%H%M%S")
    file_name = os.path.join(dossier_rapports, f"Transactions_{date_str}.xlsx")
    wb = Workbook(); ws = wb.active; ws.title = "Transactions"
    headers = ["Date","Description","Catégorie","Type","Débit","Crédit","Solde D-C","Compte","Mode"]
    ws.append(headers)
    def parse_date_excel(s):
        try: return datetime.datetime.strptime(s, "%d-%m-%Y")
        except: return datetime.datetime.min
    asc = sorted(transactions, key=lambda t: parse_date_excel(t[7]))
    cumul = 0; solde_map = {}
    for t in asc:
        cumul += float(t[2]) if t[1].lower()=="entrée" else -float(t[2])
        solde_map[t[0]] = cumul
    for t in sorted(transactions, key=lambda t: parse_date_excel(t[7]), reverse=True):
        debit = float(t[2]) if t[1].lower()=="entrée" else 0
        credit = 0 if t[1].lower()=="entrée" else float(t[2])
        ws.append([t[7], t[3], t[4], t[1], debit, credit, solde_map[t[0]], t[5], t[6]])
    wb.save(file_name); messagebox.showinfo("Succès", f"Export Excel généré : {file_name}"); os.startfile(file_name)

def imprimer_pdf_par_categorie_detail(transactions, titre):
    if not transactions:
        messagebox.showinfo("PDF", "Aucune transaction")
        return
    dossier = os.path.join(APP_FOLDER, "rapports")
    os.makedirs(dossier, exist_ok=True)
    file_path = os.path.join(dossier, f"{titre}_{datetime.datetime.now().strftime('%d%m%Y_%H%M%S')}.pdf")
    doc = SimpleDocTemplate(file_path, pagesize=A4, leftMargin=15, rightMargin=15, topMargin=20, bottomMargin=20)
    elements = []; styles = getSampleStyleSheet()
    titre_style = ParagraphStyle('titre', parent=styles['Normal'], fontSize=11, textColor=colors.red, fontName="Helvetica-Bold", spaceAfter=4)
    small_style = ParagraphStyle('small', parent=styles['Normal'], fontSize=7, fontName="Helvetica-Bold")
    elements.append(Paragraph(titre, titre_style)); elements.append(Spacer(1, 6))
    rapport = defaultdict(list)
    for t in transactions:
        cat = t[4] if t[4] else "Non classé"
        rapport[cat].append(t)
    grand_debit = 0; grand_credit = 0
    for cat in sorted(rapport.keys()):
        elements.append(Paragraph(cat, titre_style)); elements.append(Spacer(1, 2))
        def safe_date(tx):
            try: return datetime.datetime.strptime(str(tx[7]), "%d-%m-%Y")
            except: return datetime.datetime.min
        transactions_triees = sorted(rapport[cat], key=safe_date)
        data = [["Date", "Description", "Type", "Montant", "Compte", "Mode"]]
        total_debit = 0; total_credit = 0; table_styles = []
        for t in transactions_triees:
            date = str(t[7]); desc = str(t[3]); type_tx = str(t[1]); compte = str(t[5]); mode = str(t[6])
            try: montant = float(t[2])
            except: montant = 0
            if type_tx.lower() == "entrée": total_debit += montant; color = colors.green
            else: total_credit += montant; color = colors.red
            data.append([date, desc, type_tx, f"{montant:,.0f}", compte, mode])
            table_styles.append(('TEXTCOLOR', (3, len(data)-1), (3, len(data)-1), color))
        table = Table(data, colWidths=[70, 200, 60, 70, 70, 70], repeatRows=1)
        style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('ALIGN', (3, 0), (3, -1), 'RIGHT'), ('GRID', (0, 0), (-1, -1), 0.25, colors.grey), ('FONTSIZE', (0, 0), (-1, -1), 6), ('TOPPADDING', (0, 0), (-1, -1), 1), ('BOTTOMPADDING', (0, 0), (-1, -1), 1),])
        for s in table_styles: style.add(*s)
        table.setStyle(style); elements.append(table)
        solde = total_debit - total_credit
        elements.append(Spacer(1, 4))
        elements.append(Paragraph(f"Débit : {total_debit:,.0f} F", ParagraphStyle("e", parent=small_style, textColor=colors.green)))
        elements.append(Paragraph(f"Crédit : {total_credit:,.0f} F", ParagraphStyle("s", parent=small_style, textColor=colors.red)))
        elements.append(Paragraph(f"Solde D-C : {solde:,.0f} F", ParagraphStyle("solde", parent=small_style, textColor=(colors.green if solde >= 0 else colors.red))))
        elements.append(Spacer(1, 8)); grand_debit += total_debit; grand_credit += total_credit
    solde_final = grand_debit - grand_credit
    elements.append(Spacer(1, 10)); elements.append(Paragraph("RAPPORT FINAL", titre_style))
    elements.append(Paragraph(f"Total Débit : {grand_debit:,.0f} F", ParagraphStyle("ge", parent=small_style, textColor=colors.green)))
    elements.append(Paragraph(f"Total Crédit : {grand_credit:,.0f} F", ParagraphStyle("gs", parent=small_style, textColor=colors.red)))
    elements.append(Paragraph(f"Solde Global D-C : {solde_final:,.0f} F", ParagraphStyle("sg", parent=small_style, textColor=(colors.green if solde_final >= 0 else colors.red))))
    def footer(canvas, doc):
        canvas.saveState(); canvas.setFont("Helvetica", 7); canvas.drawString(15, 15, "M Sainte Rita - Rapport comptable"); canvas.drawRightString(575, 15, f"Page {canvas.getPageNumber()}"); canvas.restoreState()
    doc.build(elements, onFirstPage=footer, onLaterPages=footer)
    messagebox.showinfo("Succès", f"PDF généré : {file_path}"); os.startfile(file_path)

# INTERFACE PRINCIPALE
root = tk.Tk()
root.title("Application Comptabilité Magasin Sainte Rita")
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
root.minsize(800, 500)
root.geometry(f"{int(screen_width*0.9)}x{int(screen_height*0.9)}")
font_label = ("Arial", 12); font_entry = ("Arial", 12); font_button = ("Arial", 12, "bold")
frame_top = tk.Frame(root); frame_top.pack(fill="both", expand=True, padx=10, pady=10)
frame_top = tk.Frame(root); frame_top.pack(fill="x", expand=True, padx=20, pady=10)
frame_saisie = tk.LabelFrame(frame_top, text="Gestion des Transactions", padx=20, pady=20, font=("Arial",14,"bold"))
frame_saisie.pack(side="left", fill="x", expand=True, padx=(0,10))
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
ttk.Combobox(frame_saisie,textvariable=mode_var,values=MODES_PAIEMENT,font=font_entry,width=15,state="readonly").grid(row=1, column=5, padx=5, pady=5)
tk.Label(frame_saisie,text="Mode",font=font_label).grid(row=1, column=4, sticky="w", padx=5, pady=5)
date_entry = DateEntry(frame_saisie, font=font_entry, width=15, date_pattern='dd-mm-yyyy')
date_entry.grid(row=2,column=1, padx=5, pady=10)
tk.Label(frame_saisie,text="Date", font=font_label).grid(row=2,column=0, sticky="w", padx=5, pady=10)
tk.Button(frame_saisie,text="Ajouter/Modifier", command=ajouter_transaction, bg="#4CAF50",fg="white", font=font_button, width=15).grid(row=2,column=3, padx=5, pady=10)
tk.Button(frame_saisie,text="Supprimer", command=demander_mot_de_passe, bg="#f44336",fg="white", font=font_button, width=15).grid(row=2,column=4, padx=5, pady=10)
tk.Button(frame_saisie,text="Modifier", command=modifier_transaction, bg="#FF9800",fg="white", font=font_button, width=15).grid(row=2,column=5, padx=5, pady=10)

frame_recherche = tk.LabelFrame(frame_top, text="Recherche", padx=10, pady=10, font=("Arial",12,"bold"))
frame_recherche.pack(side="right")
recherche_var = tk.StringVar()
tk.Label(frame_recherche, text="Recherche", font=font_label).grid(row=0, column=0, padx=5, pady=5)
recherche_entry = tk.Entry(frame_recherche, textvariable=recherche_var, font=font_entry, width=20)
recherche_entry.grid(row=0, column=1, padx=5, pady=5)

def charger_transactions():
    for item in tree.get_children():
        tree.delete(item)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT id, type, montant, description, categorie, compte, mode_paiement, date FROM transactions ORDER BY date DESC")
    rows = cursor.fetchall()
    conn.close()
    # reutilise la logique corrigée
    mise_a_jour_tableau(rows)

def rechercher_transaction(event=None):
    mot = recherche_var.get().lower()
    if mot == "":
        charger_transactions()
        return
    for item in tree.get_children():
        valeurs = tree.item(item)["values"]
        texte = " ".join(str(v).lower() for v in valeurs)
        if mot not in texte:
            tree.delete(item)

recherche_entry.bind("<KeyRelease>", rechercher_transaction)

frame_top = tk.Frame(root); frame_top.pack(fill="x", padx=20, pady=10)
frame_filtre = tk.LabelFrame(frame_top, text="Filtrage", padx=10, pady=10, font=("Arial",12,"bold"))
frame_filtre.pack(side="left", padx=10, pady=5)
filtre_categorie = tk.StringVar(value="")
ttk.Combobox(frame_filtre, textvariable=filtre_categorie, values=CATEGORIES, font=font_entry, width=12).grid(row=0,column=1, padx=5, pady=5)
tk.Label(frame_filtre,text="Catégorie:", font=font_label).grid(row=0,column=0, padx=5, pady=5)
tk.Label(frame_filtre, text="Date début:", font=font_label).grid(row=1, column=0, padx=5, pady=5)
filtre_date_debut = DateEntry(frame_filtre, font=font_entry, width=12, date_pattern='dd-mm-yyyy')
filtre_date_debut.grid(row=1, column=1, padx=5)
tk.Label(frame_filtre, text="Date fin:", font=font_label).grid(row=2, column=0, padx=5, pady=5)
filtre_date_fin = DateEntry(frame_filtre, font=font_entry, width=12, date_pattern='dd-mm-yyyy')
filtre_date_fin.grid(row=2, column=1, padx=5, pady=5)
filtre_type = tk.StringVar(value="")
tk.Label(frame_filtre, text="Type:", font=font_label).grid(row=0, column=2, padx=5, pady=5)
ttk.Combobox(frame_filtre, textvariable=filtre_type, values=TYPES_TRANSACTION, font=font_entry, width=12).grid(row=0, column=3, padx=5, pady=5)
filtre_compte = tk.StringVar(value="")
tk.Label(frame_filtre, text="Compte:", font=font_label).grid(row=1, column=2, padx=5, pady=5)
ttk.Combobox(frame_filtre, textvariable=filtre_compte, values=COMPTES, font=font_entry, width=12).grid(row=1, column=3, padx=5, pady=5)
filtre_mode = tk.StringVar(value="")
tk.Label(frame_filtre, text="Mode:", font=font_label).grid(row=2, column=2, padx=5, pady=5)
ttk.Combobox(frame_filtre, textvariable=filtre_mode, values=MODES_PAIEMENT, font=font_entry, width=12).grid(row=2, column=3, padx=5, pady=5)
tk.Button(frame_filtre,text="Rechercher",command=rechercher_transactions,font=font_button,bg="#2196F3",fg="white", width=15).grid(row=4,column=0,columnspan=2,pady=10)
tk.Button(frame_filtre,text="Tout afficher",command=tout_afficher,font=font_button,bg="#607D8B",fg="white", width=15).grid(row=4,column=2,columnspan=2,pady=5)
tk.Button(frame_filtre, text="Imprimer sélection", command=imprimer_selection, bg="#2196F3", fg="white", font=("Arial",11,"bold"), width=20).grid(row=4, column=4, columnspan=4, pady=10)

resume_frame = tk.LabelFrame(frame_top, text="Résumé", padx=10, pady=10, font=("Arial",12,"bold"))
resume_frame.pack(side="left", padx=10, pady=5, fill="y")
resume_label = tk.Label(resume_frame, text="", font=("Arial",11,"bold"), justify="left", anchor="w")
resume_label.pack(padx=5, pady=5, anchor="w")

def mise_a_jour_resume(transactions=None):
    if transactions is None:
        transactions = lire_transactions()
    total_debit = sum(float(t[2]) for t in transactions if t[1]=="Entrée")
    total_credit = sum(float(t[2]) for t in transactions if t[1]=="Sortie")
    solde = total_debit - total_credit
    resume_label.config(text=f"Débit total:\n {total_debit:,.2f} F\n\nCrédit total:\n {total_credit:,.2f} F\n\nSolde D-C:\n {solde:,.2f} F", fg="black")
    mise_a_jour_infos(transactions)

infos_frame = tk.LabelFrame(frame_top, text="Informations", padx=10, pady=10, font=("Arial",12,"bold"))
infos_frame.pack(side="left", padx=10, pady=5, fill="y")
infos_label = tk.Label(infos_frame, text="", font=("Arial",11), justify="left", anchor="w")
infos_label.pack(padx=5, pady=5, anchor="w")

def mise_a_jour_infos(transactions=None):
    if transactions is None:
        transactions = lire_transactions()
    nb_transactions = len(transactions)
    derniere_trans = max(transactions, key=lambda x: x[0]) if transactions else None
    if derniere_trans:
        dernier_detail = (f"{derniere_trans[0]}\nType: {derniere_trans[1]}\nMontant: {derniere_trans[2]:,.2f} F")
    else:
        dernier_detail = "Aucune"
    infos_label.config(text=f"Nb total de trans: {nb_transactions}\nDernière trans numéro: {dernier_detail}")

def imprimer_bilan_detaille_pdf():
    transactions = lire_transactions()
    if not transactions:
        messagebox.showinfo("PDF", "Aucune donnée")
        return
    from collections import defaultdict
    import os, datetime
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    produits = defaultdict(float)
    charges = defaultdict(float)
    count_produits = defaultdict(int)
    count_charges = defaultdict(int)
    montants = []
    for t in transactions:
        try:
            type_tx = t[1].strip().lower()
            montant = float(t[2])
            categorie = t[4]
            montants.append(montant)
            if type_tx == "entrée":
                produits[categorie] += montant
                count_produits[categorie] += 1
            elif type_tx == "sortie":
                charges[categorie] += montant
                count_charges[categorie] += 1
        except:
            continue
    total_produits = sum(produits.values())
    total_charges = sum(charges.values())
    resultat = total_produits - total_charges
    dossier = os.path.join(APP_FOLDER, "rapports")
    os.makedirs(dossier, exist_ok=True)
    file_path = os.path.join(dossier, f"Bilan_{datetime.datetime.now().strftime('%d%m%Y_%H%M%S')}.pdf")
    doc = SimpleDocTemplate(file_path)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", parent=styles["Title"], alignment=1, fontSize=14)
    normal = ParagraphStyle("normal", parent=styles["Normal"], fontSize=9, leading=12)
    elements.append(Paragraph("BILAN FINANCIER GLOBAL", title_style))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph(f"Date : {datetime.datetime.now().strftime('%d-%m-%Y %H:%M')}", normal))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("RÉSUMÉ FINANCIER", styles["Heading3"]))
    elements.append(Paragraph(f"Débit total : <b>{total_produits:,.2f} F</b><br/>Crédit total : <b>{total_charges:,.2f} F</b><br/>Solde D-C : <b>{resultat:,.2f} F</b><br/>Total transactions : <b>{len(transactions)}</b>", normal))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("1. ENTRÉES (DEBIT)", styles["Heading3"]))
    data1 = [["Catégorie", "Nb transactions", "Montant (F)", "%"]]
    sorted_produits = sorted(produits.items(), key=lambda x: x[1], reverse=True)
    for cat, val in sorted_produits:
        nb = count_produits[cat]
        pct = (val / total_produits * 100) if total_produits else 0
        data1.append([cat, nb, f"{val:,.2f}", f"{pct:.1f}%"])
    data1.append(["TOTAL DEBIT", sum(count_produits.values()), f"{total_produits:,.2f}", "100%"])
    table1 = Table(data1, colWidths=[180, 110, 120, 70])
    table1.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.green), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.4, colors.grey), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('ALIGN', (1,1), (-1,-1), 'CENTER'), ('ALIGN', (2,1), (-1,-1), 'RIGHT')]))
    elements.append(table1)
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("2. SORTIES (CREDIT)", styles["Heading3"]))
    data2 = [["Catégorie", "Nb transactions", "Montant (F)", "%"]]
    sorted_charges = sorted(charges.items(), key=lambda x: x[1], reverse=True)
    for cat, val in sorted_charges:
        nb = count_charges[cat]
        pct = (val / total_charges * 100) if total_charges else 0
        data2.append([cat, nb, f"{val:,.2f}", f"{pct:.1f}%"])
    data2.append(["TOTAL CREDIT", sum(count_charges.values()), f"{total_charges:,.2f}", "100%"])
    table2 = Table(data2, colWidths=[180, 110, 120, 70])
    table2.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.red), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.4, colors.grey), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('ALIGN', (1,1), (-1,-1), 'CENTER'), ('ALIGN', (2,1), (-1,-1), 'RIGHT')]))
    elements.append(table2)
    elements.append(Spacer(1, 15))
    elements.append(Paragraph("RÉCAPITULATIF GLOBAL", styles["Heading3"]))
    solde = total_produits - total_charges
    data_final = [["Indicateur", "Montant (F)"], ["Total Débit", f"{total_produits:,.2f}"], ["Total Crédit", f"{total_charges:,.2f}"], ["Solde D-C", f"{solde:,.2f}"]]
    table_final = Table(data_final, colWidths=[250, 150])
    table_final.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1E3A8A")), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('FONTSIZE', (0,0), (-1,-1), 10), ('TEXTCOLOR', (1,1), (1,1), colors.green), ('TEXTCOLOR', (1,2), (1,2), colors.red),]))
    table_final.setStyle([('TEXTCOLOR', (1,3), (1,3), colors.green if solde >= 0 else colors.red)])
    elements.append(table_final)
    def footer(canvas, doc):
        canvas.saveState(); canvas.setFont("Helvetica", 8); canvas.drawString(20, 15, "Bilan financier automatique"); canvas.drawRightString(570, 15, f"Page {canvas.getPageNumber()}"); canvas.restoreState()
    doc.build(elements, onFirstPage=footer, onLaterPages=footer)
    messagebox.showinfo("Succès", f"Bilan PDF généré:\n{file_path}")
    os.startfile(file_path)

frame_rapports = tk.LabelFrame(frame_top, text="Rapports", padx=10, pady=10, font=("Arial",12,"bold"))
frame_rapports.pack(side="right", padx=10)
frame_rapports.columnconfigure(0, weight=1)
frame_rapports.columnconfigure(1, weight=1)
tk.Button(frame_rapports, text="Journalier", command=lambda: afficher_rapport(filtrer_transactions("jour"), "Journalier"), font=font_button, bg="#2196F3", fg="white").grid(row=0, column=0, padx=5, pady=5, sticky="ew")
tk.Button(frame_rapports, text="Hebdomadaire", command=lambda: afficher_rapport(filtrer_transactions("semaine"), "Hebdomadaire"), font=font_button, bg="#2196F3", fg="white").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
tk.Button(frame_rapports, text="Mensuel", command=lambda: afficher_rapport(filtrer_transactions("mois"), "Mensuel"), font=font_button, bg="#2196F3", fg="white").grid(row=1, column=0, padx=5, pady=5, sticky="ew")
tk.Button(frame_rapports, text="Annuel", command=lambda: afficher_rapport(filtrer_transactions("annee"), "Annuel"), font=font_button, bg="#2196F3", fg="white").grid(row=1, column=1, padx=5, pady=5, sticky="ew")
tk.Button(frame_rapports, text="Complet", command=lambda: afficher_rapport(lire_transactions(), "Complet"), font=font_button, bg="#4CAF50", fg="white").grid(row=2, column=0, padx=5, pady=5, sticky="ew")
tk.Button(frame_rapports, text="Export Excel", command=lambda: exporter_excel(lire_transactions()), font=font_button, bg="#FF9800", fg="white").grid(row=2, column=1, padx=5, pady=5, sticky="ew")
tk.Button(frame_rapports, text="Par Catégorie", command=lambda: afficher_rapport_par_categorie(lire_transactions()), font=font_button, bg="#9C27B0", fg="white").grid(row=3, column=0, padx=5, pady=5, sticky="ew")
tk.Button(frame_rapports, text="Dashboard", command=afficher_dashboard, font=font_button, bg="#000000", fg="white").grid(row=3, column=1, padx=5, pady=5, sticky="ew")
tk.Button(frame_rapports, text="Bilan PDF", command=imprimer_bilan_detaille_pdf, font=font_button, bg="#1F2937", fg="white").grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

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

mise_a_jour_tableau()
root.mainloop()