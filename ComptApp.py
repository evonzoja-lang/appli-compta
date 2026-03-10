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
from openpyxl import Workbook
from openpyxl.styles import Alignment
from collections import defaultdict
from reportlab.lib import pagesizes
import os
import sys


# dossier de l'application (là où se trouve l'exe ou le script)

def get_base_path():
    if getattr(sys, 'frozen', False):
        # si application compilée en .exe
        return os.path.dirname(sys.executable)
    else:
        # si script python
        return os.path.dirname(os.path.abspath(__file__))
    
BASE_DIR = get_base_path()
DB_FILE = os.path.join(BASE_DIR, "comptabilite.db")
print("Base utilisée :", DB_FILE)


DOCUMENTS = os.path.join(os.path.expanduser("~"), "Documents")
APP_FOLDER = os.path.join(DOCUMENTS, "ComptabiliteMSC")
RAPPORTS_FOLDER = os.path.join(APP_FOLDER, "rapports")

# créer les dossiers s'ils n'existent pas
os.makedirs(RAPPORTS_FOLDER, exist_ok=True)

def get_app_folder():
    # dossier Documents de l'utilisateur
    documents = os.path.join(os.path.expanduser("~"), "Documents")
    app_folder = os.path.join(documents, "ComptabiliteMSC")

    if not os.path.exists(app_folder):
        os.makedirs(app_folder)

    return app_folder

APP_FOLDER = get_app_folder()



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
        if tree.exists(k):   # évite l'erreur
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
        text=f"Tot Entrées: {total_entrees:.2f} F   Tot Sorties: {total_sorties:.2f} F   Solde: {solde_cumul:.2f} F"
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

    tk.Button(fenetre_mdp, text="Valider", command=verifier,
              bg="#4CAF50", fg="white", width=12).pack(pady=10)

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
    

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    requete = "SELECT * FROM transactions"
    conditions = []
    params = []

    # --- NORMALISATION ---
    categorie = filtre_categorie.get().strip()
    type_tx = filtre_type.get().strip()
    compte = filtre_compte.get().strip()
    mode = filtre_mode.get().strip()

    date_debut = filtre_date_debut.get().strip()
    date_fin = filtre_date_fin.get().strip()

    # -----------------------
    # FILTRES TEXTE (indépendants)
    # -----------------------
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

    # -----------------------
    # FILTRES DATE (flexibles)
    # -----------------------
    try:
        # Priorité aux plages de dates
        if date_debut or date_fin:
            if date_debut:
                date_debut_sql = datetime.datetime.strptime(date_debut, "%d-%m-%Y").strftime("%Y-%m-%d")
                conditions.append("date >= ?")
                params.append(date_debut_sql)
            if date_fin:
                date_fin_sql = datetime.datetime.strptime(date_fin, "%d-%m-%Y").strftime("%Y-%m-%d")
                conditions.append("date <= ?")
                params.append(date_fin_sql)
        # Date exacte uniquement si pas de plage
        elif date_exacte:
            date_sql = datetime.datetime.strptime(date_exacte, "%d-%m-%Y").strftime("%Y-%m-%d")
            conditions.append("date = ?")
            params.append(date_sql)

    except ValueError:
        messagebox.showerror("Erreur", "Format de date invalide (DD-MM-YYYY).")
        conn.close()
        return

    # -----------------------
    # Construction finale de la requête
    # -----------------------
    if conditions:
        requete += " WHERE " + " AND ".join(conditions)

    requete += " ORDER BY date DESC"

    # --- DEBUG ---
    print("REQUETE SQL:", requete)
    print("PARAMS:", params)

    # Exécution
    cursor.execute(requete, params)
    transactions = cursor.fetchall()
    conn.close()

    # Mise à jour du tableau Tkinter
    mise_a_jour_tableau(transactions)

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

    # 📁 Dossier rapports
    dossier = os.path.join(APP_FOLDER, "rapports")
    if not os.path.exists(dossier):
        os.makedirs(dossier, exist_ok=True)

    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = os.path.join(dossier, f"rapport_filtre_{now}.pdf")

    doc = SimpleDocTemplate(file_path, pagesize=pagesizes.A4,
                            rightMargin=20, leftMargin=20,
                            topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()

    # ===== EN-TÊTE =====
    elements.append(Paragraph("<b>INT MSC - RAPPORT FILTRE</b>", styles["Title"]))
    elements.append(Spacer(1, 10))
    date_rapport = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    elements.append(Paragraph(f"Date du rapport : {date_rapport}", styles["Normal"]))
    elements.append(Spacer(1, 20))

    # ===== TABLEAU =====
    data = []
    headers = tree["columns"][:-2]
    data.append(headers)

    total_debit = 0
    total_credit = 0

    for item in items:
        values = tree.item(item)["values"][:-2]
        data.append(values)

        try:
            debit = float(values[5])   # colonne Débit
            credit = float(values[6])  # colonne Crédit
            total_debit += debit
            total_credit += credit
        except:
            pass

    # ===== TABLEAU STYLE =====
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

    # ===== TOTALS =====
    solde = total_debit - total_credit


# Choix dynamique de la couleur du solde
    couleur_solde = "green" if solde >= 0 else "red"

    elements.append(Paragraph(
    f"<b><font color='green'>Total Entrées : {total_debit:,.2f}</font></b>",
    styles["Heading5"]
    ))

    elements.append(Paragraph(
    f"<b><font color='red'>Total Sorties  : {total_credit:,.2f}</font></b>",
    styles["Heading5"]
    ))

    elements.append(Paragraph(
    f"<b><font color='{couleur_solde}'>Solde : {solde:,.2f}</font></b>",
    styles["Heading5"]
    ))

    elements.append(Spacer(1, 20))

    # ===== PIED DE PAGE =====
    elements.append(Paragraph("Généré automatiquement par INT MSC", styles["Normal"]))

    # ==== GÉNÉRATION PDF ====
    doc.build(elements)
    messagebox.showinfo("Succès", "Rapport généré avec succès")
    os.startfile(file_path)
    
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


#-----------------------
# RAPPORT PAR CATEGORIE
#-----------------------
def afficher_rapport_par_categorie(transactions):
    if not transactions:
        messagebox.showinfo("Rapport par Catégorie", "Aucune transaction")
        return

    

    # Structure : {categorie: {mois: {"entree": x, "sortie": y}}}
    rapport = defaultdict(lambda: defaultdict(lambda: {"entree": 0, "sortie": 0}))

    for t in transactions:
        categorie = t[4]
        type_tx = t[1].strip().lower()
        montant = float(t[2])

        try:
            t_date = datetime.datetime.strptime(t[7], "%Y-%m-%d")
        except:
            continue

        mois = t_date.strftime("%Y-%m")

        if type_tx == "entrée":
            rapport[categorie][mois]["entree"] += montant
        elif type_tx == "sortie":
            rapport[categorie][mois]["sortie"] += montant

    # --- Fenêtre ---
    fen = tk.Toplevel(root)
    fen.title("Rapport par Catégorie par Mois")
    fen.geometry("950x550")

    cols = ("Catégorie / Mois", "Total Entrées", "Total Sorties", "Solde")

    tree_cat = ttk.Treeview(fen, columns=cols, show="headings")
    scroll = ttk.Scrollbar(fen, orient="vertical", command=tree_cat.yview)
    tree_cat.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")

    for c in cols:
        tree_cat.heading(c, text=c)
        tree_cat.column(c, width=200, anchor="center")

    tree_cat.pack(fill="both", expand=True, padx=10, pady=10)

    tree_cat.tag_configure("categorie", background="#DDEEFF", font=("Arial", 10, "bold"))
    tree_cat.tag_configure("total_cat", background="#EEEEEE", font=("Arial", 10, "bold"))
    tree_cat.tag_configure("total_general", background="#C8E6C9", font=("Arial", 11, "bold"))

    grand_total_ent = 0
    grand_total_sort = 0

    # ---- Affichage structuré ----
    for cat in sorted(rapport.keys()):

        # Ligne titre catégorie
        tree_cat.insert("", "end", values=(f"=== {cat.upper()} ===", "", "", ""), tags=("categorie",))

        total_cat_ent = 0
        total_cat_sort = 0

        for mois in sorted(rapport[cat].keys()):
            ent = rapport[cat][mois]["entree"]
            sort = rapport[cat][mois]["sortie"]
            solde = ent - sort

            total_cat_ent += ent
            total_cat_sort += sort

            tree_cat.insert(
                "",
                "end",
                values=(mois, f"{ent:,.2f}", f"{sort:,.2f}", f"{solde:,.2f}")
            )

        solde_cat = total_cat_ent - total_cat_sort

        # Ligne total catégorie
        tree_cat.insert(
            "",
            "end",
            values=("TOTAL " + cat.upper(),
                    f"{total_cat_ent:,.2f}",
                    f"{total_cat_sort:,.2f}",
                    f"{solde_cat:,.2f}"),
            tags=("total_cat",)
        )

        tree_cat.insert("", "end", values=("", "", "", ""))

        grand_total_ent += total_cat_ent
        grand_total_sort += total_cat_sort

    solde_general = grand_total_ent - grand_total_sort

    # Ligne total général
    tree_cat.insert(
        "",
        "end",
        values=("TOTAL GENERAL",
                f"{grand_total_ent:,.2f}",
                f"{grand_total_sort:,.2f}",
                f"{solde_general:,.2f}"),
        tags=("total_general",)
    )

    # Résumé en bas
    lbl_tot = tk.Label(
        fen,
        text=f"Total Entrées: {grand_total_ent:,.2f} F | "
             f"Total Sorties: {grand_total_sort:,.2f} F | "
             f"Solde Général: {solde_general:,.2f} F",
        font=("Arial", 12, "bold")
    )
    lbl_tot.pack(pady=8)

    tk.Button(
        fen,
        text="Imprimer PDF",
        command=lambda: imprimer_pdf_par_categorie_tableau(transactions, "Rapport_Categorie_Par_Mois"),
        font=("Arial", 12)
    ).pack(pady=5)
    
    
def imprimer_pdf_par_categorie_tableau(transactions, titre):
    if not transactions:
        messagebox.showinfo("PDF", "Aucune transaction")
        return

    # 📁 Dossier rapports
    dossier_rapports = os.path.join(APP_FOLDER, "rapports")
    if not os.path.exists(dossier_rapports):
        os.makedirs(dossier_rapports, exist_ok=True)

    # ⏰ Nom fichier PDF
    date_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = os.path.join(dossier_rapports, f"{titre}_{date_str}.pdf")

    # 📄 Création du document avec marges réduites
    doc = SimpleDocTemplate(file_path, pagesize=A4,
                            topMargin=10, bottomMargin=10, leftMargin=10, rightMargin=10)
    elements = []
    styles = getSampleStyleSheet()

    # Styles personnalisés
    titre_style = ParagraphStyle('title', parent=styles['Title'], fontSize=14, alignment=1)
    cat_style = ParagraphStyle('categorie', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor("#2196F3"))
    right_style = ParagraphStyle('right', parent=styles['Normal'], fontSize=8, alignment=2)
    left_style = ParagraphStyle('left', parent=styles['Normal'], fontSize=8, alignment=0)
    
    # Titre du rapport
    elements.append(Paragraph(titre, titre_style))
    elements.append(Spacer(1, 3 * mm))

    # Organisation des transactions par catégorie et mois
    rapport = defaultdict(lambda: defaultdict(lambda: {"entree": 0, "sortie": 0}))
    for t in transactions:
        cat = t[4]
        type_tx = t[1].strip().lower()
        montant = float(t[2])
        try:
            t_date = datetime.datetime.strptime(t[7], "%Y-%m-%d")
        except:
            continue
        mois = t_date.strftime("%Y-%m")
        if type_tx == "entrée":
            rapport[cat][mois]["entree"] += montant
        elif type_tx == "sortie":
            rapport[cat][mois]["sortie"] += montant

    grand_total_ent = 0
    grand_total_sort = 0

    usable_width = A4[0] - doc.leftMargin - doc.rightMargin  # largeur utilisable pour les tableaux

    for cat in sorted(rapport.keys()):
        elements.append(Paragraph(f"{cat.upper()}", cat_style))
        elements.append(Spacer(1, 1.5 * mm))

        data_pdf = []
        header = ["Mois", "Entrées", "Sorties", "Solde"]
        data_pdf.append([Paragraph(h, styles["Heading6"]) for h in header])

        total_cat_ent = 0
        total_cat_sort = 0

        for mois in sorted(rapport[cat].keys()):
            ent = rapport[cat][mois]["entree"]
            sort = rapport[cat][mois]["sortie"]
            solde = ent - sort
            total_cat_ent += ent
            total_cat_sort += sort

            data_pdf.append([
                Paragraph(mois, left_style),
                Paragraph(f"{ent:,.2f}", ParagraphStyle('entree', textColor=colors.green, fontSize=8, alignment=2)),
                Paragraph(f"{sort:,.2f}", ParagraphStyle('sortie', textColor=colors.red, fontSize=8, alignment=2)),
                Paragraph(f"{solde:,.2f}", right_style)
            ])

        solde_cat = total_cat_ent - total_cat_sort
        # Ligne total catégorie
        data_pdf.append([
            Paragraph("TOTAL " + cat.upper(), left_style),
            Paragraph(f"{total_cat_ent:,.2f}", right_style),
            Paragraph(f"{total_cat_sort:,.2f}", right_style),
            Paragraph(f"{solde_cat:,.2f}", right_style)
        ])

        col_widths_cat = [usable_width * 0.4, usable_width * 0.2, usable_width * 0.2, usable_width * 0.2]
        table = Table(data_pdf, colWidths=col_widths_cat, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2196F3")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
            ('BACKGROUND', (0,1), (-1,-2), colors.whitesmoke),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#EEEEEE")),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 3 * mm))

        grand_total_ent += total_cat_ent
        grand_total_sort += total_cat_sort

    solde_general = grand_total_ent - grand_total_sort

    # ===================== TOTAL GENERAL AVEC COULEURS =====================
    couleur_solde = colors.green if solde_general > 0 else (colors.yellow if solde_general == 0 else colors.red)
    elements.append(Spacer(1, 3 * mm))
    data_tot = [
        ["Entrées", "Sorties", "Solde"],
        [
            Paragraph(f"{grand_total_ent:,.2f}", ParagraphStyle('entree', textColor=colors.green, fontSize=9, alignment=2)),
            Paragraph(f"{grand_total_sort:,.2f}", ParagraphStyle('sortie', textColor=colors.red, fontSize=9, alignment=2)),
            Paragraph(f"{solde_general:,.2f}", ParagraphStyle('solde', textColor=couleur_solde, fontSize=9, alignment=2))
        ]
    ]
    col_widths_tot = [usable_width / 3] * 3
    table_tot = Table(data_tot, colWidths=col_widths_tot)
    table_tot.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#CCCCCC")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    elements.append(table_tot)

    # ===================== FOOTER =====================
    date_impression = datetime.datetime.today().strftime("%d-%m-%Y %H:%M")
    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.line(10 * mm, 15 * mm, A4[0] - 10 * mm, 15 * mm)
        canvas.drawString(10 * mm, 7 * mm, f"Date d'impression : {date_impression}")
        page_num = canvas.getPageNumber()
        canvas.drawRightString(A4[0] - 10 * mm, 7 * mm, f"Page {page_num}")
        canvas.restoreState()

    doc.build(elements, onFirstPage=add_page_number, onLaterPages=add_page_number)
    messagebox.showinfo("Succès", f"PDF généré : {file_path}")


# -----------------------
# EXPORTATIONS PDF EXCEL
# -----------------------

def imprimer_pdf(transactions, titre):
    if not transactions:
        messagebox.showinfo("PDF", "Aucune transaction")
        return

    # 📁 Dossier rapports
    dossier_rapports = os.path.join(APP_FOLDER, "rapports")
    if not os.path.exists(dossier_rapports):
        os.makedirs(dossier_rapports, exist_ok=True)

    # ⏰ Nom fichier PDF
    date_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = os.path.join(dossier_rapports, f"{titre}_{date_str}.pdf")

    # 📄 Création du document
    doc = SimpleDocTemplate(
        file_path,
        pagesize=A4,
        topMargin=20,
        bottomMargin=40
    )

    elements = []
    styles = getSampleStyleSheet()

    left_style = ParagraphStyle('left', parent=styles['Normal'], fontSize=7, leading=8, alignment=0)
    right_style = ParagraphStyle('right', parent=styles['Normal'], fontSize=7, leading=8, alignment=2)
    resume_style = ParagraphStyle('resume', parent=styles['Normal'], fontSize=9, leading=11, alignment=2)

    # ===================== TITRE =====================
    elements.append(Paragraph(titre, styles['Title']))
    elements.append(Spacer(1, 5 * mm))

    # ===================== TABLEAU =====================
    data_pdf = []
    header = ["Date", "Description", "Catégorie", "Type", "Montant", "Compte", "Mode"]
    data_pdf.append([Paragraph(h, styles["Heading6"]) for h in header])

    total_ent = total_sort = 0

    for t in transactions:
        type_tx = str(t[1]).strip()
        montant = float(t[2])
        if type_tx.lower() == "entrée":
            total_ent += montant
        elif type_tx.lower() == "sortie":
            total_sort += montant

        data_pdf.append([
            Paragraph(str(t[7]), left_style),
            Paragraph(str(t[3]), left_style),
            Paragraph(str(t[4]), left_style),
            Paragraph(type_tx, left_style),
            Paragraph(f"{montant:,.2f}", right_style),
            Paragraph(str(t[5]), left_style),
            Paragraph(str(t[6]), left_style),
        ])

    table = Table(data_pdf, colWidths=[55, 180, 65, 70, 70, 55, 55], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2196F3")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
        ('ALIGN', (4, 1), (4, -1), 'RIGHT'),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 5 * mm))

    solde = total_ent - total_sort
    elements.append(Paragraph(f"<b>Total Entrées :</b> <font color='green'>{total_ent:,.2f} F</font>", resume_style))
    elements.append(Paragraph(f"<b>Total Sorties :</b> <font color='red'>{total_sort:,.2f} F</font>", resume_style))
    couleur_solde = "green" if solde >= 0 else "red"
    elements.append(Paragraph(f"<b>Solde :</b> <font color='{couleur_solde}'>{solde:,.2f} F</font>", resume_style))

    # ===================== FOOTER =====================
    date_impression = datetime.datetime.today().strftime("%d-%m-%Y %H:%M")
    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.line(20 * mm, 18 * mm, A4[0] - 20 * mm, 18 * mm)
        canvas.drawString(20 * mm, 10 * mm, f"Date d'impression : {date_impression}")
        page_num = canvas.getPageNumber()
        canvas.drawRightString(A4[0] - 20 * mm, 10 * mm, f"Page {page_num}")
        canvas.restoreState()

    # Générer le PDF
    doc.build(elements, onFirstPage=add_page_number, onLaterPages=add_page_number)
    messagebox.showinfo("Succès", f"PDF généré : {file_path}")
# ----------------------------
# EXCEL
# ----------------------------

def exporter_excel(transactions):

    if not transactions:
        messagebox.showinfo("Excel", "Aucune transaction")
        return

    # =====================
    # DOSSIER RAPPORTS
    # =====================
    dossier_rapports = os.path.join(APP_FOLDER, "rapports")
    if not os.path.exists(dossier_rapports):
        os.makedirs(dossier_rapports)

    # Nom automatique du fichier
    date_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = os.path.join(dossier_rapports, f"Transactions_{date_str}.xlsx")

    # =====================
    # CREATION EXCEL
    # =====================
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"

    headers = ["Date","Description","Catégorie","Type","Montant","Compte","Mode"]
    ws.append(headers)

    # Alignement en-têtes
    for col in range(1, len(headers)+1):
        ws.cell(row=1, column=col).alignment = Alignment(horizontal='center')

    # Ajout des transactions
    for t in transactions:
        ws.append([t[7], t[3], t[4], t[1], t[2], t[5], t[6]])

    # Ajustement largeur colonnes
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length + 2

    # Sauvegarde
    wb.save(file_name)

    messagebox.showinfo("Succès", f"Export Excel généré : {file_name}")

    # Ouvrir automatiquement le fichier
    os.startfile(file_name)
    

# -----------------------
# INTERFACE PRINCIPALE
# -----------------------
root = tk.Tk()
root.title("Application Comptabilité INT MSC")

# Ajuster automatiquement à la taille de l'écran
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Définir une taille minimale, mais permettre le redimensionnement
root.minsize(800, 500)  # minimum pour petits écrans
root.geometry(f"{int(screen_width*0.9)}x{int(screen_height*0.9)}")  # 90% de l'écran

# ---- FONTS ----
font_label = ("Arial", 12)
font_entry = ("Arial", 12)
font_button = ("Arial", 12, "bold")

# Exemple de frame responsive
frame_top = tk.Frame(root)
frame_top.pack(fill="both", expand=True, padx=10, pady=10)



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

tk.Button(frame_saisie,text="Ajouter/Modifier", command=ajouter_transaction,bg="#4CAF50",fg="white", font=font_button, width=15).grid(row=2,column=3, padx=5, pady=10)
tk.Button(frame_saisie,text="Supprimer", command=demander_mot_de_passe,
          bg="#f44336",fg="white", font=font_button, width=15
).grid(row=2,column=4, padx=5, pady=10)
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


# --- BOUTONS ---
tk.Button(frame_filtre,text="Rechercher",command=rechercher_transactions,font=font_button,bg="#2196F3",fg="white",   width=15).grid(row=4,column=0,columnspan=2,pady=10)
tk.Button(frame_filtre,text="Tout afficher",command=tout_afficher,font=font_button,bg="#607D8B",fg="white",   width=15).grid(row=4,column=2,columnspan=2,pady=5)
tk.Button(frame_filtre,
          text="Imprimer sélection",
          command=imprimer_selection,
          bg="#2196F3",
          fg="white",
          font=("Arial",11,"bold"),
          width=20
).grid(row=4, column=4, columnspan=4, pady=10)

# --- CARD RÉSUMÉ ---
resume_frame = tk.LabelFrame(frame_top, text="Résumé", padx=10, pady=10, font=("Arial",12,"bold"))
resume_frame.pack(side="left", padx=10, pady=5, fill="y")

resume_label = tk.Label(resume_frame, text="", font=("Arial",11,"bold"), justify="left", anchor="w")
resume_label.pack(padx=5, pady=5, anchor="w")

def mise_a_jour_resume(transactions=None):
    if transactions is None:
        transactions = lire_transactions()
    
    total_entrees = sum(float(t[2]) for t in transactions if t[1]=="Entrée")
    total_sorties = sum(float(t[2]) for t in transactions if t[1]=="Sortie")
    solde = total_entrees - total_sorties
    
    # Couleurs et affichage professionnel avec indent
    resume_label.config(
        text=f"Entrées totales:\n    {total_entrees:,.2f} F\n\n"
             f"Sorties totales:\n    {total_sorties:,.2f} F\n\n"
             f"Solde:\n    {solde:,.2f} F",
        fg="black"  # Texte général
    )
    # Couleurs séparées avec tags
    resume_label.config(fg="black")  # par défaut noir
    
    # Si tu veux, on peut utiliser Label séparés pour chaque ligne et mettre les couleurs distinctes
    # Met à jour infos supplémentaires
    mise_a_jour_infos(transactions)

# --- INFORMATIONS SUPPLÉMENTAIRES ---
infos_frame = tk.LabelFrame(frame_top, text="Informations", padx=10, pady=10, font=("Arial",12,"bold"))
infos_frame.pack(side="left", padx=10, pady=5, fill="y")

infos_label = tk.Label(infos_frame, text="", font=("Arial",11), justify="left", anchor="w")
infos_label.pack(padx=5, pady=5, anchor="w")

def mise_a_jour_infos(transactions=None):
    if transactions is None:
        transactions = lire_transactions()
    
    nb_transactions = len(transactions)
    derniere_trans = transactions[-1] if transactions else None
    
    if derniere_trans:
        dernier_detail = (
            f"Date: {derniere_trans[0]}\n"
            f"Type: {derniere_trans[1]}\n"
            f"Montant: {derniere_trans[2]:,.2f} F"
        )
    else:
        dernier_detail = "Aucune"

    infos_label.config(
        text=f"Nombre total de transactions:\n    {nb_transactions}\n\n"
             f"Dernière transaction:\n    {dernier_detail}"
    )
    
# --- RAPPORTS À DROITE ---
frame_rapports = tk.LabelFrame(frame_top, text="Rapports", padx=10, pady=10, font=("Arial",12,"bold"))
frame_rapports.pack(side="right", padx=10)
tk.Button(frame_rapports,text="Journalier",command=lambda: afficher_rapport(filtrer_transactions("jour"), "Journalier"),font=font_button,bg="#2196F3",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Hebdomadaire",command=lambda: afficher_rapport(filtrer_transactions("semaine"), "Hebdomadaire"),font=font_button,bg="#2196F3",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Mensuel",command=lambda: afficher_rapport(filtrer_transactions("mois"), "Mensuel"),font=font_button,bg="#2196F3",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Annuel",command=lambda: afficher_rapport(filtrer_transactions("annee"), "Annuel"),font=font_button,bg="#2196F3",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Complet",command=lambda: afficher_rapport(lire_transactions(), "Complet"),font=font_button,bg="#4CAF50",fg="white").pack(padx=5,pady=5, fill="x")
tk.Button(frame_rapports,text="Export Excel",command=lambda: exporter_excel(lire_transactions()),font=font_button,bg="#FF9800",fg="white").pack(padx=5, pady=5, fill="x")
tk.Button(frame_rapports,text="Par Catégorie",command=lambda: afficher_rapport_par_categorie(lire_transactions()),font=font_button,bg="#9C27B0",fg="white").pack(padx=5, pady=5, fill="x")

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
