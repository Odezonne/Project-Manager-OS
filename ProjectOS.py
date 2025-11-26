import os
import json
import shutil
import datetime
import threading
import urllib.parse
import urllib.request
import webbrowser
import platform 
import tkinter as tk
from tkinter import messagebox, scrolledtext
import customtkinter as ctk 

# Pour Outlook (Gère l'erreur si pas installé)
try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False

# --- CONFIGURATION ---
DOSSIER_RACINE_GENERAL = r"Q:\BETIC" 
DOSSIER_TEMPLATE = r"Q:\BETIC\2023\Dossier type nouveau projet - Copie"
FICHIER_FAVORIS = "favoris.json"
FICHIER_CONFIG = "config.json"

# 🔴 METTRE LE VRAI LIEN DU TIMESHEET ICI
URL_TIMESHEET = "http://bepromis.swecogroup.com/TimeSheet" 
URL_UPDATE_RAW = "https://raw.githubusercontent.com/Odezonne/Project-Manager-OS/refs/heads/main/ProjectOS_V4.py"

# --- INFOS ---
APP_VERSION = "4.0 (Galactic Overlord)"
DEV_MAIL = "olivier.schwickert@betic.lu"
DEV_NAME = "Olivier Schwickert"
DEV_TEAM = "Team2 (Les Vrais)"

# --- LEXIQUE ---
LEXIQUE_METIERS = {
    "🌍 Général (Tout)": [], 
    "⚡ Électricité": ["elec", "cfo", "cfa", "eclairage", "prise", "schema", "ele", "implantation", "cable", "strom", "elektro", "lighting", "power"],
    "❄️ HVAC / Ventil": ["hvac", "vent", "clim", "a -", "a_", "chauffage", "aera", "cvc", "heizung", "lüftung", "heating", "ventilation"],
    "💧 Sanitaire": ["san", "plomb", "evac", "b -", "b_", "eau", "sanitaire", "egout", "wasser", "abwasser", "plumbing"],
    "🏗️ Gros Oeuvre": ["beton", "archi", "coupe", "facade", "coffrage", "poutre", "dalle", "bau", "structure"],
    "🔥 Sécurité": ["feu", "incendie", "detec", "sprinkler", "fdi", "ssi", "evacuation", "fire"]
}

# --- THEMES "SOFT" (Yeux reposés) ---
THEMES_VISUELS = {
    "Graphite Soft (Sombre)":   {"sidebar": "#2b2b2b", "main": "#1f1f1f", "right": "#252525", "card": "#333333", "text": "#ecf0f1", "accent": "#3498db"},
    "Nordic Blue (Apaisant)":   {"sidebar": "#2E3440", "main": "#3B4252", "right": "#434C5E", "card": "#4C566A", "text": "#ECEFF4", "accent": "#88C0D0"},
    "Architecte (Clair)":       {"sidebar": "#F0F0F0", "main": "#FFFFFF", "right": "#F7F7F7", "card": "#E5E5E5", "text": "#2C3E50", "accent": "#95a5a6"},
    "Sage Green (Nature)":      {"sidebar": "#334D38", "main": "#E9F5EB", "right": "#DCEMD9", "card": "#C8E6C9", "text": "#1B3A25", "accent": "#66BB6A"},
}

BOUTONS_SIMPLE_CONFIG = {
    "🗂️ Administratif": "A -",
    "🏗️ Etude Chantier": "B -",
    "📷 Photos": "D -"
}

class AppGestionProjet(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title(f"Project OS - {APP_VERSION}")
        self.geometry("1450x900")
        self.user_name = os.getlogin().upper()
        
        # Variables
        self.path_projet_actuel = None
        self.numero_actuel = None
        self.current_theme_key = "Graphite Soft (Sombre)"
        self.current_metier_key = "🌍 Général (Tout)"

        # --- LAYOUT 3 COLONNES ---
        self.grid_columnconfigure(0, minsize=260) # Sidebar
        self.grid_columnconfigure(1, weight=1)    # Centre
        self.grid_columnconfigure(2, minsize=350) # Droite
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        self.main_area = ctk.CTkFrame(self, corner_radius=0)
        self.main_area.grid(row=0, column=1, sticky="nsew")
        
        self.right_panel = ctk.CTkFrame(self, corner_radius=0)
        self.right_panel.grid(row=0, column=2, sticky="nsew")
        
        self.setup_sidebar()
        self.setup_main_area()
        self.setup_right_panel_design()
        
        self.charger_config()
        self.charger_favoris_json()
        self.appliquer_theme_visuel()

    # --- CONFIG & STYLE ---
    def charger_config(self):
        if os.path.exists(FICHIER_CONFIG):
            try:
                with open(FICHIER_CONFIG, "r") as f:
                    data = json.load(f)
                    self.current_theme_key = data.get("custom_theme", "Graphite Soft (Sombre)")
                    self.current_metier_key = data.get("last_metier", "🌍 Général (Tout)")
                    self.combo_theme.set(self.current_theme_key)
                    if self.current_metier_key in LEXIQUE_METIERS:
                        self.combo_metier.set(self.current_metier_key)
            except: pass

    def sauver_config(self, event=None):
        try:
            with open(FICHIER_CONFIG, "w") as f: 
                json.dump({"custom_theme": self.current_theme_key, "last_metier": self.combo_metier.get()}, f)
        except: pass

    def appliquer_theme_visuel(self):
        cols = THEMES_VISUELS.get(self.current_theme_key, THEMES_VISUELS["Graphite Soft (Sombre)"])
        
        self.sidebar.configure(fg_color=cols["sidebar"])
        self.main_area.configure(fg_color=cols["main"])
        self.right_panel.configure(fg_color=cols["right"])
        self.fav_frame.configure(fg_color=cols["sidebar"])
        self.card_timesheet.configure(fg_color=cols["card"])
        self.card_outlook.configure(fg_color=cols["card"])
        
        # Gestion Light/Dark pour les widgets natifs
        if "Architecte" in self.current_theme_key or "Sage" in self.current_theme_key:
            ctk.set_appearance_mode("Light")
            self.listbox_favoris.config(bg="white", fg="black")
            self.list_mails.config(bg="#f5f5f5", fg="black")
        else:
            ctk.set_appearance_mode("Dark")
            self.listbox_favoris.config(bg="#2b2b2b", fg="white")
            self.list_mails.config(bg="#333", fg="white")

    def changer_theme_callback(self, choix):
        self.current_theme_key = choix
        self.appliquer_theme_visuel()
        self.sauver_config()

    # --- 1. SIDEBAR ---
    def setup_sidebar(self):
        ctk.CTkLabel(self.sidebar, text="PROJECT O.S.", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(30, 5))
        ctk.CTkLabel(self.sidebar, text=f"v{APP_VERSION}", font=ctk.CTkFont(size=10)).pack(pady=(0, 10))

        ctk.CTkLabel(self.sidebar, text=f"Bienvenue, {self.user_name}", font=ctk.CTkFont(weight="bold")).pack(pady=5)

        ctk.CTkLabel(self.sidebar, text="FILTRE MÉTIER", font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(15,0))
        self.combo_metier = ctk.CTkComboBox(self.sidebar, values=list(LEXIQUE_METIERS.keys()), width=220, command=self.sauver_config)
        self.combo_metier.pack(pady=5)

        ctk.CTkLabel(self.sidebar, text="STYLE VISUEL", font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(15,0))
        self.combo_theme = ctk.CTkComboBox(self.sidebar, values=list(THEMES_VISUELS.keys()), command=self.changer_theme_callback, width=220)
        self.combo_theme.pack(pady=5)

        ctk.CTkLabel(self.sidebar, text="ACCÈS RAPIDE", anchor="w", font=ctk.CTkFont(weight="bold")).pack(fill="x", padx=20, pady=(20, 0))
        self.fav_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.fav_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.listbox_favoris = tk.Listbox(self.fav_frame, bd=0, highlightthickness=0, font=("Segoe UI", 11))
        self.listbox_favoris.pack(fill="both", expand=True)
        self.listbox_favoris.bind('<Double-1>', self.charger_favori_depuis_liste)
        
        ctk.CTkButton(self.sidebar, text="Supprimer Favori", fg_color="#c0392b", height=25, command=self.supprimer_favori).pack(padx=20, pady=10)
        
        # Outils avec TITRES CLAIRS
        ctk.CTkLabel(self.sidebar, text="BOITE A OUTILS", anchor="w", font=ctk.CTkFont(size=10)).pack(fill="x", padx=20)
        
        # On remet des boutons normaux pour que ce soit lisible
        ctk.CTkButton(self.sidebar, text="💡 Idée de Génie", command=self.soumettre_idee, fg_color="#f1c40f", text_color="black", height=30).pack(fill="x", padx=20, pady=3)
        ctk.CTkButton(self.sidebar, text="🐞 Signaler Bug", command=self.signaler_bug, fg_color="transparent", border_width=1, height=30).pack(fill="x", padx=20, pady=3)
        ctk.CTkButton(self.sidebar, text="❓ À Propos", command=self.popup_about, fg_color="transparent", border_width=1, height=30).pack(fill="x", padx=20, pady=3)
        ctk.CTkButton(self.sidebar, text="🔄 Mise à jour", command=self.check_update, fg_color="transparent", border_width=1, height=30).pack(fill="x", padx=20, pady=3)

    # --- 2. MAIN AREA ---
    def setup_main_area(self):
        header = ctk.CTkFrame(self.main_area, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(20, 10))

        # Recherche
        self.entry_numero = ctk.CTkEntry(header, width=120, font=("Segoe UI", 16, "bold"), placeholder_text="25.XXX")
        self.entry_numero.pack(side="left", padx=(0, 10))
        self.entry_numero.bind('<Return>', self.rechercher_projet)
        ctk.CTkButton(header, text="OUVRIR", width=80, command=self.rechercher_projet).pack(side="left")
        ctk.CTkButton(header, text="+ Nouveau", width=80, fg_color="#2ecc71", hover_color="#27ae60", command=self.popup_creation_projet).pack(side="left", padx=10)

        # Scan
        f_scan = ctk.CTkFrame(header, fg_color="transparent")
        f_scan.pack(side="right")
        self.btn_scan = ctk.CTkButton(f_scan, text="🔎 SCANNER", width=120, height=35, fg_color="#8e44ad", font=("Segoe UI", 12, "bold"), command=lambda: self.lancer_recherche_v4(mode="metier"))
        self.btn_scan.pack(side="right", padx=10)
        self.entry_search_file = ctk.CTkEntry(f_scan, width=180, placeholder_text="Recherche globale...")
        self.entry_search_file.pack(side="right")
        self.entry_search_file.bind('<Return>', lambda e: self.lancer_recherche_v4(mode="libre"))

        self.lbl_titre = ctk.CTkLabel(self.main_area, text="En attente...", font=("Segoe UI", 28, "bold"), text_color="gray")
        self.lbl_titre.pack(anchor="w", padx=20, pady=(0, 10))

        # Zone Travail
        work_frame = ctk.CTkFrame(self.main_area, fg_color="transparent")
        work_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        top_work = ctk.CTkFrame(work_frame, fg_color="transparent")
        top_work.pack(fill="x", pady=(0, 20))
        top_work.grid_columnconfigure(0, weight=1)
        top_work.grid_columnconfigure(1, weight=1)

        # Infos
        f_info = ctk.CTkFrame(top_work)
        f_info.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ctk.CTkLabel(f_info, text="📝 INFOS", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        ctk.CTkLabel(f_info, text="Adresse :").pack(anchor="w", padx=10)
        self.var_adresse = tk.StringVar()
        ctk.CTkEntry(f_info, textvariable=self.var_adresse).pack(fill="x", padx=10, pady=2)
        ctk.CTkLabel(f_info, text="Phase :").pack(anchor="w", padx=10)
        # NOUVELLES PHASES ICI
        self.combo_phase = ctk.CTkComboBox(f_info, values=["APS", "APD", "PDE", "PDD", "SOUMISSION", "CHANTIER", "RECEPTION"])
        self.combo_phase.pack(fill="x", padx=10, pady=2)
        
        self.btn_save = ctk.CTkButton(f_info, text="Sauvegarder", state="disabled", height=25, command=self.sauvegarder_infos_json)
        self.btn_save.pack(fill="x", padx=10, pady=10)

        # Actions
        f_act = ctk.CTkFrame(top_work)
        f_act.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        ctk.CTkLabel(f_act, text="🚀 ACTIONS", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        self.btns_ui = []
        self.btns_ui.append(self.make_btn(f_act, "📂 Ouvrir Racine (Q:)", lambda: os.startfile(self.path_projet_actuel), "#555"))
        for nom, prefixe in BOUTONS_SIMPLE_CONFIG.items():
             self.btns_ui.append(self.make_btn(f_act, nom, lambda p=prefixe, n=nom: self.ouvrir_dossier_intelligent(p, n)))
        
        mini_row = ctk.CTkFrame(f_act, fg_color="transparent")
        mini_row.pack(fill="x", pady=2)
        self.btns_ui.append(self.make_btn(mini_row, "Plans", self.menu_plans_special, "#2980b9", "left"))
        self.btns_ui.append(self.make_btn(mini_row, "Mail", self.creer_email_projet, "#e67e22", "left"))
        self.btns_ui.append(self.make_btn(mini_row, "Note", self.popup_creer_note, "#8e44ad", "left"))
        self.btn_fav = self.make_btn(f_act, "⭐ Mettre en Favori", self.ajouter_favori, "#f1c40f", text_color="black")
        self.btns_ui.append(self.btn_fav)

        ctk.CTkLabel(work_frame, text="📌 PENSE-BÊTE DU PROJET", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 5))
        self.txt_pense_bete = scrolledtext.ScrolledText(work_frame, font=("Consolas", 11), bd=0, bg="#2b2b2b", fg="white", insertbackground="white")
        self.txt_pense_bete.pack(fill="both", expand=True)

    # --- 3. RIGHT PANEL ---
    def setup_right_panel_design(self):
        ctk.CTkLabel(self.right_panel, text="COMMUNICATION", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(30, 15))
        
        # TIMESHEET
        self.card_timesheet = ctk.CTkFrame(self.right_panel)
        self.card_timesheet.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkLabel(self.card_timesheet, text="⏱️ TIMESHEET", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=(10, 5))
        self.var_timesheet_done = ctk.BooleanVar(value=False)
        self.switch_time = ctk.CTkSwitch(self.card_timesheet, text="Pointage fait ?", variable=self.var_timesheet_done, onvalue=True, offvalue=False, progress_color="#2ecc71")
        self.switch_time.pack(padx=10, pady=5)
        ctk.CTkButton(self.card_timesheet, text="Ouvrir le Site", fg_color="#e67e22", hover_color="#d35400", command=self.ouvrir_timesheet).pack(fill="x", padx=10, pady=(5, 10))

        # TEAMS
        ctk.CTkButton(self.right_panel, text="💬 Ouvrir Teams", fg_color="#546e7a", height=35, command=self.ouvrir_teams).pack(fill="x", padx=15, pady=(0, 15))
        
        # OUTLOOK
        self.card_outlook = ctk.CTkFrame(self.right_panel)
        self.card_outlook.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        header_mail = ctk.CTkFrame(self.card_outlook, fg_color="transparent")
        header_mail.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(header_mail, text="📧 EMAILS", font=("Segoe UI", 12, "bold")).pack(side="left")
        ctk.CTkButton(header_mail, text="↻", width=30, height=20, fg_color="#555", command=self.charger_outlook_preview).pack(side="right")

        self.list_mails = tk.Listbox(self.card_outlook, bd=0, highlightthickness=0, font=("Segoe UI", 10), bg="#2b2b2b", fg="white", selectbackground="#444")
        self.list_mails.pack(fill="both", expand=True, padx=5, pady=(0, 10))
        self.list_mails.bind('<Double-1>', self.ouvrir_mail_selectionne)

        if OUTLOOK_AVAILABLE: self.after(1000, self.charger_outlook_preview)
        else: self.list_mails.insert(tk.END, "Outlook introuvable.")


    # --- LOGIQUE ---
    def ouvrir_timesheet(self): webbrowser.open(URL_TIMESHEET)
    def ouvrir_teams(self):
        try: os.startfile("msteams:")
        except: messagebox.showinfo("Info", "Impossible de lancer Teams.")
        
    def charger_outlook_preview(self):
        if not OUTLOOK_AVAILABLE: return
        self.list_mails.delete(0, tk.END)
        self.outlook_items = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)
            count = 0
            for msg in messages:
                try:
                    subject = msg.Subject
                    sender = msg.SenderName
                    display_text = f"✉️ {sender[:15]}.. | {subject[:25]}.."
                    self.list_mails.insert(tk.END, display_text)
                    self.outlook_items.append(msg)
                    count += 1
                    if count >= 20: break
                except: pass
        except Exception: self.list_mails.insert(tk.END, "Erreur lecture Outlook")

    def ouvrir_mail_selectionne(self, event):
        try:
            index = self.list_mails.curselection()[0]
            self.outlook_items[index].Display()
        except: pass

    # --- TEXTES FUN (MAIS PRO) ---
    def soumettre_idee(self):
        sujet = f"[GENIUS] Project OS v{APP_VERSION} - Idea"
        corps = (f"Salut {DEV_NAME},\n\nJ'ai une idée de génie :\n\n- TITRE :\n...\n\n- POURQUOI :\n...\n\nSi tu codes ça, je te paye un croissant.")
        try: os.startfile(f"mailto:{DEV_MAIL}?subject={urllib.parse.quote(sujet)}&body={urllib.parse.quote(corps)}")
        except: pass

    def signaler_bug(self):
        sys_info = f"{platform.system()} {platform.release()}"
        sujet = f"[DRAMA] Bug sur Project OS v{APP_VERSION}"
        corps = (f"Allo {DEV_NAME} ? On a un problème.\n\nMachine : {sys_info}\n\n- ACTION :\n...\n\n- RESULTAT :\n...\n\nRépare ça vite stp.")
        try: os.startfile(f"mailto:{DEV_MAIL}?subject={urllib.parse.quote(sujet)}&body={urllib.parse.quote(corps)}")
        except: pass

    def popup_about(self):
        msg = (f"PROJECT O.S. - v{APP_VERSION}\n\nDictateur du Code : {DEV_NAME}\n\"Je ne code pas des bugs, je code des fonctionnalités inattendues.\"\n\nSupport Moral : {DEV_TEAM}")
        messagebox.showinfo("À Propos", msg)

    # --- MOTEUR (Standard) ---
    def make_btn(self, parent, text, cmd, color=None, side="top", text_color="white"):
        btn = ctk.CTkButton(parent, text=text, command=cmd, state="disabled")
        if color: btn.configure(fg_color=color, hover_color=color) 
        if text_color == "black": btn.configure(text_color="black")
        if side=="top": btn.pack(fill="x", padx=10, pady=3)
        else: btn.pack(side="left", fill="x", expand=True, padx=2)
        return btn

    def rechercher_projet(self, event=None):
        if self.path_projet_actuel: self.sauvegarder_pense_bete()
        num = self.entry_numero.get().strip()
        self.path_projet_actuel = None
        if "." not in num: self.lbl_titre.configure(text="Format invalide"); return
        try:
            an = "20" + num.split('.')[0]
            path_an = os.path.join(DOSSIER_RACINE_GENERAL, an)
            if os.path.exists(path_an):
                for d in os.listdir(path_an):
                    if d.startswith(num):
                        self.path_projet_actuel = os.path.join(path_an, d)
                        self.numero_actuel = num
                        self.lbl_titre.configure(text=d, text_color="white")
                        for b in self.btns_ui: b.configure(state="normal")
                        self.btn_save.configure(state="normal")
                        self.btn_scan.configure(fg_color="#9b59b6")
                        self.charger_infos_json(); self.charger_pense_bete()
                        return
            self.lbl_titre.configure(text="Projet Introuvable (Q:)", text_color="red")
        except: pass

    def lancer_recherche_v4(self, event=None, mode="libre"):
        if not self.path_projet_actuel: return
        terme_libre = self.entry_search_file.get().strip().lower()
        metier_choisi = self.combo_metier.get()
        mots_cles_metier = LEXIQUE_METIERS.get(metier_choisi, [])
        if mode == "metier": titre = f"Documents Récents : {metier_choisi}"
        else:
            if not terme_libre: return
            titre = f"Recherche : '{terme_libre}'"
        self.top_res = ctk.CTkToplevel(self)
        self.top_res.geometry("1100x750")
        self.top_res.title(titre)
        self.top_res.attributes("-topmost", True)
        filter_frame = ctk.CTkFrame(self.top_res, fg_color="transparent")
        filter_frame.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(filter_frame, text="FILTRER :", font=("Segoe UI", 12, "bold")).pack(side="left", padx=(10, 20))
        def filter_click(ftype): self.display_sorted_results(ftype)
        ctk.CTkButton(filter_frame, text="TOUT", width=80, fg_color="#555", command=lambda: filter_click("ALL")).pack(side="left", padx=5)
        ctk.CTkButton(filter_frame, text="📕 PDF", width=80, fg_color="#e74c3c", command=lambda: filter_click("PDF")).pack(side="left", padx=5)
        ctk.CTkButton(filter_frame, text="📐 PLANS", width=80, fg_color="#3498db", command=lambda: filter_click("PLANS")).pack(side="left", padx=5)
        ctk.CTkButton(filter_frame, text="📊 OFFICE", width=80, fg_color="#27ae60", command=lambda: filter_click("OFFICE")).pack(side="left", padx=5)
        self.scroll_results = ctk.CTkScrollableFrame(self.top_res)
        self.scroll_results.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.lbl_loading = ctk.CTkLabel(self.scroll_results, text="🚀 Analyse en cours...", font=("Segoe UI", 14))
        self.lbl_loading.pack(pady=20)
        threading.Thread(target=lambda: self.execute_smart_scan(terme_libre, mots_cles_metier, mode), daemon=True).start()

    def execute_smart_scan(self, terme_libre, mots_cles_metier, mode):
        self.found_files_cache = [] 
        mots_recherche = terme_libre.split(" ") if mode == "libre" else []
        for r, d, f in os.walk(self.path_projet_actuel):
            if "RECYCLE" in r or "Trash" in r: continue 
            for file in f:
                file_lower = file.lower()
                path_lower = r.lower()
                full_path = os.path.join(r, file)
                match = False
                if mode == "libre":
                    if all(mot in file_lower or mot in path_lower for mot in mots_recherche):
                        match = True
                        if mots_cles_metier:
                             if not any(k in file_lower or k in path_lower for k in mots_cles_metier): match = False
                elif mode == "metier":
                    if not mots_cles_metier: match = True
                    else:
                        if any(k in file_lower for k in mots_cles_metier) or any(k in path_lower for k in mots_cles_metier): match = True
                if match:
                    try:
                        mtime = os.path.getmtime(full_path)
                        ftype = self.get_file_type(file)
                        self.found_files_cache.append({"name": file, "path": r, "full": full_path, "type": ftype, "time": mtime, "date_str": datetime.datetime.fromtimestamp(mtime).strftime('%d/%m/%Y %H:%M')})
                    except: pass
        self.found_files_cache.sort(key=lambda x: x['time'], reverse=True)
        self.top_res.after(0, lambda: self.display_sorted_results("ALL"))

    def display_sorted_results(self, filter_type):
        for widget in self.scroll_results.winfo_children(): widget.destroy()
        if filter_type == "ALL": files_to_show = self.found_files_cache
        else: files_to_show = [f for f in self.found_files_cache if f["type"] == filter_type]
        files_to_show = files_to_show[:100]
        if not files_to_show:
            ctk.CTkLabel(self.scroll_results, text="Aucun résultat.").pack(pady=20)
            return
        ctk.CTkLabel(self.scroll_results, text=f"Top {len(files_to_show)} Résultats récents", text_color="gray").pack(pady=(0, 10))
        for item in files_to_show: self.draw_smart_card(item)

    def draw_smart_card(self, item):
        icon = self.get_file_icon(item["type"])
        card = ctk.CTkFrame(self.scroll_results, fg_color=("gray95", "#333"))
        card.pack(fill="x", pady=3, padx=5)
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(card, text=icon, font=("Segoe UI", 26)).grid(row=0, column=0, rowspan=2, padx=15, pady=5)
        title_frame = ctk.CTkFrame(card, fg_color="transparent")
        title_frame.grid(row=0, column=1, sticky="w", pady=(5, 0))
        ctk.CTkLabel(title_frame, text=item["name"], font=("Segoe UI", 13, "bold"), anchor="w").pack(side="left")
        ctk.CTkLabel(title_frame, text=f"  ({item['date_str']})", font=("Segoe UI", 11), text_color="orange").pack(side="left")
        try: rel = os.path.relpath(item["path"], self.path_projet_actuel)
        except: rel = item["path"]
        ctk.CTkLabel(card, text=f"📂 {rel}", font=("Segoe UI", 11), text_color="gray", anchor="w").grid(row=1, column=1, sticky="w", pady=(0, 5))
        ctk.CTkButton(card, text="Ouvrir", width=70, height=28, command=lambda f=item["full"]: os.startfile(f)).grid(row=0, column=2, rowspan=2, padx=15)

    def get_file_type(self, filename):
        ext = filename.lower().split('.')[-1]
        if ext == "pdf": return "PDF"
        if ext in ["dwg", "dxf", "rvt", "pln", "ifc"]: return "PLANS"
        if ext in ["xlsx", "xls", "csv", "doc", "docx", "txt", "ppt"]: return "OFFICE"
        if ext in ["jpg", "png", "jpeg", "bmp"]: return "IMAGES"
        if ext in ["msg", "eml"]: return "EMAILS"
        return "AUTRE"
    def get_file_icon(self, ftype):
        if ftype == "PDF": return "📕"
        if ftype == "PLANS": return "📐"
        if ftype == "OFFICE": return "📊"
        if ftype == "IMAGES": return "🖼️"
        if ftype == "EMAILS": return "📧"
        return "📄"
    
    def check_update(self): 
        try:
            print(f"Checking {URL_UPDATE_RAW}")
            response = urllib.request.urlopen(URL_UPDATE_RAW)
            remote_code = response.read().decode('utf-8')
            remote_v = "Inconnue"
            for line in remote_code.splitlines():
                if line.startswith('APP_VERSION = "'):
                    remote_v = line.split('"')[1]
                    break
            if remote_v != APP_VERSION:
                if messagebox.askyesno("Update", f"Nouvelle version : {remote_v}\nInstaller ?"):
                    name = os.path.basename(__file__)
                    shutil.copy(name, name+".bak")
                    with open(name, "w", encoding="utf-8") as f: f.write(remote_code)
                    messagebox.showinfo("Reboot", "Mise à jour OK. Redémarrez.")
                    self.destroy()
            else: messagebox.showinfo("Update", "Vous êtes à jour.")
        except: messagebox.showerror("Erreur", "Impossible de vérifier.")

    def charger_infos_json(self):
        try:
            with open(os.path.join(self.path_projet_actuel, "_data.json")) as f: d=json.load(f); self.var_adresse.set(d.get("adresse","")); self.combo_phase.set(d.get("phase",""))
        except: pass
    def sauvegarder_infos_json(self):
        f = os.path.join(self.path_projet_actuel, "_data.json")
        with open(f, "w") as fi: json.dump({"adresse": self.var_adresse.get(), "phase": self.combo_phase.get()}, fi)
        self.sauvegarder_pense_bete()
    def charger_pense_bete(self):
        self.txt_pense_bete.delete("1.0", tk.END)
        p = os.path.join(self.path_projet_actuel, "_pense_bete.txt")
        if os.path.exists(p): 
            with open(p, "r", encoding="utf-8") as f: self.txt_pense_bete.insert("1.0", f.read())
    def sauvegarder_pense_bete(self):
        if self.path_projet_actuel:
             with open(os.path.join(self.path_projet_actuel, "_pense_bete.txt"), "w", encoding="utf-8") as f: f.write(self.txt_pense_bete.get("1.0", tk.END))
    def ouvrir_dossier_intelligent(self, p, n):
        for i in os.listdir(self.path_projet_actuel):
            if i.startswith(p): os.startfile(os.path.join(self.path_projet_actuel, i)); return
        d = os.path.join(self.path_projet_actuel, f"{p} {n.split(' ')[-1]}"); os.makedirs(d); os.startfile(d)
    def menu_plans_special(self): self.ouvrir_dossier_intelligent("C -", "Plans")
    def creer_email_projet(self):
        if self.nom_dossier_complet: os.startfile(f"mailto:?subject={urllib.parse.quote(self.nom_dossier_complet)}")
    def charger_favoris_json(self):
        self.listbox_favoris.delete(0, tk.END)
        if os.path.exists(FICHIER_FAVORIS):
            try: 
                with open(FICHIER_FAVORIS) as f: 
                    for i in json.load(f): self.listbox_favoris.insert(tk.END, i)
            except: pass
    def ajouter_favori(self):
        if self.numero_actuel and self.numero_actuel not in self.listbox_favoris.get(0, tk.END):
            self.listbox_favoris.insert(tk.END, self.numero_actuel)
            with open(FICHIER_FAVORIS, "w") as f: json.dump(list(self.listbox_favoris.get(0, tk.END)), f)
    def supprimer_favori(self):
        s = self.listbox_favoris.curselection()
        if s: self.listbox_favoris.delete(s); self.ajouter_favori()
        with open(FICHIER_FAVORIS, "w") as f: json.dump(list(self.listbox_favoris.get(0, tk.END)), f)
    def charger_favori_depuis_liste(self, e):
        s = self.listbox_favoris.curselection()
        if s: self.entry_numero.delete(0, tk.END); self.entry_numero.insert(0, self.listbox_favoris.get(s)); self.rechercher_projet()
    def popup_creation_projet(self):
        t = ctk.CTkToplevel(self); t.geometry("300x250"); t.title("Nouveau"); t.attributes("-topmost", True)
        ctk.CTkLabel(t, text="Numéro:").pack(); e1 = ctk.CTkEntry(t); e1.pack()
        ctk.CTkLabel(t, text="Nom:").pack(); e2 = ctk.CTkEntry(t); e2.pack()
        def ok():
            num, nom = e1.get(), e2.get()
            an = "20" + num.split('.')[0]
            p = os.path.join(DOSSIER_RACINE_GENERAL, an, f"{num} - {nom}")
            try:
                os.makedirs(p, exist_ok=True)
                if os.path.exists(DOSSIER_TEMPLATE): 
                    try: shutil.copytree(DOSSIER_TEMPLATE, p, dirs_exist_ok=True)
                    except: pass
                with open(os.path.join(p, "_data.json"), "w") as f: json.dump({"adresse": "", "phase": "ESQ"}, f)
                t.destroy(); self.entry_numero.delete(0, tk.END); self.entry_numero.insert(0, num); self.rechercher_projet()
            except Exception as e: messagebox.showerror("Erreur Création", str(e))
        ctk.CTkButton(t, text="GO", command=ok).pack(pady=20)
    def popup_creer_note(self):
        t=ctk.CTkToplevel(self); t.geometry("300x150"); t.title("Note"); t.attributes("-topmost", True)
        ctk.CTkLabel(t, text="Sujet:").pack(); e=ctk.CTkEntry(t); e.pack()
        def ok():
            p=os.path.join(self.path_projet_actuel, "B - Etude", "X - Notes"); os.makedirs(p, exist_ok=True)
            f=os.path.join(p, f"{datetime.date.today()}_{e.get()}.txt")
            with open(f,"w") as fi: fi.write(f"PROJET: {self.nom_dossier_complet}\nSUJET: {e.get()}")
            os.startfile(f); t.destroy()
        ctk.CTkButton(t, text="OK", command=ok).pack(pady=10)

if __name__ == "__main__":
    app = AppGestionProjet()
    app.mainloop()