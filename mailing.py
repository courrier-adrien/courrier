"""
MAILING — Outil d'envoi de newsletter
Utilise la config SMTP de COURRIER (comptes.json)
"""
import sys, os, json, csv, time, smtplib, threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QSpinBox, QFileDialog, QTextEdit,
    QComboBox, QProgressBar, QFrame, QMessageBox, QInputDialog
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QFont, QColor, QPalette

DOSSIER_APP = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "COURRIER")
FICHIER_COMPTES = os.path.join(DOSSIER_APP, "comptes.json")
FICHIER_PROGRESSION = os.path.join(DOSSIER_APP, "mailing_progression.json")
FICHIER_CARNET = os.path.join(DOSSIER_APP, "carnet_adresses.json")

STYLE = """
    QMainWindow, QWidget { background-color: #2a2a2a; color: white; font-family: Verdana; font-size: 11px; }
    QPushButton {
        background-color: #A52A2A; color: white; border: none;
        padding: 6px 14px; font-family: Verdana; font-size: 11px; font-weight: bold;
        border-bottom: 3px solid #6b1a1a; border-right: 2px solid #6b1a1a;
    }
    QPushButton:hover { background-color: #c03030; }
    QPushButton:pressed { padding-top: 8px; border-bottom: 1px solid #6b1a1a; }
    QPushButton:disabled { background-color: #555; color: #888; border-bottom: 1px solid #333; }
    QLineEdit, QSpinBox, QComboBox {
        background-color: #1e1e1e; color: white; border: 1px solid #5a5a5a;
        padding: 4px 6px; font-family: Verdana; font-size: 11px;
    }
    QTextEdit { background-color: #1a1a1a; color: #cccccc; border: 1px solid #444; font-size: 10px; }
    QProgressBar {
        background-color: #1a1a1a; border: 1px solid #555; text-align: center; color: white;
    }
    QProgressBar::chunk { background-color: #A52A2A; }
    QLabel { color: white; }
    QComboBox QAbstractItemView { background-color: #1e1e1e; color: white; selection-background-color: #A52A2A; }
"""

def _dechiffrer_mdp(valeur):
    """Mot de passe chiffré par COURRIER (DPAPI Windows, voir
    chiffrer_secret dans courrier.py — même entropie obligatoire)."""
    import base64, ctypes

    class _BLOB(ctypes.Structure):
        _fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_char))]

    donnees = base64.b64decode(valeur[len("dpapi:"):])
    t1 = ctypes.create_string_buffer(donnees, len(donnees))
    t2 = ctypes.create_string_buffer(b"COURRIER-comptes-v1", 19)
    entree = _BLOB(len(donnees), ctypes.cast(t1, ctypes.POINTER(ctypes.c_char)))
    entropie = _BLOB(19, ctypes.cast(t2, ctypes.POINTER(ctypes.c_char)))
    sortie = _BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(entree), None, ctypes.byref(entropie),
                                                    None, None, 0x1, ctypes.byref(sortie)):
        raise OSError("DPAPI : déchiffrement impossible")
    try:
        return ctypes.string_at(sortie.pbData, sortie.cbData).decode("utf-8")
    finally:
        ctypes.windll.kernel32.LocalFree(sortie.pbData)

def charger_comptes():
    try:
        with open(FICHIER_COMPTES, encoding="utf-8") as f:
            comptes = json.load(f)
    except Exception:
        return []
    for c in comptes:
        if isinstance(c, dict) and c.get("mdp_chiffre"):
            try:
                c["mdp"] = _dechiffrer_mdp(c["mdp_chiffre"])
            except Exception as e:
                print(f"Mot de passe illisible pour {c.get('adresse', '?')} : {e}")
                c["mdp"] = ""
        if isinstance(c, dict):
            c.setdefault("mdp", "")
    return comptes

def charger_carnet():
    """Carnet d'adresses de COURRIER : {"dossiers": [...], "contacts": [{"nom", "email", "dossier"}, ...]}"""
    try:
        with open(FICHIER_CARNET, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"dossiers": [], "contacts": []}

# Titres de colonne reconnus pour l'adresse (mêmes que l'import du carnet
# dans COURRIER), en minuscules
COLONNES_EMAIL = ("email", "e-mail", "adresse courriel", "courriel",
                  "adresse e-mail", "adresse email", "mail")


def lire_classeur(chemin):
    """Onglets d'un classeur Excel : [(titre, [ligne (tuple de valeurs), ...])].
    .xls (Excel 2003 et avant) par xlrd, sinon openpyxl (08/10/2026)."""
    if chemin.lower().endswith(".xls"):
        import xlrd
        classeur = xlrd.open_workbook(chemin)
        try:
            return [(f.name, [tuple(f.row_values(r)) for r in range(f.nrows)])
                    for f in classeur.sheets()]
        finally:
            classeur.release_resources()
    import openpyxl
    classeur = openpyxl.load_workbook(chemin, data_only=True, read_only=True)
    try:
        return [(f.title, list(f.iter_rows(values_only=True))) for f in classeur.worksheets]
    finally:
        classeur.close()


def connexion_smtp(compte):
    if not compte.get("ssl_direct", False):
        serveur = smtplib.SMTP("127.0.0.1", 1587)
    elif compte.get("smtp_starttls", False) or compte.get("smtp_port") == 587:
        # Port 587 = toujours STARTTLS, même pour un compte saisi à la main.
        host = compte.get("smtp_host", "")
        port = compte.get("smtp_port", 587)
        serveur = smtplib.SMTP(host, port)
        serveur.starttls()
    else:
        host = compte.get("smtp_host", "")
        port = compte.get("smtp_port", 465)
        serveur = smtplib.SMTP_SSL(host, port)
    serveur.login(compte["adresse"], compte["mdp"])
    return serveur


class ThreadEnvoi(QThread):
    signal_log       = Signal(str)
    signal_progress  = Signal(int, int)   # envoyé, total
    signal_fini      = Signal(str)        # message final
    signal_countdown = Signal(int)        # secondes restantes

    def __init__(self, compte, destinataires, sujet, corps_html,
                 intervalle, index_debut=0):
        super().__init__()
        self.compte       = compte
        self.destinataires = destinataires
        self.sujet        = sujet
        self.corps_html   = corps_html
        self.intervalle   = intervalle
        self.index_debut  = index_debut
        self._stop        = False
        self._pause       = False

    def stop(self):   self._stop  = True
    def pause(self):  self._pause = True
    def reprendre(self): self._pause = False

    def run(self):
        total   = len(self.destinataires)
        envoyes = self.index_debut

        for i, email in enumerate(self.destinataires[self.index_debut:], self.index_debut):
            # Gestion pause
            while self._pause and not self._stop:
                time.sleep(0.5)
            if self._stop:
                self._sauvegarder_progression(i)
                self.signal_fini.emit(f"Arrêté à {i}/{total}. Reprise possible.")
                return

            # Connexion et envoi
            try:
                smtp = connexion_smtp(self.compte)
                from email.utils import formatdate
                msg = MIMEMultipart("alternative")
                msg["From"]    = f"{self.compte.get('nom', '')} <{self.compte['adresse']}>"
                msg["To"]      = email
                msg["Subject"] = Header(self.sujet, "utf-8")
                msg["Date"]    = formatdate(localtime=True)
                msg.attach(MIMEText(self.corps_html, "html", "utf-8"))
                smtp.sendmail(self.compte["adresse"], [email], msg.as_bytes())
                smtp.quit()
                envoyes += 1
                self.signal_progress.emit(envoyes, total)
                self.signal_log.emit(f"✓ {email}")
                self._sauvegarder_progression(i + 1)
            except Exception as e:
                self.signal_log.emit(f"✗ {email} — {e}")

            # Attendre avec décompte sauf si dernier
            if i < total - 1 and not self._stop:
                for sec in range(self.intervalle, 0, -1):
                    if self._stop:
                        break
                    while self._pause and not self._stop:
                        time.sleep(0.5)
                    self.signal_countdown.emit(sec)
                    time.sleep(1)
                self.signal_countdown.emit(0)

        self.signal_fini.emit(f"Terminé ! {envoyes}/{total} mails envoyés.")

    def _sauvegarder_progression(self, index):
        try:
            with open(FICHIER_PROGRESSION, "w", encoding="utf-8") as f:
                json.dump({
                    "index": index,
                    "total": len(self.destinataires),
                    "sujet": self.sujet,
                    "compte": self.compte.get("adresse", ""),
                }, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


class FenetreMailing(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MAILING")
        self.setMinimumSize(800, 700)
        self.setStyleSheet(STYLE)
        self.thread_envoi = None
        self.destinataires = []
        self.corps_html = ""
        self._init_ui()
        self._charger_comptes()
        self._verifier_progression()

        # Cacher la console
        try:
            import ctypes
            ctypes.windll.user32.ShowWindow(
                ctypes.windll.kernel32.GetConsoleWindow(), 0)
        except Exception:
            pass

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)

        # Titre
        titre = QLabel("MAILING")
        titre.setStyleSheet("color: #A52A2A; font-size: 16px; font-weight: bold; padding: 5px 0;")
        layout.addWidget(titre)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine); sep.setStyleSheet("color: #A52A2A;")
        layout.addWidget(sep)

        # Compte expéditeur
        h_compte = QHBoxLayout()
        h_compte.addWidget(QLabel("Compte expéditeur :"))
        self.combo_compte = QComboBox()
        self.combo_compte.setMinimumWidth(300)
        h_compte.addWidget(self.combo_compte)
        h_compte.addStretch(1)
        layout.addLayout(h_compte)

        # Sujet
        h_sujet = QHBoxLayout()
        h_sujet.addWidget(QLabel("Sujet :"))
        self.champ_sujet = QLineEdit()
        self.champ_sujet.setPlaceholderText("Objet de la newsletter...")
        h_sujet.addWidget(self.champ_sujet)
        layout.addLayout(h_sujet)

        # Fichiers
        h_csv = QHBoxLayout()
        self.lbl_csv = QLabel("Aucun fichier CSV chargé")
        self.lbl_csv.setStyleSheet("color: #888;")
        btn_csv = QPushButton("CHARGER CSV")
        btn_csv.clicked.connect(self._charger_csv)
        h_csv.addWidget(btn_csv)
        h_csv.addWidget(self.lbl_csv, 1)
        layout.addLayout(h_csv)

        # Destinataires pris dans un dossier du carnet d'adresses (07/10/2026)
        h_carnet = QHBoxLayout()
        self.lbl_carnet = QLabel("Aucun dossier d'adresses chargé")
        self.lbl_carnet.setStyleSheet("color: #888;")
        btn_carnet = QPushButton("CHARGER DOSSIER ADRESSES")
        btn_carnet.clicked.connect(self._charger_dossier_carnet)
        h_carnet.addWidget(btn_carnet)
        h_carnet.addWidget(self.lbl_carnet, 1)
        layout.addLayout(h_carnet)

        # Destinataires pris dans un classeur Excel, .xlsx ou .xls (08/10/2026)
        h_excel = QHBoxLayout()
        self.lbl_excel = QLabel("Aucun classeur Excel chargé")
        self.lbl_excel.setStyleSheet("color: #888;")
        btn_excel = QPushButton("CHARGER EXCEL")
        btn_excel.clicked.connect(self._charger_excel)
        h_excel.addWidget(btn_excel)
        h_excel.addWidget(self.lbl_excel, 1)
        layout.addLayout(h_excel)

        h_html = QHBoxLayout()
        self.lbl_html = QLabel("Aucun fichier HTML chargé")
        self.lbl_html.setStyleSheet("color: #888;")
        btn_html = QPushButton("CHARGER HTML")
        btn_html.clicked.connect(self._charger_html)
        h_html.addWidget(btn_html)
        h_html.addWidget(self.lbl_html, 1)
        layout.addLayout(h_html)

        # Intervalle
        h_interval = QHBoxLayout()
        h_interval.addWidget(QLabel("Intervalle entre envois :"))
        self.spin_intervalle = QSpinBox()
        self.spin_intervalle.setRange(10, 300)
        self.spin_intervalle.setValue(30)
        self.spin_intervalle.setSuffix(" secondes")
        self.spin_intervalle.setFixedWidth(130)
        h_interval.addWidget(self.spin_intervalle)
        h_interval.addStretch(1)
        layout.addLayout(h_interval)

        # Progression
        self.progress = QProgressBar()
        self.progress.setFixedHeight(22)
        self.progress.setFormat("%v / %m (%p%)")
        layout.addWidget(self.progress)

        self.lbl_statut = QLabel("Prêt.")
        self.lbl_statut.setStyleSheet("color: #aaaaaa; font-style: italic;")
        layout.addWidget(self.lbl_statut)

        self.lbl_countdown = QLabel("")
        self.lbl_countdown.setStyleSheet(
            "color: #A52A2A; font-size: 18px; font-weight: bold; qproperty-alignment: AlignCenter;")
        self.lbl_countdown.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_countdown.setFixedHeight(30)
        layout.addWidget(self.lbl_countdown)

        # Boutons d'action
        h_actions = QHBoxLayout()
        self.btn_lancer = QPushButton("LANCER L'ENVOI")
        self.btn_lancer.setStyleSheet(self.btn_lancer.styleSheet())
        self.btn_lancer.clicked.connect(self._lancer)

        self.btn_pause = QPushButton("PAUSE")
        self.btn_pause.setEnabled(False)
        self.btn_pause.clicked.connect(self._pause)

        self.btn_stop = QPushButton("ARRÊTER")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop)

        self.btn_reprendre = QPushButton("REPRENDRE")
        self.btn_reprendre.clicked.connect(self._reprendre_progression)
        self.btn_reprendre.setEnabled(False)

        h_actions.addWidget(self.btn_lancer)
        h_actions.addWidget(self.btn_pause)
        h_actions.addWidget(self.btn_stop)
        h_actions.addSpacing(20)
        h_actions.addWidget(self.btn_reprendre)
        h_actions.addStretch(1)
        layout.addLayout(h_actions)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.Shape.HLine); sep2.setStyleSheet("color: #444;")
        layout.addWidget(sep2)

        # Log
        log_lbl = QLabel("Journal d'envoi :")
        log_lbl.setStyleSheet("color: #888; font-size: 10px;")
        layout.addWidget(log_lbl)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(200)
        self.log.setStyleSheet("background-color: #1a1a1a; color: #cccccc; border: 1px solid #444; font-family: Verdana; font-size: 14px;")
        layout.addWidget(self.log)

        # Boutons CONSOLE / FERMER (style global de l'appli, pas d'override)
        h_bas = QHBoxLayout()
        h_bas.addStretch(1)
        btn_console = QPushButton("CONSOLE")
        btn_console.clicked.connect(self._basculer_console)
        h_bas.addWidget(btn_console)

        btn_fermer = QPushButton("FERMER")
        btn_fermer.clicked.connect(self.close)
        h_bas.addWidget(btn_fermer)
        layout.addLayout(h_bas)

    def _charger_comptes(self):
        comptes = charger_comptes()
        self.comptes = comptes
        for c in comptes:
            self.combo_compte.addItem(c.get("adresse", ""), c)

    def _charger_csv(self):
        chemin, _ = QFileDialog.getOpenFileName(
            self, "Charger la liste CSV", DOSSIER_APP, "CSV (*.csv *.txt)")
        if not chemin:
            return
        emails = []
        try:
            with open(chemin, encoding="utf-8", errors="ignore") as f:
                for ligne in f:
                    email = ligne.strip().strip(";,").strip()
                    if "@" in email and "." in email:
                        emails.append(email)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
            return
        self.destinataires = emails
        self.lbl_csv.setText(f"{len(emails)} adresses chargées — {os.path.basename(chemin)}")
        self.lbl_csv.setStyleSheet("color: #80c080;")
        self.progress.setMaximum(len(emails))
        self._reinitialiser_sources(sauf=self.lbl_csv)

    def _reinitialiser_sources(self, sauf):
        """Une seule source de destinataires à la fois (CSV, dossier ou
        Excel) : la dernière chargée remplace les autres, dont l'étiquette
        revient à "Aucun … chargé"."""
        for lbl, texte in ((self.lbl_csv, "Aucun fichier CSV chargé"),
                           (self.lbl_carnet, "Aucun dossier d'adresses chargé"),
                           (self.lbl_excel, "Aucun classeur Excel chargé")):
            if lbl is not sauf:
                lbl.setText(texte)
                lbl.setStyleSheet("color: #888;")

    def _charger_excel(self):
        """Adresses lues dans un classeur Excel (.xlsx, .xlsm ou .xls d'Excel
        2003) : colonne Email (ou E-mail, Courriel, Mail...) si l'onglet en
        a une, sinon toute cellule contenant une adresse. Avec plusieurs
        onglets, on choisit lequel (un export Excel du carnet de COURRIER a
        un onglet par dossier)."""
        chemin, _ = QFileDialog.getOpenFileName(
            self, "Charger un classeur Excel", DOSSIER_APP, "Classeur Excel (*.xlsx *.xlsm *.xls)")
        if not chemin:
            return
        try:
            feuilles = lire_classeur(chemin)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire le classeur :\n{e}")
            return
        utilisables = []
        for titre, lignes in feuilles:
            if not lignes:
                continue
            entete = [str(c or "").strip().lower() for c in lignes[0]]
            i_email = next((entete.index(n) for n in COLONNES_EMAIL if n in entete), None)
            emails, vus = [], set()
            for ligne in (lignes[1:] if i_email is not None else lignes):
                cellules = [ligne[i_email]] if i_email is not None and i_email < len(ligne) else ligne
                for v in cellules:
                    email = str(v or "").strip().strip(";,").strip()
                    if "@" in email and "." in email and " " not in email:
                        if email.lower() not in vus:
                            vus.add(email.lower())
                            emails.append(email)
                        break  # une adresse par ligne
            if emails:
                utilisables.append((titre, emails))
        if not utilisables:
            QMessageBox.warning(self, "Attention", "Aucune adresse email trouvée dans ce classeur.")
            return
        # Onglet récapitulatif d'un export du carnet : proposé en dernier
        utilisables.sort(key=lambda o: o[0].strip().upper() == "TOUS LES CONTACTS")
        if len(utilisables) == 1:
            titre, emails = utilisables[0]
        else:
            libelles = [f"{t}  ({len(e)})" for t, e in utilisables]
            choix, ok = QInputDialog.getItem(
                self, "Classeur Excel", "Envoyer le mailing aux adresses de l'onglet :",
                libelles, 0, False)
            if not ok:
                return
            titre, emails = utilisables[libelles.index(choix)]
        self.destinataires = emails
        self.lbl_excel.setText(f"{len(emails)} adresses chargées — {os.path.basename(chemin)}, onglet « {titre} »")
        self.lbl_excel.setStyleSheet("color: #80c080;")
        self.progress.setMaximum(len(emails))
        self._reinitialiser_sources(sauf=self.lbl_excel)

    def _charger_dossier_carnet(self):
        carnet = charger_carnet()
        contacts = carnet.get("contacts", [])
        nb_par_dossier = {}
        for c in contacts:
            d = c.get("dossier", "")
            nb_par_dossier[d] = nb_par_dossier.get(d, 0) + 1
        dossiers = sorted((d for d in carnet.get("dossiers", []) if nb_par_dossier.get(d)),
                          key=str.casefold)
        if not dossiers:
            QMessageBox.warning(self, "Attention",
                "Aucun dossier contenant des adresses dans le carnet.")
            return
        libelles = [f"{d}  ({nb_par_dossier[d]})" for d in dossiers]
        choix, ok = QInputDialog.getItem(
            self, "Dossier d'adresses", "Envoyer le mailing au dossier :",
            libelles, 0, False)
        if not ok:
            return
        dossier = dossiers[libelles.index(choix)]

        # Adresses valides du dossier, sans doublons (casse ignorée)
        emails, vus = [], set()
        for c in contacts:
            if c.get("dossier", "") != dossier:
                continue
            email = (c.get("email") or "").strip()
            if "@" in email and "." in email and email.lower() not in vus:
                vus.add(email.lower())
                emails.append(email)
        if not emails:
            QMessageBox.warning(self, "Attention",
                f"Le dossier « {dossier} » ne contient aucune adresse valide.")
            return
        self.destinataires = emails
        self.lbl_carnet.setText(f"{len(emails)} adresses chargées — dossier « {dossier} »")
        self.lbl_carnet.setStyleSheet("color: #80c080;")
        self.progress.setMaximum(len(emails))
        self._reinitialiser_sources(sauf=self.lbl_carnet)

    def _charger_html(self):
        chemin, _ = QFileDialog.getOpenFileName(
            self, "Charger le modèle HTML", DOSSIER_APP, "HTML (*.html *.htm)")
        if not chemin:
            return
        try:
            with open(chemin, encoding="utf-8") as f:
                self.corps_html = f.read()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
            return
        self.lbl_html.setText(f"✓ {os.path.basename(chemin)}")
        self.lbl_html.setStyleSheet("color: #80c080;")

    def _verifier_progression(self):
        if not os.path.exists(FICHIER_PROGRESSION):
            return
        try:
            with open(FICHIER_PROGRESSION, encoding="utf-8") as f:
                prog = json.load(f)
            index = prog.get("index", 0)
            total = prog.get("total", 0)
            sujet = prog.get("sujet", "")
            compte = prog.get("compte", "")
            if index > 0 and index < total:
                self.btn_reprendre.setEnabled(True)
                self.btn_reprendre.setToolTip(
                    f"Reprendre à {index}/{total}\nSujet: {sujet}\nCompte: {compte}")
                self.lbl_statut.setText(
                    f"⚠ Envoi interrompu à {index}/{total} — cliquez REPRENDRE")
                self.lbl_statut.setStyleSheet("color: #e08080;")
        except Exception:
            pass

    def _lancer(self):
        if not self.destinataires:
            QMessageBox.warning(self, "Attention",
                "Chargez d'abord un fichier CSV, un dossier d'adresses ou un classeur Excel.")
            return
        if not self.corps_html:
            QMessageBox.warning(self, "Attention", "Chargez d'abord un fichier HTML.")
            return
        sujet = self.champ_sujet.text().strip()
        if not sujet:
            QMessageBox.warning(self, "Attention", "Entrez le sujet du mail.")
            return
        compte = self.combo_compte.currentData()
        if not compte:
            QMessageBox.warning(self, "Attention", "Sélectionnez un compte expéditeur.")
            return
        self._demarrer_envoi(compte, self.destinataires, sujet, self.corps_html, 0)

    def _reprendre_progression(self):
        try:
            with open(FICHIER_PROGRESSION, encoding="utf-8") as f:
                prog = json.load(f)
        except Exception:
            return
        index = prog.get("index", 0)
        if not self.destinataires:
            QMessageBox.warning(self, "Attention",
                "Rechargez la même liste (CSV, dossier d'adresses ou classeur Excel) pour reprendre.")
            return
        if not self.corps_html:
            QMessageBox.warning(self, "Attention",
                "Rechargez le même fichier HTML pour reprendre.")
            return
        sujet = prog.get("sujet", self.champ_sujet.text().strip())
        compte = self.combo_compte.currentData()
        self.log.append(f"\n--- Reprise à partir de {index}/{len(self.destinataires)} ---")
        self._demarrer_envoi(compte, self.destinataires, sujet, self.corps_html, index)

    def _demarrer_envoi(self, compte, destinataires, sujet, corps_html, index_debut):
        intervalle = self.spin_intervalle.value()
        self.thread_envoi = ThreadEnvoi(
            compte, destinataires, sujet, corps_html, intervalle, index_debut)
        self.thread_envoi.signal_log.connect(self._on_log)
        self.thread_envoi.signal_progress.connect(self._on_progress)
        self.thread_envoi.signal_fini.connect(self._on_fini)
        self.thread_envoi.signal_countdown.connect(self._on_countdown)
        self.thread_envoi.start()
        self.btn_lancer.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_stop.setEnabled(True)
        self.btn_reprendre.setEnabled(False)
        self.lbl_statut.setText("Envoi en cours...")
        self.lbl_statut.setStyleSheet("color: #80c080;")

    def _pause(self):
        if not self.thread_envoi:
            return
        if self.thread_envoi._pause:
            self.thread_envoi.reprendre()
            self.btn_pause.setText("PAUSE")
            self.lbl_statut.setText("Envoi en cours...")
        else:
            self.thread_envoi.pause()
            self.btn_pause.setText("REPRENDRE")
            self.lbl_statut.setText("En pause.")

    def _stop(self):
        if self.thread_envoi:
            self.thread_envoi.stop()
        self.btn_pause.setEnabled(False)
        self.btn_stop.setEnabled(False)

    def _on_log(self, msg):
        self.log.append(msg)
        self.log.verticalScrollBar().setValue(
            self.log.verticalScrollBar().maximum())

    def _on_progress(self, envoyes, total):
        self.progress.setValue(envoyes)
        self.lbl_statut.setText(
            f"Envoi en cours : {envoyes}/{total} "
            f"— {envoyes/total*100:.1f}% "
            f"— reste ~{int((total-envoyes)*self.spin_intervalle.value()/60)} min")

    def _on_fini(self, msg):
        self.lbl_statut.setText(msg)
        self.lbl_statut.setStyleSheet("color: #90EE90;")
        self.btn_lancer.setEnabled(True)
        self.btn_pause.setEnabled(False)
        self.btn_stop.setEnabled(False)
        self.log.append(f"\n✅ {msg}")
        self._effacer_countdown()
        # Supprimer progression si terminé
        if "Terminé" in msg:
            try: os.remove(FICHIER_PROGRESSION)
            except Exception: pass

    def _on_countdown(self, sec):
        if sec > 0:
            self.lbl_countdown.setText(f"Prochain envoi dans {sec}s")
        else:
            self.lbl_countdown.setText("Envoi en cours...")

    def _effacer_countdown(self):
        self.lbl_countdown.setText("✅ Envoi terminé !")
        self.lbl_countdown.setStyleSheet(
            "color: #80c080; font-size: 18px; font-weight: bold; qproperty-alignment: AlignCenter;")

    def _basculer_console(self):
        try:
            import ctypes
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd == 0:
                # Pas de console (lancement via pythonw.exe) : on en crée une
                # à la demande et on y reconnecte stdout/stderr.
                if ctypes.windll.kernel32.AllocConsole():
                    try:
                        nouvelle_console = open("CONOUT$", "w", encoding="utf-8", buffering=1)
                        sys.stdout = nouvelle_console
                        sys.stderr = nouvelle_console
                    except Exception:
                        pass
            else:
                visible = ctypes.windll.user32.IsWindowVisible(hwnd)
                ctypes.windll.user32.ShowWindow(hwnd, 0 if visible else 5)
        except Exception:
            pass


def main():
    app = QApplication(sys.argv)
    palette = app.palette()
    palette.setColor(QPalette.ColorRole.Highlight, QColor("#A52A2A"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("white"))
    app.setPalette(palette)
    fenetre = FenetreMailing()
    fenetre.show()
    sys.exit(app.exec())
