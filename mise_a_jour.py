# -*- coding: utf-8 -*-
"""Mise a jour automatique de la version diffusee de COURRIER (exe PyInstaller).

Au demarrage, verifie en arriere-plan la derniere release GitHub. Si le zip publie
est plus recent que cette build, propose de l'installer :
  1) telechargement du zip et decompression dans DOSSIER_DONNEES\\mise_a_jour ;
  2) fermeture de COURRIER (sa fermeture normale : rien n'est perdu) ;
  3) un petit script PowerShell attend la fin du processus, remplace l'exe et le
     dossier _internal (en gardant l'ancienne version jusqu'au bout, restauree en
     cas d'echec), puis relance COURRIER.
Les donnees de l'utilisateur (%LOCALAPPDATA%\\COURRIER) ne sont jamais touchees :
seul le dossier du programme est remplace.

Fichier copie dans src\\ par patch.py ; inactif hors exe (sys.frozen)."""
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

from PySide6.QtCore import QObject, QThread, QTimer, QUrl, Qt, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QApplication, QMessageBox, QProgressDialog

DEPOT_GITHUB = "courrier-adrien/courrier"
URL_API = f"https://api.github.com/repos/{DEPOT_GITHUB}/releases/latest"
URL_PAGE = f"https://github.com/{DEPOT_GITHUB}/releases/latest"
DELAI_VERIFICATION_MS = 20000  # laisse COURRIER finir de demarrer avant d'aller sur Internet
MOTIF_VERSION = re.compile(r"(\d{4}-\d{2}-\d{2})[ ._](\d{2})h(\d{2})")


def version_depuis_nom(nom):
    """'(nom du zip).2026-10-13.01h11.zip' -> '2026-10-13 01h11' (None si absent).
    Ce format se compare directement comme une chaine."""
    m = MOTIF_VERSION.search(nom or "")
    return f"{m.group(1)} {m.group(2)}h{m.group(3)}" if m else None


def version_lisible(version):
    """'2026-10-13 01h11' -> '13/10/2026 à 01h11'."""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2}) (\d{2}h\d{2})", version or "")
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)} à {m.group(4)}" if m else (version or "?")


class _Verificateur(QThread):
    resultat = Signal(object)  # dict(version, nom, url, taille) ou None

    def run(self):
        try:
            import requests
            rep = requests.get(URL_API, timeout=20,
                               headers={"Accept": "application/vnd.github+json"})
            rep.raise_for_status()
            meilleur = None
            for a in rep.json().get("assets", []):
                version = version_depuis_nom(a.get("name"))
                if (version and a.get("name", "").lower().endswith(".zip")
                        and a.get("state", "uploaded") == "uploaded"
                        and (meilleur is None or version > meilleur["version"])):
                    meilleur = {"version": version, "nom": a["name"],
                                "url": a["browser_download_url"], "taille": a.get("size") or 0}
            self.resultat.emit(meilleur)
        except Exception as e:
            print(f"[MAJ] verification impossible : {e}")
            self.resultat.emit(None)


class _Telechargeur(QThread):
    progression = Signal(int, int)  # Mo recus, Mo total
    etape = Signal(str)
    fini = Signal(str)  # dossier de la nouvelle version decompressee
    erreur = Signal(str)

    def __init__(self, info, dossier_travail, nom_exe):
        super().__init__()
        self.info = info
        self.dossier_travail = dossier_travail
        self.nom_exe = nom_exe
        self.annule = False

    def run(self):
        try:
            import requests
            shutil.rmtree(self.dossier_travail, ignore_errors=True)
            os.makedirs(self.dossier_travail, exist_ok=True)
            chemin_zip = os.path.join(self.dossier_travail, "nouvelle_version.zip")
            total = self.info["taille"]
            recu = 0
            with requests.get(self.info["url"], stream=True, timeout=60) as rep:
                rep.raise_for_status()
                total = int(rep.headers.get("Content-Length") or total or 0)
                with open(chemin_zip, "wb") as f:
                    for bloc in rep.iter_content(1024 * 1024):
                        if self.annule:
                            return
                        f.write(bloc)
                        recu += len(bloc)
                        self.progression.emit(recu // (1024 * 1024), total // (1024 * 1024))
            if total and recu != total:
                raise IOError(f"telechargement incomplet ({recu} octets sur {total})")

            self.etape.emit("Décompression de la mise à jour...")
            dossier_extrait = os.path.join(self.dossier_travail, "extrait")
            with zipfile.ZipFile(chemin_zip) as z:
                for membre in z.infolist():
                    if self.annule:
                        return
                    z.extract(membre, dossier_extrait)
            os.remove(chemin_zip)

            # Le zip contient un seul dossier racine ("COURRIER ADRIEN\...") : on cherche
            # celui qui contient l'exe et _internal, sans dependre de son nom.
            for racine, dossiers, fichiers in os.walk(dossier_extrait):
                if self.nom_exe in fichiers and "_internal" in dossiers:
                    self.fini.emit(racine)
                    return
            raise IOError(f"{self.nom_exe} introuvable dans le zip telecharge")
        except Exception as e:
            if not self.annule:
                self.erreur.emit(str(e))


# Script de remplacement : tourne apres la fermeture de COURRIER. Ecrit son resultat
# dans un fichier lu au lancement suivant (message "mis a jour" ou "echec").
SCRIPT_POWERSHELL = r'''
param([int]$IdProcessus, [string]$Installation, [string]$Nouvelle, [string]$NomExe,
      [string]$Resultat, [string]$Version, [int]$Relancer)
$ErrorActionPreference = "Stop"
function Ecrire-Resultat($texte) { Set-Content -LiteralPath $Resultat -Value $texte -Encoding UTF8 }

# 1) Attendre la fin de COURRIER et de tous les processus lances depuis son dossier
#    (fenetre MAILING, moteur web QtWebEngineProcess...). Forcer apres 90 s : la
#    fermeture normale a deja tout enregistre a ce stade.
try { Wait-Process -Id $IdProcessus -Timeout 90 -ErrorAction SilentlyContinue } catch {}
$prefixe = ($Installation.TrimEnd('\') + '\').ToLower()
for ($i = 0; $i -lt 60; $i++) {
    $restants = @(Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -and $_.Path.ToLower().StartsWith($prefixe) })
    if ($restants.Count -eq 0) { break }
    if ($i -eq 45) { $restants | Stop-Process -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 1
}

# 2) Mettre l'ancienne version de cote (renommage), puis copier la nouvelle.
#    Plusieurs essais : un antivirus peut verrouiller brievement un fichier.
$exe = Join-Path $Installation $NomExe
$interne = Join-Path $Installation "_internal"
$exeAncien = "$exe.ancien"
$interneAncien = "$interne.ancien"
Remove-Item -LiteralPath $exeAncien, $interneAncien -Recurse -Force -ErrorAction SilentlyContinue
$deCote = $false
for ($i = 0; $i -lt 20 -and -not $deCote; $i++) {
    try {
        Rename-Item -LiteralPath $interne -NewName "_internal.ancien"
        try { Rename-Item -LiteralPath $exe -NewName "$NomExe.ancien" }
        catch { Rename-Item -LiteralPath $interneAncien -NewName "_internal"; throw }
        $deCote = $true
    } catch { Start-Sleep -Seconds 2 }
}
if (-not $deCote) {
    Ecrire-Resultat "ECHEC`nLes fichiers de COURRIER sont restés verrouillés (antivirus ?). Rien n'a été modifié."
} else {
    try {
        Copy-Item -LiteralPath (Join-Path $Nouvelle "_internal") -Destination $interne -Recurse -Force
        Copy-Item -LiteralPath (Join-Path $Nouvelle $NomExe) -Destination $exe -Force
        Get-ChildItem -LiteralPath $Nouvelle -File | Where-Object { $_.Name -ne $NomExe } |
            Copy-Item -Destination $Installation -Force
        Remove-Item -LiteralPath $exeAncien, $interneAncien -Recurse -Force -ErrorAction SilentlyContinue
        Ecrire-Resultat "OK`n$Version"
    } catch {
        # Retour a l'ancienne version, intacte
        $erreur = $_.Exception.Message
        Remove-Item -LiteralPath $exe, $interne -Recurse -Force -ErrorAction SilentlyContinue
        Rename-Item -LiteralPath $interneAncien -NewName "_internal" -ErrorAction SilentlyContinue
        Rename-Item -LiteralPath $exeAncien -NewName $NomExe -ErrorAction SilentlyContinue
        Ecrire-Resultat "ECHEC`nLa copie de la nouvelle version a échoué ($erreur). L'ancienne version a été remise en place."
    }
}
if ($Relancer -eq 1) { Start-Process -FilePath $exe -WorkingDirectory $Installation }
'''


class GestionnaireMiseAJour(QObject):
    def __init__(self, fenetre, version_actuelle, dossier_donnees, style_boutons=""):
        super().__init__(fenetre)
        # Meme apparence que les autres boites de message de COURRIER (boutons rouges)
        self.style_boutons = style_boutons
        self.style_boites = ("QMessageBox { background-color: #2a2a2a; } QLabel { color: white; }"
                             + style_boutons)
        self.fenetre = fenetre
        self.version_actuelle = version_actuelle
        self.dossier_donnees = dossier_donnees
        self.dossier_travail = os.path.join(dossier_donnees, "mise_a_jour")
        self.fichier_prefs = os.path.join(dossier_donnees, "mise_a_jour.json")
        self.fichier_resultat = os.path.join(dossier_donnees, "mise_a_jour_resultat.txt")
        self.installation = os.path.dirname(os.path.abspath(sys.executable))
        self.nom_exe = os.path.basename(sys.executable)
        self.info = None
        self.nouvelle = None  # dossier de la version telechargee, prete a installer
        self.relancer = True
        self._verificateur = None
        self._telechargeur = None
        self._progression = None

    # ---------- demarrage ----------
    def demarrer(self):
        QApplication.instance().aboutToQuit.connect(self._arreter_threads)
        QTimer.singleShot(3000, self._annoncer_resultat_precedent)
        QTimer.singleShot(DELAI_VERIFICATION_MS, self.verifier)

    def _arreter_threads(self):
        """A la fermeture : un QThread detruit en cours d'execution fait planter Qt."""
        if self._telechargeur is not None and self._telechargeur.isRunning():
            self._telechargeur.annule = True
            self._telechargeur.wait(5000)
        if self._verificateur is not None and self._verificateur.isRunning():
            self._verificateur.wait(3000)

    def _boite(self, icone, titre, texte):
        boite = QMessageBox(self.fenetre)
        boite.setIcon(icone)
        boite.setWindowTitle(titre)
        boite.setText(texte)
        boite.setStyleSheet(self.style_boites)
        return boite

    def _annoncer_resultat_precedent(self):
        if not os.path.exists(self.fichier_resultat):
            # Restes d'une mise a jour abandonnee (telechargement interrompu, etc.)
            shutil.rmtree(self.dossier_travail, ignore_errors=True)
            return
        try:
            with open(self.fichier_resultat, encoding="utf-8-sig") as f:
                lignes = f.read().strip().splitlines()
            os.remove(self.fichier_resultat)
        except OSError:
            return
        shutil.rmtree(self.dossier_travail, ignore_errors=True)
        if lignes and lignes[0] == "OK":
            self._boite(QMessageBox.Icon.Information, "Mise à jour",
                        f"COURRIER a été mis à jour (version du "
                        f"{version_lisible(self.version_actuelle)}).\n\n"
                        f"Vos comptes, mails et réglages sont conservés.").exec()
        elif lignes:
            self._boite(QMessageBox.Icon.Warning, "Mise à jour",
                        "La mise à jour de COURRIER n'a pas pu être installée.\n\n"
                        + "\n".join(lignes[1:])
                        + f"\n\nVous pouvez la télécharger à la main :\n{URL_PAGE}").exec()

    def _prefs(self):
        try:
            with open(self.fichier_prefs, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _sauver_prefs(self, prefs):
        try:
            with open(self.fichier_prefs, "w", encoding="utf-8") as f:
                json.dump(prefs, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

    # ---------- verification ----------
    def verifier(self):
        if self._verificateur is not None or self.nouvelle is not None:
            return
        self._verificateur = _Verificateur()
        self._verificateur.resultat.connect(self._sur_resultat_verification)
        self._verificateur.finished.connect(self._fin_verificateur)
        self._verificateur.start()

    def _fin_verificateur(self):
        self._verificateur = None

    def _sur_resultat_verification(self, info):
        if not info or info["version"] <= self.version_actuelle:
            return
        if self._prefs().get("version_ignoree") == info["version"]:
            return
        self.info = info
        boite = self._boite(QMessageBox.Icon.Information, "Mise à jour disponible",
                            f"Une nouvelle version de COURRIER est disponible "
                            f"(version du {version_lisible(info['version'])}).")
        boite.setInformativeText(
            f"Votre version : {version_lisible(self.version_actuelle)}\n"
            f"Téléchargement : environ {max(1, info['taille'] // (1024 * 1024))} Mo\n\n"
            f"Vos comptes, mails et réglages sont conservés.\n"
            f"COURRIER se fermera puis redémarrera tout seul.")
        b_installer = boite.addButton("Installer maintenant", QMessageBox.ButtonRole.AcceptRole)
        boite.addButton("Plus tard", QMessageBox.ButtonRole.RejectRole)
        b_ignorer = boite.addButton("Ignorer cette version", QMessageBox.ButtonRole.DestructiveRole)
        boite.setDefaultButton(b_installer)
        boite.exec()
        if boite.clickedButton() is b_installer:
            self.telecharger()
        elif boite.clickedButton() is b_ignorer:
            prefs = self._prefs()
            prefs["version_ignoree"] = info["version"]
            self._sauver_prefs(prefs)

    # ---------- telechargement ----------
    def _installation_modifiable(self):
        essai = os.path.join(self.installation, ".essai_ecriture_maj")
        try:
            with open(essai, "w") as f:
                f.write("ok")
            os.remove(essai)
            return True
        except OSError:
            return False

    def telecharger(self):
        if not self._installation_modifiable():
            self._boite(
                QMessageBox.Icon.Warning, "Mise à jour",
                f"COURRIER est installé dans un dossier protégé par Windows :\n"
                f"{self.installation}\n\n"
                f"La mise à jour automatique ne peut pas le modifier. Téléchargez la "
                f"nouvelle version à la main (la page va s'ouvrir), ou déplacez le "
                f"dossier de COURRIER dans Documents par exemple.").exec()
            QDesktopServices.openUrl(QUrl(URL_PAGE))
            return
        total_mo = max(1, self.info["taille"] // (1024 * 1024))
        self._progression = QProgressDialog("Téléchargement de la mise à jour...", "Annuler",
                                            0, total_mo, self.fenetre)
        self._progression.setWindowTitle("Mise à jour de COURRIER")
        self._progression.setStyleSheet("QLabel { color: white; }" + self.style_boutons)
        self._progression.setWindowModality(Qt.WindowModality.NonModal)
        self._progression.setMinimumDuration(0)
        self._progression.setAutoClose(False)
        self._progression.setAutoReset(False)
        self._progression.setValue(0)
        self._progression.canceled.connect(self._annuler_telechargement)
        self._progression.show()

        self._telechargeur = _Telechargeur(self.info, self.dossier_travail, self.nom_exe)
        self._telechargeur.progression.connect(self._sur_progression)
        self._telechargeur.etape.connect(self._sur_etape)
        self._telechargeur.fini.connect(self._sur_telechargement_fini)
        self._telechargeur.erreur.connect(self._sur_erreur_telechargement)
        self._telechargeur.start()

    def _sur_progression(self, recu, total):
        if self._progression is None:
            return
        if total and total != self._progression.maximum():
            self._progression.setMaximum(total)
        self._progression.setValue(min(recu, self._progression.maximum()))
        self._progression.setLabelText(f"Téléchargement de la mise à jour... {recu} / "
                                       f"{self._progression.maximum()} Mo")

    def _sur_etape(self, texte):
        if self._progression is not None:
            self._progression.setRange(0, 0)  # barre "en cours" sans pourcentage
            self._progression.setLabelText(texte)

    def _fermer_progression(self):
        if self._progression is not None:
            try:
                self._progression.canceled.disconnect(self._annuler_telechargement)
            except (RuntimeError, TypeError):
                pass
            self._progression.close()
            self._progression = None

    def _annuler_telechargement(self):
        if self._telechargeur is not None:
            self._telechargeur.annule = True
        self._fermer_progression()

    def _sur_erreur_telechargement(self, message):
        self._fermer_progression()
        shutil.rmtree(self.dossier_travail, ignore_errors=True)
        self._boite(QMessageBox.Icon.Warning, "Mise à jour",
                    f"Le téléchargement de la mise à jour a échoué :\n{message}\n\n"
                    f"Rien n'a été modifié. COURRIER vous la proposera à nouveau "
                    f"au prochain démarrage.").exec()

    def _sur_telechargement_fini(self, dossier_nouvelle):
        self._fermer_progression()
        self.nouvelle = dossier_nouvelle
        # Quoi qu'il arrive ensuite, la mise a jour s'installera a la fermeture de COURRIER.
        QApplication.instance().aboutToQuit.connect(self._lancer_remplacement)
        self.redemarrer()

    # ---------- installation ----------
    def _autres_fenetres_ouvertes(self):
        return [w for w in QApplication.topLevelWidgets()
                if w.isVisible() and w is not self.fenetre
                and w.windowType() not in (Qt.WindowType.ToolTip, Qt.WindowType.Popup)]

    def redemarrer(self):
        # Tant que l'utilisateur n'a pas confirme le redemarrage, une fermeture de
        # COURRIER par lui-meme installe la mise a jour sans relancer COURRIER.
        self.relancer = False
        autres = self._autres_fenetres_ouvertes()
        if autres:
            # Cette boite se ferme AVANT que l'utilisateur s'occupe des autres fenetres :
            # tant qu'elle est affichee, elle bloque toutes les fenetres de COURRIER
            # (impossible alors de les fermer). On surveille ensuite leur fermeture.
            titres = "\n".join(f"  - {w.windowTitle() or 'fenêtre sans titre'}" for w in autres[:8])
            boite = self._boite(QMessageBox.Icon.Warning, "Mise à jour prête",
                                "La mise à jour est prête, mais d'autres fenêtres de "
                                "COURRIER sont encore ouvertes :\n" + titres)
            boite.setInformativeText("Terminez ou fermez-les (un mail en cours de rédaction "
                                     "serait perdu). Dès qu'elles seront toutes fermées, "
                                     "COURRIER vous proposera de redémarrer.\n\n"
                                     "Sinon, la mise à jour s'installera quand vous "
                                     "quitterez COURRIER.")
            b_attendre = boite.addButton("Je les ferme", QMessageBox.ButtonRole.AcceptRole)
            boite.addButton("À la fermeture de COURRIER", QMessageBox.ButtonRole.RejectRole)
            boite.exec()
            if boite.clickedButton() is b_attendre:
                self._attendre_fermeture_fenetres()
            return
        self._confirmer_redemarrage(deja_confirme=True)

    def _attendre_fermeture_fenetres(self):
        if getattr(self, "_minuterie_fenetres", None) is None:
            self._minuterie_fenetres = QTimer(self)
            self._minuterie_fenetres.setInterval(1000)
            self._minuterie_fenetres.timeout.connect(self._verifier_fenetres_fermees)
        self._minuterie_fenetres.start()

    def _verifier_fenetres_fermees(self):
        if self._autres_fenetres_ouvertes():
            return
        self._minuterie_fenetres.stop()
        self._confirmer_redemarrage(deja_confirme=False)

    def _confirmer_redemarrage(self, deja_confirme):
        if not deja_confirme:
            boite = self._boite(QMessageBox.Icon.Question, "Mise à jour prête",
                                "Les autres fenêtres sont fermées. Redémarrer COURRIER "
                                "maintenant pour installer la mise à jour ?")
            b_redemarrer = boite.addButton("Redémarrer", QMessageBox.ButtonRole.AcceptRole)
            boite.addButton("À la fermeture de COURRIER", QMessageBox.ButtonRole.RejectRole)
            boite.exec()
            if boite.clickedButton() is not b_redemarrer:
                return
            if self._autres_fenetres_ouvertes():  # rouverte entre-temps
                self.redemarrer()
                return
        self.relancer = True
        # Fermeture normale (enregistrements, arret propre des connexions), puis sortie.
        self.fenetre.close()
        QApplication.instance().quit()

    def _lancer_remplacement(self):
        if not self.nouvelle:
            return
        script = os.path.join(self.dossier_travail, "installer_mise_a_jour.ps1")
        try:
            # BOM UTF-8 : sinon Windows PowerShell 5 lit les accents du script de travers
            with open(script, "w", encoding="utf-8-sig") as f:
                f.write(SCRIPT_POWERSHELL)
            if os.path.exists(self.fichier_resultat):
                os.remove(self.fichier_resultat)
            subprocess.Popen(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                 "-WindowStyle", "Hidden", "-File", script,
                 "-IdProcessus", str(os.getpid()),
                 "-Installation", self.installation,
                 "-Nouvelle", self.nouvelle,
                 "-NomExe", self.nom_exe,
                 "-Resultat", self.fichier_resultat,
                 "-Version", self.info["version"] if self.info else "",
                 "-Relancer", "1" if self.relancer else "0"],
                creationflags=subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP,
                close_fds=True, cwd=self.dossier_donnees)
        except Exception as e:
            print(f"[MAJ] lancement du remplacement impossible : {e}")
        self.nouvelle = None


def demarrer(fenetre, version_actuelle, dossier_donnees, style_boutons=""):
    """Point d'entree appele par COURRIER au demarrage. Sans effet hors exe.
    style_boutons : STYLE_BOUTON de COURRIER (boutons rouges des boites de message)."""
    if not getattr(sys, "frozen", False) or not version_actuelle:
        return None
    gestionnaire = GestionnaireMiseAJour(fenetre, version_actuelle, dossier_donnees, style_boutons)
    fenetre._gestionnaire_mise_a_jour = gestionnaire  # garde l'objet en vie
    gestionnaire.demarrer()
    return gestionnaire
