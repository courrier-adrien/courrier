# COURRIER

Client mail multi-comptes pour Windows (IMAP / Outlook), écrit en Python avec PySide6 (Qt).

> *English: COURRIER is a free multi-account email client for Windows (IMAP / Outlook), written in Python / Qt. The interface is in French. Free for non-commercial use.*

## Fonctionnalités

- Plusieurs comptes mail dans la même fenêtre (IMAP, Outlook via Microsoft Graph).
- Carnet d'adresses, agenda, modèles, macros, filtrage des expéditeurs indésirables.
- Aperçu des pièces jointes (PDF, Word, Excel, PowerPoint…).
- Lecteur de flux RSS.
- **TEL / SMS** (Android uniquement) : afficher l'écran du téléphone sur le PC et envoyer des SMS depuis le PC, via [scrcpy](https://github.com/Genymobile/scrcpy).
- Outil **MAILING** pour l'envoi de lettres d'information.
- Toutes vos données restent sur votre PC (`%LOCALAPPDATA%\COURRIER`).

## Télécharger (version prête à l'emploi)

Rendez-vous dans l'onglet **[Releases](../../releases)**, téléchargez le zip, décompressez-le et lancez `COURRIER ADRIEN.exe`.
Les instructions détaillées (premier lancement, téléphone, désinstallation) sont dans [LISEZMOI.txt](LISEZMOI.txt).

**[Mode d'emploi complet](https://courrier-adrien.github.io/courrier/mode-emploi.html)** : toutes les fonctions, avec une case de recherche (aussi accessible dans le logiciel par le bouton AIDE).

Le programme n'est pas signé numériquement : Windows SmartScreen peut afficher un avertissement
(« Informations complémentaires » puis « Exécuter quand même »).

## Lancer depuis le code source

Prérequis : Windows et Python 3.13.

```bat
python -m venv venv
venv\Scripts\pip install -r requirements.txt
venv\Scripts\python COURRIER.py
```

Pour construire l'exécutable :

```bat
venv\Scripts\pip install pyinstaller
venv\Scripts\pyinstaller COURRIER.spec --noconfirm --clean
```

Le résultat se trouve dans `dist\COURRIER`.

### Comptes Outlook / Microsoft 365

La connexion Outlook utilise l'identifiant d'application (`CLIENT_ID_OAUTH_OUTLOOK` dans `COURRIER.py`)
enregistré par l'auteur. Si vous modifiez et redistribuez le programme, enregistrez votre propre
application dans [Microsoft Entra](https://entra.microsoft.com) et remplacez cet identifiant.

## Signaler un problème

Ouvrez une *issue* sur ce dépôt, ou écrivez à adrienmel@gmail.com.
En cas de plantage, joignez si possible `journal.txt` et `crash_traceback.txt` du dossier `%LOCALAPPDATA%\COURRIER`.

## Licence

[PolyForm Noncommercial 1.0.0](LICENSE) : vous pouvez utiliser, modifier et partager COURRIER
**pour un usage non commercial uniquement**.

Logiciels tiers : le bouton TEL utilise scrcpy et ses composants (FFmpeg, SDL, libusb, adb),
redistribués sous leurs propres licences : voir `ressources/scrcpy/`.
