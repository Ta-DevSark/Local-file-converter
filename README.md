# Convertisseur Local — Windows 10/11

Application bureautique locale, sans IA et sans accès Internet pendant l'utilisation.

## Fonctions

- DOC, DOCX, ODT, RTF, TXT, PPT, PPTX, ODP, XLS, XLSX, ODS et CSV vers PDF.
- PDF vers TXT.
- PDF vers DOCX (reconstruction du texte extrait, avec séparation des pages).

## Dépendances

- **LibreOffice** : moteur de conversion des documents bureautiques vers PDF. Il doit être installé sur le PC.
- **PyMuPDF** : extraction du texte des PDF.
- **python-docx** : création de fichiers DOCX.

L'application ne transmet aucun fichier et n'effectue aucune requête réseau. Les dépendances sont intégrées dans l'exécutable lors de sa compilation.

## Créer l'exécutable Windows

La compilation doit être effectuée sur Windows (la génération d'un EXE Windows depuis cet environnement Linux n'est pas possible).

1. Installer Python 3.11 ou 3.12 (64 bits) sur le PC Windows.
2. Installer LibreOffice.
3. Extraire ce dossier.
4. Double-cliquer sur `build_windows.bat`.
5. L'exécutable sera créé dans `dist\\ConvertisseurLocal.exe`.

La première compilation nécessite Internet pour télécharger les dépendances Python. Une fois l'EXE créé, son utilisation est hors ligne. Pour une compilation elle-même hors ligne, il faut fournir à pip un dossier local contenant les wheels requis (`--no-index --find-links`).

## Utilisation

1. Lancer `ConvertisseurLocal.exe`.
2. Ajouter un ou plusieurs fichiers.
3. Choisir le sens de conversion.
4. Choisir le format TXT ou DOCX pour les PDF.
5. Choisir le dossier de destination (facultatif).
6. Cliquer sur Convertir.

Les fichiers existants ne sont pas écrasés : un suffixe `_1`, `_2`, etc. est ajouté.

## Limites importantes

- Un PDF scanné ou constitué uniquement d'images ne contient pas de texte extractible. Cette version ne fait pas d'OCR.
- PDF vers DOCX : le texte et les sauts de page sont reconstruits. La mise en page, les colonnes, tableaux, polices, images et graphiques ne sont pas reproduits fidèlement.
- Les résultats dépendent du fichier source et des capacités d'import/export de LibreOffice.
- Les fichiers protégés par mot de passe ou endommagés peuvent échouer.

## Architecture et choix techniques

- Python + Tkinter : interface native simple, sans framework graphique tiers.
- PyMuPDF : moteur rapide pour extraire le texte des PDF.
- python-docx : génération DOCX.
- LibreOffice en mode headless : moteur local éprouvé pour les conversions bureautiques.
- Traitement dans un thread séparé : l'interface reste réactive.
- Aucun serveur, aucune télémétrie, aucune dépendance à une API.

## Confidentialité

Les fichiers sont traités localement. Le code ne contient aucune fonction d'envoi, de synchronisation ou de télémétrie.
