# UCKK DOCX Table Styler — version simple

Cette version simplifie l'application :

- sélection d'un fichier `.docx` source ;
- sélection d'un fichier `.docx` de sortie ;
- sélection directe d'un seul fichier `.png` décoratif ;
- option pour ajouter ou non le PNG ;
- si le PNG est désactivé, l'application formate seulement les tableaux ;
- pas de prévisualisation.

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

Windows :

```bash
run_windows_simple.bat
```

macOS / Linux :

```bash
bash run_mac_linux_simple.sh
```

ou directement :

```bash
python uckk_docx_table_gui_simple.py
```

## Réglages conseillés

Pour un rendu KDP/DOCX sobre :

- Palette : `BW` pour noir et blanc, `COULEUR` pour pétrole/gris froid ;
- Largeur PNG max : `260` à `310` points ;
- Bordure : `0.5` ;
- Taille police en-tête : `8` ;
- Taille police corps : `8`.

## Utilisation sans PNG

Décoche simplement :

`Ajouter ce PNG avant chaque tableau`

L'application appliquera uniquement le style des tableaux.
