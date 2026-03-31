# CEMAC Dashboard Generator 📊

Ce projet automatise la création d'un **Tableau de Bord mensuel CEMAC** au format Excel à partir des rapports de la BEAC (Rapport sur la Politique Monétaire — RPM et son Annexe statistique CPM).

## 🚀 Fonctionnalités principales

Le script `generate_cemac_dashboard.py` réalise les opérations suivantes :

1. **Détection automatique des PDFs** : identifie les derniers fichiers `Point 3_RPM*.pdf` et `Point 3_annexe*.pdf` présents dans le dossier, et en extrait la période (mois/année) pour nommer l'Excel de sortie.

2. **Extraction des tableaux** via `pdfplumber` (PDFs natifs — pas d'OCR) :
   - **T18 — Agrégats monnaie & crédit CEMAC** (Annexe p.26, idx 25) : extraction par positions de mots (`extract_words()`), indispensable pour gérer les nombres avec espace comme séparateur de milliers (ex : `2 485 304`). Deux sous-tableaux : Niveaux et Variations annuelles.
   - **T16 — Balance des paiements** (Annexe p.22, idx 21) : parsing ligne par ligne du texte extrait.
   - **T25 — Taux directeurs BEAC/BCE** (Annexe p.33, idx 32) : parsing ligne par ligne avec filtre anti-bruit pour ignorer les axes de graphiques G29/G30 qui s'affichent parfois en texte.
   - **T32 — Indicateurs CEMAC globaux** (Annexe p.39, idx 38) : extraction générique par années-colonnes.
   - **T33 à T38 — Données pays** (Annexe p.40–45, idx 39–44) : Cameroun, Centrafrique, Congo, Gabon, Guinée Équatoriale, Tchad — même fonction générique.

3. **Traçabilité & Audit** :
   - Chaque onglet `Source_*` documente le fichier PDF source, la page et la date/heure d'extraction.
   - **Surlignage orange automatique** : si un fichier `Dashboard_*.xlsx` du mois précédent existe dans le dossier, le script compare les valeurs et surligne en orange (#FFE0B2) toutes les cellules nouvelles ou modifiées. Une légende est ajoutée en ligne 6 de chaque onglet source.

4. **Dashboard principal** (`01_Dashboard`) :
   - Formaté A4 Paysage, prêt à l'export PDF.
   - Bloc fixe : composition des ministres de l'Économie de la zone CEMAC.
   - Bloc conjoncturel : données T18, T16, T25, T32 automatiquement insérées.
   - Zones grises `[PLACEHOLDER]` pour les données manuelles (graphiques, prix pétrole, TCER, COBAC).

## 📋 Cartographie des sources

| Tableau | Page Annexe | Index (0-based) | Contenu |
| :--- | :---: | :---: | :--- |
| T18 | p.26 | 25 | Agrégats monnaie & crédit CEMAC (niveaux + variations) |
| T16 | p.22 | 21 | Soldes balance des paiements (% PIB) |
| T25 | p.33 | 32 | Taux directeurs BEAC et BCE |
| T32 | p.39 | 38 | Indicateurs économiques CEMAC globaux |
| T33–T38 | p.40–45 | 39–44 | Indicateurs par pays (6 pays) |

## 🛠 Installation et Usage

### Prérequis

Python 3.8+ avec les bibliothèques suivantes :

```bash
pip install pdfplumber pandas xlsxwriter openpyxl
```

(`openpyxl` est utilisé en lecture pour comparer avec l'Excel du mois précédent.)

### Utilisation mensuelle

1. Déposez les deux PDFs du nouveau mois dans ce dossier :
   - `Point 3_RPM [mois] [année]_vf_.pdf`
   - `Point 3_annexe_Tableau_de_bord_CPM_[mois]_[année]_vN.pdf`
2. Lancez le script :
   ```bash
   python generate_cemac_dashboard.py
   ```
3. Le fichier `Dashboard_[mois]_[année].xlsx` est généré. Les cellules modifiées par rapport au mois précédent sont automatiquement surlignées en orange.
4. Complétez manuellement les zones `[PLACEHOLDER]` dans `01_Dashboard` (graphiques, prix pétrole, TCER, données COBAC).

## ⚠️ Notes techniques

- **Regroupement vertical par proximité** (`group_words_by_proximity`, seuil 5 pt) : utilisé pour T18 afin de fusionner les mots d'une même ligne logique dont les coordonnées `top` varient légèrement selon le mot. Plus robuste que le bucketing fixe qui crée des artefacts aux frontières de buckets.
- **Filtre anti-bruit T25** : les graphiques G29/G30 de l'Annexe génèrent des séquences de chiffres/lettres isolés dans le texte extrait. Le script détecte et stoppe l'extraction dès qu'une telle séquence apparaît (regex `^[\d\s\-]{20,}$` ou préfixe `G2`/`G3`).
- **Reconstruction du label T18** : "Taux de couverture extérieure de la monnaie (en %)" apparaît sur deux lignes dans le PDF ; le script le reconstruit automatiquement via `fix_tcem_labels()`.
- En cas d'échec d'extraction sur une page, le script continue et marque l'onglet source "Données non extraites".

## 📈 Étendre le scope

### Ajouter un nouveau tableau de l'annexe

Utilisez la fonction générique `extract_indicators_table(pdf_path, page_idx, table_label)` pour tout tableau structuré avec des années en colonnes (format similaire à T32–T38).

### Modifier le Dashboard principal

Toutes les sections du Dashboard sont dans `build_dashboard()`. Les données sont écrites via `write_df_block(ws, df, row)` qui prend directement un DataFrame.

### Mettre à jour les Ministres

La liste `ministres` dans `build_dashboard()` contient les noms et titres actuels. Modifier directement les entrées concernées.

---
*Développé pour la Mission CEMAC Expert.*
