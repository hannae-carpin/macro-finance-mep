# Macro Finance MEP — Anonymisation & Contrôle métier (Python / Excel VBA)

## Objectif
Mettre en place un **pipeline de démonstration sécurisé** permettant d’anonymiser des données financières sensibles sous Excel à l’aide de **Python**, puis de les **contrôler via une macro VBA** reproduisant des règles métier réelles (finance / comptabilité / administration).

Le projet illustre une approche pragmatique combinant **automatisation Python** et **outils Excel existants**, sans exposition de données sensibles.

---

## Données
- **Source** : fichier Excel financier interne (hors GitHub)
- **Format de démonstration** : fichier Excel `.xlsx` anonymisé
- **Nature des données** :
  - bénéficiaires
  - numéros de facture
  - montants financiers
  - dates (facture, échéance)
  - IBAN / BIC
  - pays de la banque

Toutes les données présentes dans ce dépôt sont **fictives et générées automatiquement**.

---

## Méthodologie
1. Chargement sécurisé du fichier Excel source
2. Détection automatique des colonnes sensibles (montants, dates, pays)
3. Anonymisation contrôlée :
   - randomisation cohérente des bénéficiaires
   - génération d’IBAN / BIC fictifs
   - dates réalistes (passé / présent / futur)
4. Injection volontaire d’anomalies métier
5. Export d’un fichier Excel de démonstration
6. Application d’une macro VBA de contrôle
7. Vérification du déclenchement des alertes attendues

Chaque type d’anomalie est **garanti au moins une fois** dans le fichier final.

---

## Anomalies simulées et règles (via macro VBA)

### Vérifier RIB
- **Colonne A – Bénéficiaire**
  - Valeur égale à `AKAMAI`

### Vérifier Numéro Facture
- **Colonne C – Numéro**
  - Vide
  - Commence par `TIT`
  - Se termine par `TIT`
  - Premier caractère non alphanumérique

### ≥ 800K€
- **Colonne P – Total à payer**
  - Montant supérieur ou égal à **800 000 €**

### Vérifier Date passée
- **Colonne R – Date de fin du compte**
  - Date antérieure à la date du jour

### Mettre en PG18 IBAN
- **Colonne S – IBAN**
  - Se termine par l’un des suffixes suivants :
    - `1623`
    - `3310`
    - `9742`
    - `43840`

### Mettre en PG03 BIC
- **Colonne T – BIC**
  - Vide
  - Commence par `TRPU`
  - Commence par `BDFEFRPP`

### Mettre RIB Bloqué
- **Colonne T – BIC**
  - Commence par l’un des codes suivants :
    - `NORDFRPP`
    - `TARNFR`
    - `COURTFR`
    - `KOLBFR`
    - `BNUGFR`
    - `RAPLFR`
    - `SMCTFR`
    - `SGBTMC`
    - `SBGDFRP`

### PAYS non autorisé
- **Colonne U – Pays de la banque**
  - Valeur différente de :
    - `FR`
    - `RE`
    - `MQ`
    - `GP`
    - `GF`
    - `PF`

---

## Résultats attendus (MVP)
- Générer un fichier Excel **testable immédiatement**
- Vérifier le bon déclenchement des règles VBA
- Simuler un cas réel de contrôle financier
- Fournir un support de démonstration sans risque RGPD

---

## Structure du repo
```
macro-finance-mep/
├─ data/
│  ├─ raw/               # fichier source local (ignoré par git)
│  └─ demo/              # fichier Excel anonymisé (.xlsx)
├─ scripts/
│  └─ anonymize_mep.py   # script Python principal
├─ vba/
│  └─ VirBU01_AUT.bas    # macro VBA exportée
├─ README.md
└─ requirements.txt
```

---

## How to run

### Prérequis
- Python 3.10+ recommandé
- Microsoft Excel (Windows)
- Git (optionnel)

### Installation (Windows PowerShell)
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Exécution
```powershell
python scripts/anonymize_mep.py
```

Le fichier de démonstration est généré dans :
```
data/demo/mep_anonymized.xlsx
```

---

## Utilisation de la macro VBA
1. Ouvrir `data/demo/mep_anonymized.xlsx`
2. Importer la macro depuis `vba/VirBU01_AUT.bas`
3. Lancer la macro
4. Consulter les anomalies détectées

---

## Sécurité & conformité
- Aucun fichier sensible versionné
- Données fictives uniquement
- Logique métier reproductible sans risque
- Projet diffusable publiquement

---

## Key takeaways (TL;DR)
- Combinaison **Python + VBA** orientée métier
- Automatisation pragmatique en environnement Excel
- Simulation réaliste de contrôles financiers
- Attention portée à la sécurité des données
- Projet directement testable par un recruteur
