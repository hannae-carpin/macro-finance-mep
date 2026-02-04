import random
import re
import string
from datetime import datetime, timedelta

from dateutil.relativedelta import relativedelta
import pandas as pd


# =========================
# CONFIG
# =========================
SEED = 20260204
random.seed(SEED)

INPUT_XLSX = "data/raw/input_sensible.xlsx"          
OUTPUT_XLSX = "data/demo/mep_anonymized.xlsx"             
SHEET_NAME = "MEP"
HEADER_ROW = 0

ERROR_RATE = 0.10  # 10% de lignes avec 1 anomalie max (en + des erreurs garanties)

# Colonnes macro (par NOM exact dans ton fichier)
COL_BY_NAME = {
    "A_BENEF": "Bénéficiaire",          # Excel A
    "C_FACTURE": "Numéro",              # Excel C
    "P_MONTANT": "Total à payer",       # Excel P
    "R_DATE": "Date de fin du compte",  # Excel R 
    "S_IBAN": "IBAN",                   # Excel S
    "T_BIC": "BIC",                     # Excel T
    "U_PAYS": "Pays de la banque",      # Excel U
}

# Pays OK (macro)
PAYS_OK = ["FR", "RE", "MQ", "GP", "GF", "PF"]

# Valeurs BIC bloquées / PG03
BIC_PG03_PREFIXES = ["TRPU", "BDFEFRPP"]
BIC_RIB_BLOQUE_PREFIXES = ["NORDFRPP", "TARNFR", "COURTFR", "KOLBFR", "BNUGFR", "RAPLFR", "SMCTFR", "SGBTMC", "SBGDFRP"]

# IBAN suffix PG18
IBAN_PG18_SUFFIXES = ["1623", "3310", "9742", "43840"]

# Bénéficiaires (base)
BENEFICIAIRES = [
    "ORANGE SA", "ENGIE FRANCE", "SNCF RESEAU", "EDF COMMERCE",
    "TOTALENERGIES", "BNP PARIBAS", "CREDIT AGRICOLE",
    "AXA FRANCE", "LA POSTE", "VEOLIA EAU",
    "SUEZ EAU FRANCE", "SPIE BATIGNOLLES", "VINCI ENERGIES",
    "BOUYGUES TELECOM", "CAPGEMINI FRANCE"
]

# Détection colonnes montant : keywords + EXCLUSIONS texte
AMOUNT_KEYWORDS = [
    "montant", "total", "impay", "reten", "escomp", "intér", "interet", "à payer", "a payer", "payer"
]
AMOUNT_EXCLUDE_KEYWORDS = [
    "adresse", "comment", "motif", "description", "benef", "fournisseur", "site", "source", "contrat"
]

COUNTRY_KEYWORDS = ["pays", "country"]


# =========================
# HELPERS
# =========================
def norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def is_amount_col(col: str) -> bool:
    c = norm(col)
    if any(x in c for x in AMOUNT_EXCLUDE_KEYWORDS):
        return False
    return any(k in c for k in AMOUNT_KEYWORDS)

def is_country_col(col: str) -> bool:
    c = norm(col)
    return any(k in c for k in COUNTRY_KEYWORDS)

def is_date_col(col: str, series: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    return "date" in norm(col)

def rand_digits(n: int) -> str:
    return "".join(random.choice(string.digits) for _ in range(n))

def rand_upper(n: int) -> str:
    return "".join(random.choice(string.ascii_uppercase) for _ in range(n))

def rand_alnum(n: int) -> str:
    return "".join(random.choice(string.ascii_uppercase + string.digits) for _ in range(n))

def ensure_columns_exist(df: pd.DataFrame) -> dict:
    mapping = {}
    for k, col in COL_BY_NAME.items():
        if col not in df.columns:
            raise ValueError(f"Colonne introuvable: {k} -> '{col}'. Colonnes: {list(df.columns)}")
        mapping[k] = col
    return mapping

def truncate_all_dates_no_time(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        s = df[col]
        if is_date_col(col, s):
            dt = pd.to_datetime(s, errors="coerce")
            df[col] = dt.dt.normalize()
    return df

def safe_set_numeric(df: pd.DataFrame, row_i: int, col: str, value: float):
    """
    Écrit un montant uniquement si la colonne est numérique (ou convertible).
    Sinon: skip (évite crash dtype string).
    """
    # si la colonne est "string dtype", on ne touche pas
    if pd.api.types.is_string_dtype(df[col]):
        return

    # si c'est object, on tente de convertir en numeric (mais sans casser)
    if pd.api.types.is_object_dtype(df[col]):
        # si ça ressemble à une colonne texte, on skip
        sample = df[col].dropna().astype(str).head(20)
        if (sample.str.contains(r"[A-Za-z]", regex=True).mean() > 0.2):
            return

    df.at[row_i, col] = float(value)


# =========================
# GENERATEURS (OK / anomalies)
# =========================
def gen_beneficiaire_ok() -> str:
    # randomise en gardant un style entreprise
    base = random.choice(BENEFICIAIRES)
    # petite variation contrôlée (pas trop pour rester réaliste)
    if random.random() < 0.25:
        suffix = random.choice(["", " - FR", " (FR)", " / DEP", " - SERVICES"])
        return f"{base}{suffix}".strip()
    return base

def gen_beneficiaire_rib_anomaly() -> str:
    # Macro: Vérifier RIB si A = AKAMAI
    return "AKAMAI"

def gen_facture_ok(i: int) -> str:
    return f"FA{i:06d}-{rand_digits(4)}"

def gen_facture_ko() -> str:
    # Macro: vide OU commence/termine TIT OU 1er char non alphanum
    r = random.random()
    if r < 0.25:
        return ""  # vide
    if r < 0.50:
        return "TIT" + rand_alnum(9)
    if r < 0.75:
        return rand_alnum(9) + "TIT"
    return "_" + rand_alnum(10)

def gen_amount_ok() -> float:
    return float(random.randint(50, 799_000))

def gen_amount_800k() -> float:
    return float(random.randint(800_000, 2_500_000))

def gen_date_ok() -> pd.Timestamp:
    today = datetime.today().date()
    return pd.Timestamp(today + timedelta(days=random.randint(0, 90)))

def gen_date_passee() -> pd.Timestamp:
    today = datetime.today().date()
    return pd.Timestamp(today - timedelta(days=random.randint(1, 365)))

def gen_iban_ok() -> str:
    return "FR" + rand_digits(2) + rand_digits(23)

def gen_iban_pg18() -> str:
    # doit se terminer par un suffix spécifique
    base = "FR" + rand_digits(2) + rand_digits(23)
    suf = random.choice(IBAN_PG18_SUFFIXES)
    return base[:-len(suf)] + suf

def gen_bic_ok() -> str:
    # BIC neutre : 8 ou 11 chars, commence pas par les préfixes bloqués
    bank = rand_upper(4)
    bic = f"{bank}FRPP{rand_alnum(3)}"
    if any(bic.startswith(p) for p in BIC_PG03_PREFIXES + BIC_RIB_BLOQUE_PREFIXES):
        bic = "ABCD" + bic[4:]
    return bic

def gen_bic_pg03() -> str:
    # Macro: vide OU commence par TRPU OU BDFEFRPP
    r = random.random()
    if r < 0.33:
        return ""
    return random.choice(BIC_PG03_PREFIXES) + rand_alnum(11 - len(random.choice(BIC_PG03_PREFIXES)))

def gen_bic_rib_bloque() -> str:
    prefix = random.choice(BIC_RIB_BLOQUE_PREFIXES)
    # complète pour faire 11
    return prefix + rand_alnum(max(0, 11 - len(prefix)))

def gen_pays_ok() -> str:
    return random.choice(PAYS_OK)

def gen_pays_ko() -> str:
    return random.choice(["DE", "ES", "IT", "US", "GB", "CH", "BE", "NL"])


# =========================
# MAIN
# =========================
def main():
    xls = pd.ExcelFile(INPUT_XLSX, engine="openpyxl")
    if SHEET_NAME not in xls.sheet_names:
        raise ValueError(f"Feuille '{SHEET_NAME}' introuvable. Feuilles dispo: {xls.sheet_names}")

    df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME, engine="openpyxl", header=HEADER_ROW)
    mapping = ensure_columns_exist(df)

    COL_A = mapping["A_BENEF"]
    COL_C = mapping["C_FACTURE"]
    COL_P = mapping["P_MONTANT"]
    COL_R = mapping["R_DATE"]
    COL_S = mapping["S_IBAN"]
    COL_T = mapping["T_BIC"]
    COL_U = mapping["U_PAYS"]

    print("\n=== COLONNES MACRO (noms) ===")
    print("A:", COL_A)
    print("C:", COL_C)
    print("P:", COL_P)
    print("R:", COL_R)
    print("S:", COL_S)
    print("T:", COL_T)
    print("U:", COL_U)

    # Détection auto des colonnes montants / pays
    amount_cols = [c for c in df.columns if is_amount_col(c)]
    country_cols = [c for c in df.columns if is_country_col(c)]

    print("\n=== DETECTION AUTO ===")
    print("Montants détectés:", amount_cols)
    print("Pays détectés    :", country_cols)

    # Cast textes sur colonnes sensibles (évite surprises)
    for col in [COL_A, COL_C, COL_S, COL_T, COL_U]:
        df[col] = df[col].astype("string")

    # Parse dates
    df[COL_R] = pd.to_datetime(df[COL_R], errors="coerce")

    # Parse montants sur colonnes montants (si détectées)
    for c in amount_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # -------------------------
    # GARANTIE : 1 anomalie de chaque type
    # -------------------------
    anomaly_types = [
        "VERIFIER_RIB",   # A = AKAMAI
        "FACTURE_KO",     # C
        "MONTANT_800K",   # P >= 800k (on force sur COL_P)
        "DATE_PASSEE",    # R < today
        "IBAN_PG18",      # S suffix
        "BIC_PG03",       # T vide ou TRPU/BDFEFRPP
        "RIB_BLOQUE",     # T commence par NORDFRPP etc.
        "PAYS_KO"         # U not in FR/RE/MQ/GP/GF/PF
    ]

    n = len(df)
    if n < len(anomaly_types):
        raise ValueError(f"Pas assez de lignes ({n}) pour garantir {len(anomaly_types)} anomalies.")

    forced_rows = random.sample(range(n), k=len(anomaly_types))
    forced_map = {forced_rows[i]: anomaly_types[i] for i in range(len(anomaly_types))}

    print("\n=== ERREURS GARANTIES ===")
    for row, typ in forced_map.items():
        print(f"Ligne df {row} -> {typ}")

    # -------------------------
    # Génération ligne par ligne
    # -------------------------
    for i in range(n):
        excel_row = i + 2  # header=0 => data starts at Excel row 2

        # défaut = OK
        benef = gen_beneficiaire_ok()
        facture = gen_facture_ok(excel_row)
        montant_p = gen_amount_ok()
        date_r = gen_date_ok()
        iban = gen_iban_ok()
        bic = gen_bic_ok()
        pays = gen_pays_ok()

        # anomalies forcées ou aléatoires
        if i in forced_map:
            a = forced_map[i]
        else:
            a = random.choice(anomaly_types) if (random.random() < ERROR_RATE) else None

        if a == "VERIFIER_RIB":
            benef = gen_beneficiaire_rib_anomaly()
        elif a == "FACTURE_KO":
            facture = gen_facture_ko()
        elif a == "MONTANT_800K":
            montant_p = gen_amount_800k()
        elif a == "DATE_PASSEE":
            date_r = gen_date_passee()
        elif a == "IBAN_PG18":
            iban = gen_iban_pg18()
        elif a == "BIC_PG03":
            bic = gen_bic_pg03()
        elif a == "RIB_BLOQUE":
            bic = gen_bic_rib_bloque()
        elif a == "PAYS_KO":
            pays = gen_pays_ko()

        # Écriture colonnes macro
        df.at[i, COL_A] = benef
        df.at[i, COL_C] = facture
        df.at[i, COL_S] = iban
        df.at[i, COL_T] = bic
        df.at[i, COL_U] = pays
        # R doit être date, pas texte
        df.at[i, COL_R] = pd.to_datetime(date_r)

        # P : montant principal (col P de ta macro)
        # -> elle est souvent numérique, mais on sécurise quand même
        # On ne force pas dtype string ici : on écrit une valeur float OK
        try:
            df[COL_P] = pd.to_numeric(df[COL_P], errors="coerce")
            df.at[i, COL_P] = float(montant_p)
        except Exception:
            # si vraiment P est du texte (bizarre), on met quand même une string propre
            df.at[i, COL_P] = str(int(montant_p))

        # Montants : toutes les colonnes montants détectées => OK (<800k)
        # Sauf si anomalie MONTANT_800K, on laisse le "gros montant" UNIQUEMENT sur COL_P
        for c in amount_cols:
            if c == COL_P:
                continue
            safe_set_numeric(df, i, c, gen_amount_ok())

        # Pays : toutes colonnes pays détectées (évite macro qui regarde ailleurs)
        for c in country_cols:
            df.at[i, c] = pays

    # Tronquer toutes les dates (00:00:00)
    df = truncate_all_dates_no_time(df)

    # Écriture : garde les autres feuilles inchangées
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for sh in xls.sheet_names:
            if sh == SHEET_NAME:
                df.to_excel(writer, sheet_name=sh, index=False)
            else:
                other = pd.read_excel(INPUT_XLSX, sheet_name=sh, engine="openpyxl")
                other.to_excel(writer, sheet_name=sh, index=False)

    print(f"\nOK → fichier anonymisé généré : {OUTPUT_XLSX}")
    print("Anomalies garanties :", anomaly_types)


if __name__ == "__main__":
    main()