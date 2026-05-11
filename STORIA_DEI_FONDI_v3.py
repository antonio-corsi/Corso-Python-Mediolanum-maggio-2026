# -*- coding: utf-8 -*-
# [AP] Questa riga è una dichiarazione di encoding (PEP 263) — dice all'interprete Python di leggere il file sorgente come UTF-8.
# [AP] Oggi è di fatto inutile, perché Python 3 usa già UTF-8 di default per i file sorgente. È un retaggio dei tempi di Python 2, 
# [AP] dove l'encoding di default era ASCII e senza quella riga i caratteri accentati (è, à, le emoji 🟣✅) sollevavano SyntaxError. 
# [AP] Spyder la inserisce automaticamente in cima ai nuovi script per abitudine — non perché serva. 
# [AP] Si può rimuovere senza alcun effetto sul comportamento del codice.

"""
Spyder Editor
This is a temporary script file.
"""

# [AP] È una docstring generata automaticamente da Spyder quando si crea un nuovo file Python vuoto tramite File → New file. 
# [AP] È il "template" predefinito che Spyder mette in cima a ogni nuovo script per ricordarci che è un file appena creato.
# [AP] In generale, le triple virgolette """...""" in Python delimitano una stringa multilinea. Quando una stringa di questo tipo
# [AP] si trova in cima a un file (come prima istruzione, prima di qualsiasi codice o import), Python la interpreta come 
# [AP] module docstring — la "documentazione" del modulo, accessibile programmaticamente con nome_modulo.__doc__.
# [AP] Non ha impatto sull'esecuzione. Python valuta la stringa, nessuno la usa, nessun comportamento cambia.
# [AP] Si può tranquillamente eliminare, ma visto che il file ha 3500+ righe e fa parecchio lavoro, una docstring informativa è 
# [AP] effettivamente utile per chi (anche l'autore, fra sei mesi) apre il file. Per esempio:

"""
STORIA_DEI_FONDI

Pipeline di calcolo Asset/Sector/Geo Allocation per le famiglie MIF, MGF, MGE.
Produce gli storici concatenati e i confronti per Power BI.

Input:  Tracciati e Portafogli mensili in C:/Users/Lenovo/Desktop/Mediolanum
Output: file .xlsx in C:/Users/Lenovo/Desktop/Mediolanum

Esecuzione: aprire in Spyder ed eseguire le celle (# %%) in ordine.
Aggiornare anno/mese nella prima cella prima di lanciare.

Autore: ....
"""


# %%
# [AP] sys, os, warnings, datetime e glob NON devono essere installati (dà errore) perchè appartenenti
# [AP] alla Python Standard Library (già pre-installata)
import pandas as pd
# import matplotlib.pyplot as plt  # [AP] non utilizzato
import numpy as np
import warnings 
import sys
from datetime import datetime
import os
import glob
# [AP] rischioso se non si conosce bene il comportamento dello script (magari per cose nuove)
warnings.filterwarnings("ignore")

# MESE
mese = '03'   # [AP]
mese_aggiornamento = mese

# ANNO
anno = '26'   # [AP]
anno_aggiornamento = anno


giorno = '31'                # [AP]
giorno_lavorativo = '31'     # [AP]

data = f"{mese}/20{anno}"
data_aggiornamento = f"{mese_aggiornamento}/20{anno_aggiornamento}"


# [AP] commentata la precedente 'append del path' ed inserita una nuova
# sys.path.append(r'G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/Region/Nuovi_dati_25_01/Region_Comitato/Codici')
sys.path.append(r'%USERPROFILE%/Desktop/Mediolanum')  # [AP] nuova append locale
# [AP] In Python, la r prima degli apici significa raw string, cioè “stringa grezza”. Significa: Python deve leggere quella stringa 
# [AP] così com’è, senza interpretare alcuni caratteri come sequenze speciali.
# [AP] Il caso più importante riguarda i percorsi Windows, che spesso in Windows usano il backslash (/ è lo standard Unix/Linux).
# [AP] In questo caso la r non è strettamente necessaria, perché il percorso usa / e non \.
# [AP] La raw string sarebbe invece indispensabile nel caso di r'C:\Users\Lenovo\Desktop\Mediolanum' (stile Windows).
# [AP] Per questo, in Python su Windows sono molto comuni queste due forme sicure:
# [AP] - "C:/Users/Lenovo/Desktop/Mediolanum"
# [AP] - r"C:\Users\Lenovo\Desktop\Mediolanum"
# [AP] Ancora meglio, in codice moderno, si può usare pathlib, che è più ordinato e indipendente dal sistema operativo:
# [AP] - from pathlib import Path
# [AP] - cartella = Path("C:/Users/Lenovo/Desktop/Mediolanum")
# [AP] oppure:
# [AP] - from pathlib import Path
# [AP] - cartella = Path(r"C:\Users\Lenovo\Desktop\Mediolanum")
# [AP] Path serve a gestire percorsi di file e cartelle in modo più pulito rispetto alle semplici stringhe.

# [AP] esame del path (più leggibile)
for p in sys.path:
    print(p)
    
# [AP] Quello che si vede in console è il contenuto di sys.path, cioè l’elenco delle cartelle in cui Python cerca:
# [AP] - moduli standard, cioè che fanno parte della libreria standard di Python, cioè che sono già inclusi quando si installa Python.
# [AP]   Quindi non bisogna installarli con pip o conda (dà errore!).
# [AP]   Ad esempio: os, sys, datetime, warnings, glob, subprocess, math, json, pathlib.
# [AP] - pacchetti installati;
# [AP] - file .py importabili;
# [AP] - eventuali cartelle aggiunte manualmente con sys.path.append(...).
# [AP] La nuova append serve solo a dire a Python: cerca eventuali moduli .py anche dentro C:/Users/Lenovo/Desktop/Mediolanum.
# [AP] Non confondere sys.path (che contiene percorsi di CODICE Python, cioè moduli e package) con le due variabili di ambiente PATH 
# [AP] di Windows che contengono i percorsi degli ESEGUIBILI (solo per quell'utente o per tutti gli utenti).

# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==============================================================
# ==============================================================

# Asset Allocation MIF

# ==============================================================
# ==============================================================


# Asset Allocation MIF Totale
from Asset_Allocation_Sector_Geo import MIF_ASSET_fn
MIF_ASSET_df, PIVOT_NAV_MIF, NAV_MIF = MIF_ASSET_fn (anno = anno, mese = mese, data = data)

# ---- Info ----
info_asset_allocation_mif_totale = {
    "Nome codice": ["STORIA_DEI_FONDI"],
    # "Percorso": [r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Codici - Storia dei fondi"],
    "Percorso": [r"C:\Users\Lenovo\Desktop\Mediolanum"],
    "Data stampa": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
}
df_info = pd.DataFrame(info_asset_allocation_mif_totale)

# ---- Percorsi file ----
# PERCORSO_asset_allocation_mif_totale = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\TRACCIATO ASSET MIF.xlsx"
PERCORSO_asset_allocation_mif_totale = r"C:\Users\Lenovo\Desktop\Mediolanum\TRACCIATO ASSET MIF.xlsx"

# ---- Salvataggio su Excel con più fogli ----
with pd.ExcelWriter(PERCORSO_asset_allocation_mif_totale, engine="openpyxl") as writer:
    MIF_ASSET_df.to_excel(writer, sheet_name="Sheet1", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)



# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣


# ==============================================================
# ==============================================================

# Storico MIF, MGF ed MGE 

# ==============================================================
# ==============================================================

from Storico_MIF_MGF_MGE import aggiorna_famiglia


famiglie = ["MIF", "MGF", "MGE"]

for famiglia in famiglie:
    print(f"\n🔄 Aggiornamento famiglia: {famiglia}")
    aggiorna_famiglia(famiglia, anno_aggiornamento, mese_aggiornamento)




# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==============================================================
# ==============================================================

# Storici MIF, MGF ed MGE concatenati

# ==============================================================
# ==============================================================

# percorso_storici = "G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi"
percorso_storici = "C:/Users/Lenovo/Desktop/Mediolanum"

famiglie = ["MIF", "MGF", "MGE"]
tipi_storico = [
    "ASSET MACRO",
    "ASSET MICRO",
    "SECTOR MACRO",
    "SECTOR MICRO",
    "COUNTRY AREA",
    "COUNTRY PAESE"
]

storici_concatenati = {}

# concatenazione

for tipo in tipi_storico:
    df_list = []

    for famiglia in famiglie:
        file_path = f"{percorso_storici}/STORICO {tipo} {famiglia}.xlsx"

        try:
            df = pd.read_excel(file_path)
            df["FAMIGLIA"] = famiglia
            df_list.append(df)
        except FileNotFoundError:
            print(f"⚠️ File non trovato: {file_path}")

    if df_list:
        storici_concatenati[tipo] = pd.concat(df_list, ignore_index=True)
    else:
        storici_concatenati[tipo] = pd.DataFrame()


# RENAME MGF → SMFI

for tipo, df in storici_concatenati.items():
    if "FAMIGLIA" in df.columns:
        df["FAMIGLIA"] = df["FAMIGLIA"].replace({"MGF": "SMFI"})
    storici_concatenati[tipo] = df


# SALVATAGGIO CONCATENATI + LONG + INFO

for tipo, df in storici_concatenati.items():

    if df.empty:
        print(f"⚠️ Nessun dato per tipo: {tipo}")
        continue

    # ------------------- VERSIONE LONG -------------------
    colonne_base = ["COLLECTION_DB_CODE", "TIM_REPORTING_DATE", "NAV", "NAV_TOTALE", "Data aggiornamento", "FAMIGLIA"]
    colonne_base_presenti = [c for c in colonne_base if c in df.columns]

    colonne_ribasato = [c for c in df.columns if c.endswith("_Ribasato")]

    df_sub = df[colonne_base_presenti + colonne_ribasato].copy()

    if colonne_ribasato:
        id_vars = [c for c in df_sub.columns if c not in colonne_ribasato]

        df_long = pd.melt(
            df_sub,
            id_vars=id_vars,
            value_vars=colonne_ribasato,
            var_name="Classe",
            value_name="Ribasato"
        )

        df_long["Classe"] = df_long["Classe"].str.replace("_Ribasato", "", regex=False)
    else:
        df_long = df_sub.copy()
        df_long["Classe"] = pd.NA
        df_long["Ribasato"] = pd.NA

    # ------------------- INFO -------------------
    info = {
        "Nome codice": ["STORIA_DEI_FONDI"],
        "Tipo storico": [tipo],
        "Percorso": [percorso_storici],
        "Data stampa": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    }
    df_info = pd.DataFrame(info)

    # ------------------- SALVATAGGIO -------------------
    percorso_save = f"{percorso_storici}/STORICO {tipo} CONCATENATO LONG.xlsx"

    with pd.ExcelWriter(percorso_save, engine="openpyxl") as writer:
        df_long.to_excel(writer, sheet_name="Sheet1", index=False)
        df_info.to_excel(writer, sheet_name="INFO", index=False)

    print(f"✅ Salvato LONG: {percorso_save}")

    # versione wide
    percorso_save_wide = f"{percorso_storici}/STORICO {tipo} CONCATENATO.xlsx"
    df.to_excel(percorso_save_wide, index=False)

    print(f"✅ Salvato: {percorso_save_wide}")




# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ======================================================================================
# ======================================================================================

# CONFRONTO MIF, MGF ed MGE A LIVELLO DI ASSET ALLOCATION - SECTOR - GEO

# ======================================================================================
# ======================================================================================

from Confronto_Asset_Allocation_Sector_Geo_NEW import Confronto_famiglie_fn
ASSET_concatenato, SECTOR_concatenato, COUNTRY_concatenato = Confronto_famiglie_fn (mese = mese, anno = anno, data = data)


# ---- Info ----
info = {
    "Nome codice": ["STORIA_DEI_FONDI"],
    # "Percorso": [r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Codici - Storia dei fondi"],
    "Percorso": [r"C:\Users\Lenovo\Desktop\Mediolanum"],
    "Data stampa": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
}
df_info = pd.DataFrame(info)

# ---- Percorsi file ----
# PERCORSO_ASSET = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Confronto Asset.xlsx"
PERCORSO_ASSET = r"C:\Users\Lenovo\Desktop\Mediolanum\Confronto Asset.xlsx"
# PERCORSO_SECTOR = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Confronto Sector.xlsx"
PERCORSO_SECTOR = r"C:\Users\Lenovo\Desktop\Mediolanum\Confronto Sector.xlsx"
# PERCORSO_COUNTRY = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Confronto Country.xlsx"
PERCORSO_COUNTRY = r"C:\Users\Lenovo\Desktop\Mediolanum\Confronto Country.xlsx"

# ---- Salvataggio su Excel con più fogli ----
with pd.ExcelWriter(PERCORSO_ASSET, engine="openpyxl") as writer:
    ASSET_concatenato.to_excel(writer, sheet_name="Sheet1", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)
    
# ---- Salvataggio su Excel con più fogli ----
with pd.ExcelWriter(PERCORSO_SECTOR, engine="openpyxl") as writer:
    SECTOR_concatenato.to_excel(writer, sheet_name="Sheet1", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)


# ---- Salvataggio su Excel con più fogli ----
with pd.ExcelWriter(PERCORSO_COUNTRY, engine="openpyxl") as writer:
    COUNTRY_concatenato.to_excel(writer, sheet_name="Sheet1", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)



# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣


# ======================================================================================
# ======================================================================================

# ASSET, SECTOR e GEO per MIF suddiviso per regione/paese 

# *** DETTAGLIO FUT SU OBB IN FONDO A QUESTA PARTE DI CODICE

# ⚠️⚠️⚠️ ATTENZIONE: TODO: -> qui c'è da aggiungere la distinzione tra MIF OBB e MIF AZ ⚠️⚠️⚠️

# CREIAMO UN PTF MIF AZIONARIO E UN PTF MIF OBBLIGAZIONARIO:
    # per Asset ovviamente non c'è bisogno
    # per il Sector andiamo a prendere COD_SECTOR che inizia con 'OB' e COD_SECTOR che inizia con 'AZ'
    # per la Geo possiamo andare direttamente a recuperare le diversificazione dai Quarterly
    # da qui poi per ogni fondo ci attacchiamo il nav per region

# ======================================================================================
# ======================================================================================

# Rimuovi eventuali import precedenti
if "ptf_mif_diviso_per_region" in sys.modules:
    del sys.modules["ptf_mif_diviso_per_region"]
    
    
# 1️⃣ Aggiungi il percorso corretto
# sys.path.append(r'G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/Region/Nuovi_dati_25_01/Region_Comitato/Codici')
sys.path.append(r'C:/Users/Lenovo/Desktop/Mediolanum')

# 2️⃣ Importa il modulo
from ptf_mif_diviso_per_region import ptf_mif_region_fn

# -------------------------------
# 1. Caricamento dati
# -------------------------------
TRACCIATO_MIF_ASSET, PTF_MIF, totale = ptf_mif_region_fn(
    anno=anno,
    mese=mese,
    giorno=giorno,
    giorno_lavorativo=giorno_lavorativo
)

# Escludo FOHF
PTF_MIF = PTF_MIF[PTF_MIF['COLLECTION_DB_CODE'] != 'DBFoHF']

# Totali regionali
PTF_MIF['TOTALE_NAV_SPAGNA'] = PTF_MIF[['NAV_SPAGNA_FEDE', 'NAV_SPAGNA_PIROZZI', 'gamax_es', 'FOF_SPAIN']].sum(axis=1)
PTF_MIF['TOTALE_NAV_GERMANIA'] = PTF_MIF[['NAV_GERMANIA_FEDE', 'NAV_GERMANIA_PIROZZI']].sum(axis=1)
PTF_MIF['TOTALE_NAV_ESTERO'] = PTF_MIF[['gamax_other']].sum(axis=1)
PTF_MIF['TOTALE_NAV_ITALIA'] = PTF_MIF['NAV'] - PTF_MIF['TOTALE_NAV_SPAGNA'].fillna(0) - PTF_MIF['TOTALE_NAV_GERMANIA'].fillna(0) - PTF_MIF['TOTALE_NAV_ESTERO'].fillna(0)

# Dizionario regioni
regioni_nav = {
    'TOTALE_NAV_ITALIA': 'ITALIA',
    'TOTALE_NAV_SPAGNA': 'SPAGNA',
    'TOTALE_NAV_GERMANIA': 'GERMANIA',
    'TOTALE_NAV_ESTERO': 'TERZI ESTERO'
}


# -------------------------------
# 2. Funzione ribasamento
# -------------------------------
def ribasa_region(df_tracciato, df_nav, col_totale_nav, famiglia_label, tracciato_type):
    # NAV per DB univoco
    mappa_nav = df_nav[['COLLECTION_DB_CODE', col_totale_nav]].drop_duplicates(subset='COLLECTION_DB_CODE')
    totale_nav = mappa_nav[col_totale_nav].sum()
    
    df = df_tracciato.merge(mappa_nav, on='COLLECTION_DB_CODE', how='left')
    df['TOTALE_Ribasato'] = df['TOTALE'] * df[col_totale_nav] / totale_nav
    df['TOTALE_Ribasato'] /= df['TOTALE_Ribasato'].sum()
    df['FAMIGLIA'] = famiglia_label
    df['NAV_FAMIGLIA'] = totale_nav
    
    # GEO: aggiungo Gruppi Area
    if tracciato_type == 'GEO':
        def label_group(area):
            if area.startswith('Mercati Emergenti'):
                return 'Mercati Emergenti'
            elif area in ['Area Euro', 'Area non Euro']:
                return 'Europa'
            else:
                return area
        df['Gruppi Area'] = df['DES_AREA'].apply(label_group)
    
    return df

# %%
# -------------------------------
# 3. Carico SECTOR e GEO
# -------------------------------
# path_tracciato_mif = f"G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Tracciati/Tracciato MIF {anno} {mese}.xlsx"
path_tracciato_mif = f"C:/Users/Lenovo/Desktop/Mediolanum/Tracciato MIF {anno} {mese}.xlsx"
df_sector = pd.read_excel(path_tracciato_mif, sheet_name='SECTOR')
df_country = pd.read_excel(path_tracciato_mif, sheet_name='COUNTRY')

tracciati_dict = {
    'ASSET': TRACCIATO_MIF_ASSET,
    'SECTOR': df_sector,
    'GEO': df_country
}

# -------------------------------
# 4. Ribasamento per tutte le regioni
# -------------------------------
df_asset_list = []
df_sector_list = []
df_geo_list = []

for col_nav, nome_regione in regioni_nav.items():
    for tracciato_type, df_tracciato in tracciati_dict.items():
        df_rib = ribasa_region(df_tracciato, PTF_MIF, col_nav, f"MIF {nome_regione}", tracciato_type)
        
        # Seleziono pivot e colonne chiave diverse per ogni tracciato
        if tracciato_type == 'ASSET':
            df_pivot = df_rib.pivot_table(index=['DES_ASSET', 'COD_MACRO_ASSET'], values='TOTALE_Ribasato', aggfunc='sum').reset_index()
            df_pivot['FAMIGLIA'] = df_rib['FAMIGLIA'].iloc[0]
            df_pivot['NAV_FAMIGLIA'] = df_rib['NAV_FAMIGLIA'].iloc[0]
            df_asset_list.append(df_pivot)
        elif tracciato_type == 'SECTOR':
            df_pivot = df_rib.pivot_table(index=['COD_SECTOR', 'DES_SECTOR'], values='TOTALE_Ribasato', aggfunc='sum').reset_index()
            df_pivot['FAMIGLIA'] = df_rib['FAMIGLIA'].iloc[0]
            df_pivot['NAV_FAMIGLIA'] = df_rib['NAV_FAMIGLIA'].iloc[0]
            df_sector_list.append(df_pivot)
        elif tracciato_type == 'GEO':
            df_pivot = df_rib.pivot_table(index=['DES_PAESE','DES_AREA','Gruppi Area'], values='TOTALE_Ribasato', aggfunc='sum').reset_index()
            df_pivot['FAMIGLIA'] = df_rib['FAMIGLIA'].iloc[0]
            df_pivot['NAV_FAMIGLIA'] = df_rib['NAV_FAMIGLIA'].iloc[0]
            df_geo_list.append(df_pivot)

# -------------------------------
# 5. Concatenazione finale
# -------------------------------
riga_mif_asset = pd.concat(df_asset_list, ignore_index=True)
riga_mif_sector = pd.concat(df_sector_list, ignore_index=True)
riga_mif_geo = pd.concat(df_geo_list, ignore_index=True)

# Uniformità COD_MACRO_ASSET
mappa_rinomina_asset = {'AZ': 'Azioni', 'OB': 'Obbligazioni', 'LQ': 'Liquidità', 'AA': 'Altro'}
riga_mif_asset['COD_MACRO_ASSET'] = riga_mif_asset['COD_MACRO_ASSET'].replace(mappa_rinomina_asset)


# aggiungiamo la colonna NAV_MIF come somma dei valori univoci di NAV_FAMIGLIA
# poi facciamo la colonna ptf_mif: TOTALE_Ribasato * NAV_FAMIGLIA / NAV_MIF

# %%
# -------------------------------
# 6. Calcolo NAV_MIF e PTF_MIF
# -------------------------------

def aggiungi_ptf(df):
    # NAV_MIF come somma dei valori univoci di NAV_FAMIGLIA
    nav_mif = df['NAV_FAMIGLIA'].unique().sum()
    df['NAV_MIF'] = nav_mif
    
    # Calcolo PTF_MIF
    df['ptf_mif'] = df['TOTALE_Ribasato'] * df['NAV_FAMIGLIA'] / df['NAV_MIF']
    
    return df

riga_mif_asset = aggiungi_ptf(riga_mif_asset)
riga_mif_sector = aggiungi_ptf(riga_mif_sector)
riga_mif_geo = aggiungi_ptf(riga_mif_geo)

# Controllo
print(riga_mif_asset.head())
print(riga_mif_sector.head())
print(riga_mif_geo.head())

# Controllo
print(riga_mif_asset['ptf_mif'].sum())
print(riga_mif_sector['ptf_mif'].sum())
print(riga_mif_geo['ptf_mif'].sum())
    
# aggiungiamo la colonna data prima del salvataggio    

riga_mif_asset['Data'] = data
riga_mif_sector['Data'] = data
riga_mif_geo['Data'] = data



# PERCORSO_FILE = r"G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/tracciato_asset_paese_mif.xlsx"
PERCORSO_FILE = r"C:/Users/Lenovo/Desktop/Mediolanum/tracciato_asset_paese_mif.xlsx"

# ---- Info ----
info = {
    "Nome codice": ["STORIA_DEI_FONDI"], 
    "Nome codice di richiamo": ["MIF_per_region_ex_fohf"],
    # "Percorso": [r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Codici - Storia dei fondi"],
    "Percorso": [r"C:\Users\Lenovo\Desktop\Mediolanum"],
    "Data stampa": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
}
df_info = pd.DataFrame(info)

# ---- Salvataggio Asset con Info ----
with pd.ExcelWriter(
    # r"G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/tracciato_asset_paese_mif.xlsx",
    r"C:/Users/Lenovo/Desktop/Mediolanum/tracciato_asset_paese_mif.xlsx",
    engine="openpyxl"
) as writer:
    riga_mif_asset.to_excel(writer, sheet_name="Dati", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)

# ---- Salvataggio Sector con Info ----
with pd.ExcelWriter(
    # r"G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/tracciato_asset_paese_mif_Sector.xlsx",
    r"C:/Users/Lenovo/Desktop/Mediolanum/tracciato_asset_paese_mif_Sector.xlsx",
    engine="openpyxl"
) as writer:
    riga_mif_sector.to_excel(writer, sheet_name="Dati", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)

# ---- Salvataggio Country con Info ----
with pd.ExcelWriter(
    # r"G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/tracciato_asset_paese_mif_Country.xlsx",
    r"C:/Users/Lenovo/Desktop/Mediolanum/tracciato_asset_paese_mif_Country.xlsx",
    engine="openpyxl"
) as writer:
    riga_mif_geo.to_excel(writer, sheet_name="Dati", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)



# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==========================================================
# 🌍💼 PTF MIF - OBBLIGAZIONARIO PER REGION (Italia, Spagna, Germania, Estero)
# ==========================================================

# -----------------------------
# 📥 1️⃣ PREPARAZIONE DATI (TUTTI I FONDI)
# -----------------------------
# Seleziono solo le colonne fino a NAV + le colonne dei NAV regionali
cols_fino_nav = PTF_MIF.columns[:PTF_MIF.columns.get_loc("NAV") + 1]
cols_nav_region = ["TOTALE_NAV_ITALIA", "TOTALE_NAV_SPAGNA", "TOTALE_NAV_GERMANIA", "TOTALE_NAV_ESTERO"]
cols_finali = list(cols_fino_nav) + cols_nav_region

PTF_MIF_OBB = PTF_MIF[cols_finali].copy()

# -----------------------------
# 🏦 2️⃣ FILTRO OBBLIGAZIONARIO
# -----------------------------
PTF_MIF_OBB = PTF_MIF_OBB[PTF_MIF_OBB['COD_SECTOR'].str.startswith('OB', na=False)].copy()

# -----------------------------
# 🔁 3️⃣ RIBILANCIAMENTO PER FONDO (solo OB)
# -----------------------------
PTF_MIF_OBB['Ribilanciamento_Fondo'] = (
    PTF_MIF_OBB.groupby('COLLECTION_DB_CODE')['HOST_FUND_WEIGHTING']
    .transform(lambda x: x / x.sum())
)

# -----------------------------
# 🌍 4️⃣ CALCOLO TOTALE_RIBASATO PER REGION
# -----------------------------
regioni = ["TOTALE_NAV_ITALIA", "TOTALE_NAV_SPAGNA", "TOTALE_NAV_GERMANIA", "TOTALE_NAV_ESTERO"]

pivot_region_list = []

for region in regioni:
    # Drop righe senza NAV per quella region
    df_region = PTF_MIF_OBB.dropna(subset=[region])
    df_region = df_region[df_region[region] > 0].copy()

    if df_region.empty:
        continue  # Se non ci sono righe per questa region, salta

    # NAV del settore per fondo nella region
    df_region['NAV_OBB_SETT_FONDO'] = df_region['Ribilanciamento_Fondo'] * df_region[region]

    # Totale NAV obbligazionario MIF per la region (somma su tutti i fondi)
    NAV_tot_region = df_region['NAV_OBB_SETT_FONDO'].sum()
    df_region['NAV_MIF_REGION_OBB'] = NAV_tot_region

    # Peso sul totale MIF REGION (TOTALE_Ribasato) -> somma su tutta la region = 1
    df_region['TOTALE_Ribasato'] = df_region['NAV_OBB_SETT_FONDO'] / NAV_tot_region

    # Aggiungo colonna famiglia
    df_region['FAMIGLIA'] = 'MIF'

    # Aggiungo nome region per chiarezza
    df_region['REGION'] = region.replace('TOTALE_NAV_', '')

    # Aggiungo alla lista per concatenazione finale
    pivot_region_list.append(df_region)

# -----------------------------
# 🔗 5️⃣ CONCATENA TUTTE LE REGIONI
# -----------------------------
PTF_OBB_REGION = pd.concat(pivot_region_list, ignore_index=True)

# -----------------------------
# ✅ 6️⃣ CHECK FINALI
# -----------------------------
# Ribilanciamento fondo: deve tornare 1
print("✅ Check Ribilanciamento per fondo (solo OB):")
print(PTF_OBB_REGION.groupby('COLLECTION_DB_CODE')['Ribilanciamento_Fondo'].sum())

# TOTALE_Ribasato per region: deve tornare 1
print("\n✅ Check TOTALE_Ribasato per region:")
print(PTF_OBB_REGION.groupby('REGION')['TOTALE_Ribasato'].sum())



# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ----------------------------------------------------
# LO STESSO PER AZ - AZALT DA ESCLUDERE E RIBASERE
# ----------------------------------------------------

# ==========================================================
# 🌍 PTF MIF REGION - AZIONARIO SOLO (AZ)
# ==========================================================

# 🔹 1️⃣ Selezione colonne rilevanti (fino a NAV + NAV regionali)
cols_fino_nav = PTF_MIF.columns[:PTF_MIF.columns.get_loc("NAV") + 1]
cols_nav_region = ["TOTALE_NAV_ITALIA", "TOTALE_NAV_SPAGNA", "TOTALE_NAV_GERMANIA", "TOTALE_NAV_ESTERO"]
cols_finali = list(cols_fino_nav) + cols_nav_region
PTF_MIF_AZ = PTF_MIF[cols_finali].copy()

# 🔹 2️⃣ Definizione regioni
regioni = ["TOTALE_NAV_ITALIA", "TOTALE_NAV_SPAGNA", "TOTALE_NAV_GERMANIA", "TOTALE_NAV_ESTERO"]

# Lista per concatenazione finale
pivot_region_list = []

# 🔹 3️⃣ Iterazione sulle regioni
for region in regioni:
    
    # Filtro righe con NAV disponibile
    df_region = PTF_MIF_AZ.dropna(subset=[region])
    df_region = df_region[df_region[region] > 0].copy()
    
    if df_region.empty:
        continue  # Se non ci sono righe per questa region, salto
    
    # Filtro solo settori azionari (gestione NaN)
    df_region = df_region[df_region['COD_SECTOR'].str.startswith('AZ', na=False)].copy()
    
    # Escludo COD_SECTOR specifici
    df_region = df_region[~df_region['COD_SECTOR'].isin(['AZALT'])].copy()
    
    if df_region.empty:
        continue  # Se dopo i filtri non rimane nulla
    
    # Ribilanciamento per fondo (somma a 1 per fondo)
    df_region['Ribilanciamento_Fondo'] = df_region.groupby('COLLECTION_DB_CODE')['HOST_FUND_WEIGHTING'] \
                                                  .transform(lambda x: x / x.sum())
    
    # Calcolo NAV azionario ribasato per fondo
    df_region['NAV_AZ_SETT_FONDO'] = df_region['Ribilanciamento_Fondo'] * df_region[region]
    
    # Totale NAV azionario per la region
    NAV_tot_region = df_region['NAV_AZ_SETT_FONDO'].sum()
    df_region['NAV_MIF_REGION_AZ'] = NAV_tot_region
    
    # TOTALE_Ribasato (peso sul totale della region)
    df_region['TOTALE_Ribasato'] = df_region['NAV_AZ_SETT_FONDO'] / NAV_tot_region
    
    # Colonna famiglia e region
    df_region['FAMIGLIA'] = 'MIF'
    df_region['REGION'] = region.replace('TOTALE_NAV_', '')
    
    # Aggiungo alla lista finale
    pivot_region_list.append(df_region)

# 🔹 4️⃣ Concatenazione finale dataframe AZ pulito
PTF_AZ_REGION = pd.concat(pivot_region_list, ignore_index=True)

# %%
# ---------------------------------------------------------
# ✅ 5️⃣ Check finali
# ---------------------------------------------------------
print("✅ Check Ribilanciamento per fondo (somma = 1):")
print(PTF_AZ_REGION.groupby('COLLECTION_DB_CODE')['Ribilanciamento_Fondo'].sum())

print("\n✅ Check TOTALE_Ribasato per region (somma = 1):")
print(PTF_AZ_REGION.groupby('REGION')['TOTALE_Ribasato'].sum())


# da qui
# per ogni region, per ogni COD_MACRO_SECTOR devo ribilinciare il TOTALE_Ribasato
# cioè il TOTALE_Ribasato_macro deve tornare a 1 per ogni COD_MACRO_SECTOR per ogni REGION

#🔹 Obbligazionario 

PTF_OBB_REGION['TOTALE_Ribasato_macro'] = (
    PTF_OBB_REGION['TOTALE_Ribasato'] /
    PTF_OBB_REGION.groupby(['REGION', 'COD_MACRO_SECTOR'])['TOTALE_Ribasato'].transform('sum')
)

# Controllo rapido: somma interna per ogni macro settore deve essere 1


print("\n✅ Check TOTALE_Ribasato_macro OBB per REGION + COD_MACRO_SECTOR (deve fare 1):")
print(
    PTF_OBB_REGION
    .groupby(['REGION', 'COD_MACRO_SECTOR'])['TOTALE_Ribasato_macro']
    .sum()
)

#🔹 Azionario 

PTF_AZ_REGION['TOTALE_Ribasato_macro'] = (
    PTF_AZ_REGION['TOTALE_Ribasato'] /
    PTF_AZ_REGION.groupby(['REGION', 'COD_MACRO_SECTOR'])['TOTALE_Ribasato'].transform('sum')
)

# Controllo rapido: somma interna per ogni macro settore deve essere 1

print("\n✅ Check TOTALE_Ribasato_macro per REGION + COD_MACRO_SECTOR (deve fare 1):")
print(
    PTF_AZ_REGION
    .groupby(['REGION', 'COD_MACRO_SECTOR'])['TOTALE_Ribasato_macro']
    .sum()
)




# PERCORSO_FILE = r"G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/mif_suddiviso_per_region_az_obb.xlsx" 
PERCORSO_FILE = r"C:/Users/Lenovo/Desktop/Mediolanum/mif_suddiviso_per_region_az_obb.xlsx" 
# ---- Info ---- 
info = { "Nome codice": ["STORIA_DEI_FONDI"], 
        # "Percorso": [r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Codici - Storia dei fondi"], 
        "Percorso": [r"C:\Users\Lenovo\Desktop\Mediolanum"], 
        "Data stampa": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")] 
        } 
df_info = pd.DataFrame(info) 

with pd.ExcelWriter(PERCORSO_FILE, engine="openpyxl") as writer: 
    # Scrivo AZ pulito 
    PTF_AZ_REGION.to_excel(writer, sheet_name="AZ", index=False) 
    # Scrivo OBB pulito 
    PTF_OBB_REGION.to_excel(writer, sheet_name="OBB", index=False) 
    # Scrivo info 
    df_info.to_excel(writer, sheet_name="Info", index=False)



# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣


# DETTAGLIO FUT SU OBB -> sheet da aggiungere a Sector
# qui a partire dal PTF MIF suddiviso per region ci teniamo anche DES_ASSET 
# perché vogliamo sapere delle obbligazioni governative minori di 1 anno 
# quanta parte è Future su Obbligazioni Future su Obbligazioni

# quindi su PTF_MIF facciamo 4 pivot, una per regione: quindi una che avrà TOTALE_NAV_ITALIA;
# una che avrà TOTALE_NAV_SPAGNA; una per TOTALE_NAV_GERMANIA; e una per una che avrà TOTALE_NAV_TERZI_ESTERO (-> questi però in ogni pivot si chiamranno uguale e "NAV_FAMIGLIA").
# ed in queste pivot teniamo: 
    # COLLECTION_DB_CODE
    # COD_SECTOR
    # DES_SECTOR
    # DES_ASSET

    # poi ci calcoliamo il TOTALE_Ribasato in ogni pivot ed aggiungiamo la colonna FAMIGLIA 
    # poi conceteniamo tutte e quattro le pivot ed aggiungaimo il NAV_MIF come somma dei valori univoci dei NAV_FAMIGLIA
    # calcoliamo il ptf_mif sul NAV_MIF ed aggiungiamo la data


# andiamo prima a sistemare i vuoti 
# Creiamo una copia del PTF_MIF


PTF_MIF_corr = PTF_MIF.copy()

# -------------------------------
# 1️⃣ Pivot di partenza
# -------------------------------
pivot_base = (
    PTF_MIF_corr
    .groupby(
        ['COLLECTION_DB_CODE', 'DES_SECTOR', 'COD_SECTOR', 'DES_ASSET', 'COD_MACRO_ASSET', 'NAV'],
        as_index=False,
        dropna=False
    )
    .agg({'HOST_FUND_WEIGHTING':'sum'})
    .rename(columns={'HOST_FUND_WEIGHTING':'Somma_HOST_FUND_WEIGHTING'})
)

# Aggiungiamo la colonna di check: somma per COLLECTION_DB_CODE
pivot_base['Check_Somma_HOST_FUND_WEIGHTING'] = pivot_base.groupby('COLLECTION_DB_CODE')['Somma_HOST_FUND_WEIGHTING'].transform('sum')

pivot_base.loc[pivot_base['COLLECTION_DB_CODE'] == 'DB8825', 
               ['COLLECTION_DB_CODE', 'Check_Somma_HOST_FUND_WEIGHTING']].drop_duplicates()


condizione_future = (
    (pivot_base['DES_ASSET'] == 'Future su Azioni') &
    (pivot_base['DES_SECTOR'].isna() | (pivot_base['DES_SECTOR'] == ''))
)

pivot_base.loc[condizione_future, 'DES_SECTOR'] = 'Altre Azioni'
pivot_base.loc[condizione_future, 'COD_SECTOR'] = 'AZALT'
pivot_base.loc[condizione_future, 'COD_MACRO_ASSET'] = 'AZAA'

check_DB8825 = pivot_base[
    pivot_base['COLLECTION_DB_CODE'] == 'DB8825'
]


# -------------------------------
# 2️⃣ Funzione per creare tracciati regionali
# -------------------------------
def crea_tracciato_region(df_base, df_mif, col_nav_totale, regione_label):
    """
    df_base: pivot di base
    df_mif: PTF_MIF con NAV per regione
    col_nav_totale: colonna NAV per la regione, es. 'TOTALE_NAV_ITALIA'
    regione_label: etichetta per FAMIGLIA, es. 'ITALIA'
    """
    df_region = df_base.copy()
    
    # Mappa NAV_FONDO per COLLECTION_DB_CODE
    mappa_nav = df_mif[['COLLECTION_DB_CODE', col_nav_totale]].drop_duplicates(subset='COLLECTION_DB_CODE')
    df_region = df_region.merge(mappa_nav, on='COLLECTION_DB_CODE', how='left')
    
    # Rinomino colonna NAV_FONDO
    df_region.rename(columns={col_nav_totale:'NAV_FONDO'}, inplace=True)
    
    # Aggiungo colonna FAMIGLIA
    df_region['FAMIGLIA'] = f"MIF {regione_label}"
    
    return df_region

# %%
# -------------------------------
# 3️⃣ Dizionario regioni
# -------------------------------
regioni_nav = {
    'TOTALE_NAV_ITALIA': 'ITALIA',
    'TOTALE_NAV_SPAGNA': 'SPAGNA',
    'TOTALE_NAV_GERMANIA': 'GERMANIA',
    'TOTALE_NAV_ESTERO': 'TERZI ESTERO'
}

# -------------------------------
# 4️⃣ Creo i tracciati regionali
# -------------------------------
tracciati_regionali = {}
for col_nav, label in regioni_nav.items():
    tracciati_regionali[label] = crea_tracciato_region(pivot_base, PTF_MIF, col_nav, label)

# -------------------------------
# 5️⃣ Aggiungo NAV_FAMIGLIA dai totali
# -------------------------------
col_nav_totale_map = {
    'ITALIA': 'TOTALE_NAV_ITALIA',
    'SPAGNA': 'TOTALE_NAV_SPAGNA',
    'GERMANIA': 'TOTALE_NAV_GERMANIA',
    'TERZI ESTERO': 'TOTALE_NAV_ESTERO'
}

for regione, df_reg in tracciati_regionali.items():
    col_nav = col_nav_totale_map[regione]
    
    # Prendo il valore scalare dal dataframe 'totale'
    nav_famiglia = totale[col_nav].iloc[0]
    
    # Aggiungo colonna NAV_FAMIGLIA
    df_reg['NAV_FAMIGLIA'] = nav_famiglia
    
    # Aggiorno il dizionario
    tracciati_regionali[regione] = df_reg
    
    
# -------------------------------
# Check e aggiunta colonna somma per COLLECTION_DB_CODE
# -------------------------------
for regione, df_reg in tracciati_regionali.items():
    # Calcolo la somma per COLLECTION_DB_CODE
    sum_per_code = df_reg.groupby('COLLECTION_DB_CODE')['Somma_HOST_FUND_WEIGHTING'] \
                         .transform('sum')
    
    # Aggiungo la colonna SUM_HOST_FUND_WEIGHTING
    df_reg['SUM_HOST_FUND_WEIGHTING'] = sum_per_code
    
    # Aggiorno il dict
    tracciati_regionali[regione] = df_reg
    

for regione, df_reg in tracciati_regionali.items():
    
    df_reg['ptf_mif'] = (
        df_reg['Somma_HOST_FUND_WEIGHTING'] *
        df_reg['NAV_FONDO'] /
        df_reg['NAV_FAMIGLIA']
    )
    
    tracciati_regionali[regione] = df_reg
    
    somma = df_reg['ptf_mif'].sum()
    
    # normalizza sempre
    df_reg['ptf_mif'] = df_reg['ptf_mif'] / df_reg['ptf_mif'].sum()
    
    tracciati_regionali[regione] = df_reg
    
    print(regione, df_reg['ptf_mif'].sum())


# Lista dei df regionali
df_list = list(tracciati_regionali.values())

# Concatenazione in un unico df
tracciato_asset_sector_paese_mif = pd.concat(df_list, ignore_index=True)

# Aggiunta colonna Data
tracciato_asset_sector_paese_mif['Data'] = data


# ---- Salvataggio Sector con Info ----
with pd.ExcelWriter(
    # r"G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/tracciato_asset_paese_mif_Sector.xlsx",
    r"C:/Users/Lenovo/Desktop/Mediolanum/tracciato_asset_paese_mif_Sector.xlsx",
    engine="openpyxl"
) as writer:
    riga_mif_sector.to_excel(writer, sheet_name="Dati", index=False)
    tracciato_asset_sector_paese_mif.to_excel(writer, sheet_name="Asset_Sector", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)
    



# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==============================================================
# ==============================================================

# Sector MIF

# ==============================================================
# ==============================================================


# Sector MIF Totale

from Asset_Allocation_Sector_Geo import MIF_SECTOR_fn
MIF_SECTOR_df, PIVOT_NAV_MIF, NAV_MIF = MIF_SECTOR_fn (anno = anno, mese = mese, data = data)

# ---- Info ----
info_sector_mif_totale = {
    "Nome codice": ["STORIA_DEI_FONDI"],
    # "Percorso": [r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Codici - Storia dei fondi"],
    "Percorso": [r"C:\Users\Lenovo\Desktop\Mediolanum"],
    "Data stampa": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
}
df_info = pd.DataFrame(info_sector_mif_totale)

# ---- Percorsi file ----
# PERCORSO_sector_mif_totale = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\TRACCIATO SECTOR MIF.xlsx"
PERCORSO_sector_mif_totale = r"C:\Users\Lenovo\Desktop\Mediolanum\TRACCIATO SECTOR MIF.xlsx"

# ---- Salvataggio su Excel con più fogli ----
with pd.ExcelWriter(PERCORSO_sector_mif_totale, engine="openpyxl") as writer:
    MIF_SECTOR_df.to_excel(writer, sheet_name="Sheet1", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)



# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==============================================================
# ========================================================

# MIF totale settoriale diviso tra AZ e OBB - AZCI, AZDI, OBCO e OBGO + dettaglio Paese

# ==============================================================
# ========================================================

# PER MIF TOTALE

# UNIONE SECTOR + COUNTRY
# SUDDIVISIONE DEL PTF MIF IN PTF MIF OBB E PTF MIF AZ
# SUDDIVISIONE DEL PTF MIF OBB IN OBGO E OBCO

# TRACCIATO_MIF_SECTOR = pd.read_excel(f'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Tracciati/Tracciato MIF {anno} {mese}.xlsx', sheet_name = 'SECTOR')
TRACCIATO_MIF_SECTOR = pd.read_excel(f'C:/Users/Lenovo/Desktop/Mediolanum/Tracciato MIF {anno} {mese}.xlsx', sheet_name = 'SECTOR')

# PTF_MIF = pd.read_excel(f'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Portafoglio/PTF_FUNDLOOKTHROUGH MIF {anno} {mese}.xlsx')
PTF_MIF = pd.read_excel(f'C:/Users/Lenovo/Desktop/Mediolanum/PTF_FUNDLOOKTHROUGH MIF {anno} {mese}.xlsx')

# Per avere NAV unico: seleziona il primo valore non nullo di NAV per ogni codice e data
NAV_UNICI = PTF_MIF.groupby(['COLLECTION_DB_CODE', 'TIM_REPORTING_DATE'])['NAV'].first().reset_index()
# Per avere la somma del NAV per ongi ptf: calcola la somma dei NAV per ciascun TIM Reporting Date
NAV_TOTALE = NAV_UNICI.groupby(['TIM_REPORTING_DATE'])['NAV'].sum().reset_index()
print(NAV_TOTALE)


#NAV Obbligazionario di ogni fondo - NAV Obbligazionario ptf MIF
#Creazione PIVOT
pivot_settori = pd.pivot_table(PTF_MIF, 
                                values='HOST_FUND_WEIGHTING', 
                                index=['COLLECTION_DB_CODE', 'COD_MACRO_SECTOR', 'COD_SECTOR', 'DES_SECTOR', 'DES_PAESE'], 
                                aggfunc='sum').reset_index()

#Unisci pivot_settori con NAV_UNICI
pivot_settori = pd.merge(pivot_settori, NAV_UNICI, on='COLLECTION_DB_CODE', how='left')
#Eliminiamo la colonna della data che non ci interessa
pivot_settori = pivot_settori.drop(columns=['TIM_REPORTING_DATE'])

#Consideriamo solo i settori obbligazionari
#Selazioniamo solo le righe con COD_SECTOR che inizia con 'OB'
pivot_settori = pivot_settori[pivot_settori['COD_SECTOR'].str.startswith('OB')]

#NAV OBBLIGAZIONARIO PER SETTORE
#Aggiungi una colonna che rappresenti il prodotto tra TOTALE e NAV_Fondo
#In questo modo vediamo la quota del NAV del fondo investita in ogni settore obbligazionario
pivot_settori['NAV_Obbligazionario_Settore'] = pivot_settori['HOST_FUND_WEIGHTING'] * pivot_settori['NAV']
#NAV Obbligazionario del fondo
#E' la somma per COLLECTION_DB_CODE della colonna Prod_TOTALE_NAV per ottenere il NAV Obbligazionario del fondo
nav_obb_per_settore = pivot_settori.groupby('COLLECTION_DB_CODE')['NAV_Obbligazionario_Settore'].sum()
#Aggiungi la colonna Totale_NAV_Obbligazionario_Fondo al DataFrame pivot_settori
pivot_settori['Totale_NAV_Obbligazionario_Fondo'] = pivot_settori['COLLECTION_DB_CODE'].map(nav_obb_per_settore)

# %%
# ----------------------------------------------------
# LO STESSO PER AZ - AZALT DA ESCLUDERE E RIBASERE
# ----------------------------------------------------

#Creazione PIVOT
pivot_settori_az = pd.pivot_table(PTF_MIF, 
                                values='HOST_FUND_WEIGHTING', 
                                index=['COLLECTION_DB_CODE', 'COD_MACRO_SECTOR','COD_SECTOR', 'DES_SECTOR', 'DES_PAESE'], 
                                aggfunc='sum').reset_index()

#Aggiungi il totale per COLLECTION_DB_CODE
totale_per_collection = pivot_settori_az.groupby('COLLECTION_DB_CODE')['HOST_FUND_WEIGHTING'].sum().reset_index()
totale_per_collection['COD_SECTOR'] = 'Totale'

#Concatena il DataFrame del totale per COLLECTION_DB_CODE
pivot_settori_az = pd.concat([pivot_settori_az, totale_per_collection], ignore_index=True)
# Unisci pivot_settori_az con NAV_UNICI
pivot_settori_az = pd.merge(pivot_settori_az, NAV_UNICI, on='COLLECTION_DB_CODE', how='left')
#Rinominiamo le colonne NAV ed eliminiamo la colonna della data che non ci interessa
#pivot_settori_az.rename(columns={'NAV_y':'NAV_Totale'}, inplace=True)
#pivot_settori_az.rename(columns={'NAV_x':'NAV_Fondo'}, inplace=True)
pivot_settori_az = pivot_settori_az.drop(columns=['TIM_REPORTING_DATE'])

#Consideriamo solo i settori azionari
#Selazioniamo solo le righe con COD_SECTOR che inizia con 'AZ'
pivot_settori_az = pivot_settori_az[pivot_settori_az['COD_SECTOR'].str.startswith('AZ')]
#Calcola i totali della colonna TOTALE per ciascun COLLECTION_DB_CODE
totali_per_db_code = pivot_settori_az.groupby('COLLECTION_DB_CODE')['HOST_FUND_WEIGHTING'].sum()
#Aggiungi la colonna Totale_AZ al DataFrame pivot_settori_az
pivot_settori_az['Totale_AZ'] = pivot_settori_az['COLLECTION_DB_CODE'].map(totali_per_db_code)

#NAV AZIONARIO PER SETTORE
#Aggiungi una colonna che rappresenta il prodotto tra TOTALE e NAV_Fondo
pivot_settori_az['NAV_Azionario_Settore'] = pivot_settori_az['HOST_FUND_WEIGHTING'] * pivot_settori_az['NAV']
#NAV Azionario per settore per singolo fondo
#Calcola la somma per COLLECTION_DB_CODE della colonna Prod_TOTALE_NAV per ottenere il NAV Azionario per settore
nav_azionario_per_settore_az = pivot_settori_az.groupby('COLLECTION_DB_CODE')['NAV_Azionario_Settore'].sum()
#Aggiungi la colonna NAV_Azionario_Settore per singolo fondo al DataFrame pivot_settori_az
pivot_settori_az['Totale_NAV_Azionario_Settore'] = pivot_settori_az['COLLECTION_DB_CODE'].map(nav_azionario_per_settore_az)

#PTF AZIONARIO NORM A 100
#% azionaria per settore: NAV_Azionario_Settore / Totale_NAV_Azionario_Settore
#Aggiungi una colonna che rappresenta la normalizzazione a 100
pivot_settori_az['Normalizzazione a 100'] = (pivot_settori_az['NAV_Azionario_Settore'] / pivot_settori_az['Totale_NAV_Azionario_Settore']) * 100


#Calcola il numero totale di COD Sector (senza considerare AZALT) per ciascun COLLECTION_DB_CODE
pivot_settori_az['Num_Cod_Settore'] = pivot_settori_az.groupby('COLLECTION_DB_CODE')['COD_SECTOR'].transform('count') - 1


#RIDISTRIBUZIONE AZALT
#Droppa le righe AZALT da pivot_settori_az
pivot_settori_az = pivot_settori_az[pivot_settori_az['COD_SECTOR'] != 'AZALT']
#Calcola la somma della colonna "Normalizzazione a 100" per ciascun collection db code
somma_normalizzazione_per_db_code = pivot_settori_az.groupby('COLLECTION_DB_CODE')['Normalizzazione a 100'].sum().reset_index()
#Crea un dizionario con 'COLLECTION_DB_CODE' come chiave e 'Normalizzazione a 100' come valore
mapping_dict = somma_normalizzazione_per_db_code.set_index('COLLECTION_DB_CODE')['Normalizzazione a 100'].to_dict()
#Mappa i valori di 'Normalizzazione a 100' nel DataFrame pivot_settori_az utilizzando il dizionario creato
pivot_settori_az['NEW'] = pivot_settori_az['COLLECTION_DB_CODE'].map(mapping_dict)
#Crea la colonna del Ribilanciamento
pivot_settori_az['Ribilanciamento'] = pivot_settori_az['Normalizzazione a 100'] / pivot_settori_az['NEW']
#Check che ogni fondo torni a 100
#Calcola la somma per collection db code della colonna "Ribilanciamento"
somma_ribilanciamento_per_db_code = pivot_settori_az.groupby('COLLECTION_DB_CODE')['Ribilanciamento'].sum()
#Calcolo NAV per settore senza AZALT
pivot_settori_az['NAV_Azionario_senza_AZALT'] = pivot_settori_az['Totale_NAV_Azionario_Settore'] * pivot_settori_az['Ribilanciamento']

print(pivot_settori_az['Ribilanciamento'].sum())

# ---------------------------------------------------------
# 3️⃣ NORMALIZZAZIONE PESI AZ (→ devono tornare a 1)
# ---------------------------------------------------------

pivot_settori_az["NAV_Azionario_senza_AZALT_fondo"] = (
    pivot_settori_az
    .groupby("COLLECTION_DB_CODE")["NAV_Azionario_senza_AZALT"]
    .transform("sum")
)

NAV_tot_MIF_az = (
    pivot_settori_az[["COLLECTION_DB_CODE", "NAV_Azionario_senza_AZALT_fondo"]]
    .drop_duplicates()["NAV_Azionario_senza_AZALT_fondo"]
    .sum()
)

pivot_settori_az["NAV_Azionario_MIF"] = NAV_tot_MIF_az

pivot_settori_az["TOTALE_Ribasato"] = (
    pivot_settori_az["NAV_Azionario_senza_AZALT"] / NAV_tot_MIF_az
)

print(pivot_settori_az["TOTALE_Ribasato"].sum())


# %%
# ---------------------------------------------------------
# 3️⃣ NORMALIZZAZIONE PESI OB (→ devono tornare a 1)
# ---------------------------------------------------------

NAV_tot_MIF_obb = (
    pivot_settori[["COLLECTION_DB_CODE", "Totale_NAV_Obbligazionario_Fondo"]]
    .drop_duplicates()["Totale_NAV_Obbligazionario_Fondo"]
    .sum()
)

pivot_settori["NAV_Obbligazionario_MIF"] = NAV_tot_MIF_obb

pivot_settori["TOTALE_Ribasato"] = (
    pivot_settori["NAV_Obbligazionario_Settore"] / NAV_tot_MIF_obb
)

print(pivot_settori["TOTALE_Ribasato"].sum())


# da qui andiamo 

# RIPONDERAZIONE A 100 di ogni ongni macro categoria di Settore
# quindi per esempio, prendiamo OBGO -> fatte 100 le obbligaizoni governative andiamo a vedere il Paese in cui investiamo 
# ovviamente lo stesso suelle altre categorie:
    # quindi prendiamo ogni macro settore COD_SECOTR
    # eliminiamo le righe che non appartengono a quel COD_SECTOR
    # creiamo quindi un portafoglio per ogni COD_SECTOR
    # da qui andiamo a ricalcolarci il nuovo nav: NAV_x_COD_SECTOR
    # ribasiamo/normalizziamo il peso TOTALE_Ribasato sul nuovo nav NAV_x_COD_SECTOR
    # chiaramente teniamo tutte le info di Paese
    
    
# per fare questo stesso procedimento ma su MIF Ita Spa e Ger
# basta utilizzare il ptf_mif_diviso_per_region
# utilizzando i nav per region anziché il NAV totale, andiamo a rifare gli stessi passaggi

# 🔹 Obbligazionario

pivot_settori["NAV_macro"] = pivot_settori.groupby('COD_MACRO_SECTOR')["NAV_Obbligazionario_Settore"].transform("sum")
pivot_settori["TOTALE_Ribasato_macro"] = pivot_settori["NAV_Obbligazionario_Settore"] / pivot_settori["NAV_macro"]

#🔹 Azionario 

pivot_settori_az["NAV_macro"] = pivot_settori_az.groupby('COD_MACRO_SECTOR')["NAV_Azionario_senza_AZALT"].transform("sum")
pivot_settori_az["TOTALE_Ribasato_macro"] = pivot_settori_az["NAV_Azionario_senza_AZALT"] / pivot_settori_az["NAV_macro"]

# Controllo rapido: somma interna per ogni macro settore deve essere 1
print(pivot_settori.groupby("COD_MACRO_SECTOR")["TOTALE_Ribasato_macro"].sum())
print(pivot_settori_az.groupby("COD_MACRO_SECTOR")["TOTALE_Ribasato_macro"].sum())


# ---------------------------------------------------------
# PARAMETRI
# ---------------------------------------------------------
nome_codice = "STORIA_DEI_FONDI"
nome_codice_origine = "MIF_AZ_OB_Sector_Geo"
# share_codice = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Codici - Storia dei fondi"
share_codice = r"C:\Users\Lenovo\Desktop\Mediolanum"
# share_file = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi"
share_file = r"C:\Users\Lenovo\Desktop\Mediolanum"

# data aggiornamento (mese/anno)
data_aggiornamento = f"{mese}/{anno}"

# timestamp creazione file
data_ora_creazione = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

descrizione = (
    "PTF MIF totale diviso in parte azionaria (escluso AZALT) "
    "e obbligazionaria a livello settoriale e affondo sul paese"
)

# ---------------------------------------------------------
# INFO PER FILE OBB
# ---------------------------------------------------------
info_obb = pd.DataFrame({
    "Campo": [
        "Nome codice",
        "Nome codice origine",
        "Nome file",
        "Share codice",
        "Share file",
        "Data aggiornamento dati",
        "Data e ora creazione file",
        "Descrizione"
    ],
    "Valore": [
        nome_codice,
        nome_codice_origine,
        "PTF_MIF_AZ_OBB",
        share_codice,
        share_file,
        data_aggiornamento,
        data_ora_creazione,
        descrizione
    ]
})

# %%
# ---------------------------------------------------------
# INFO PER FILE AZ
# ---------------------------------------------------------
info_az = pd.DataFrame({
    "Campo": [
        "Nome codice",
        "Nome codice origine",
        "Nome file",
        "Share codice",
        "Share file",
        "Data aggiornamento dati",
        "Data e ora creazione file",
        "Descrizione"
    ],
    "Valore": [
        nome_codice,
        nome_codice_origine,
        "PTF_MIF_AZ_OBB",
        share_codice,
        share_file,
        data_aggiornamento,
        data_ora_creazione,
        descrizione
    ]
})


# percorso_output = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\PTF_MIF_AZ_OBB.xlsx"
percorso_output = r"C:\Users\Lenovo\Desktop\Mediolanum\PTF_MIF_AZ_OBB.xlsx"

with pd.ExcelWriter(percorso_output, engine="openpyxl") as writer:
    
    pivot_settori.to_excel(writer, sheet_name="OBB", index=False)
    pivot_settori_az.to_excel(writer, sheet_name="AZ", index=False)
    
    # puoi scegliere quale info mettere oppure unirle
    info_obb.to_excel(writer, sheet_name="INFO_OBB", index=False)
    info_az.to_excel(writer, sheet_name="INFO_AZ", index=False)

print(f"✅ File salvato: {percorso_output}")


# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==============================================================
# ==============================================================

# Storico Rating + Duration per COLLECTION_DB_CODE MIF

# ==============================================================
# ==============================================================


NOME_FAMIGLIA = "MIF"

anno_int = 2000 + int(anno)  # '26' -> 2026
mese_int = int(mese) # '02' -> 2

ANNO_INIZIO = 2021
MESE_INIZIO = 1
ANNO_FINE = anno_int
MESE_FINE = mese_int

# PERCORSO_TRACCIATI = "G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Tracciati"
PERCORSO_TRACCIATI = "C:/Users/Lenovo/Desktop/Mediolanum"
# PERCORSO_OUTPUT = "G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi"
PERCORSO_OUTPUT = "C:/Users/Lenovo/Desktop/Mediolanum"

FILE_OUTPUT = f"{PERCORSO_OUTPUT}/STORICO_RATING_E_DURATION_{NOME_FAMIGLIA}.xlsx"


# FUNZIONI


def importa_tracciati(anno, mese):
    anno_str = f"{anno % 100:02d}"
    mese_str = f"{mese:02d}"

    file_tracciato = f"{PERCORSO_TRACCIATI}/Tracciato {NOME_FAMIGLIA} {anno_str} {mese_str}.xlsx"

    if not os.path.exists(file_tracciato):
        print(f"⚠️ File mancante {anno}-{mese:02d}")
        return None, None

    rating = pd.read_excel(file_tracciato, sheet_name="RATING")
    duration = pd.read_excel(file_tracciato, sheet_name="DURATION")

    return rating, duration


def genera_mesi():
    mesi = []
    a, m = ANNO_INIZIO, MESE_INIZIO

    while (a < ANNO_FINE) or (a == ANNO_FINE and m <= MESE_FINE):
        mesi.append((a, m))
        m += 1
        if m == 13:
            m = 1
            a += 1
    return mesi

# %%
# ------------------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------------------

storico = []

for anno, mese in genera_mesi():

    print(f"📥 Elaborazione {anno}-{mese:02d}")
    rating, duration = importa_tracciati(anno, mese)

    if rating is None:
        continue

    # -------------------------
    # RATING → %Exposure
    # -------------------------
    rating_agg = (
        rating
        .groupby(
            ["COLLECTION_DB_CODE", "TIM_REPORTING_DATE", "COD_RATING"],
            as_index=False
        )["PESO"]
        .sum()
    )

    tot_fondo = (
        rating_agg
        .groupby(["COLLECTION_DB_CODE", "TIM_REPORTING_DATE"], as_index=False)["PESO"]
        .sum()
        .rename(columns={"PESO": "TOT_PESO"})
    )

    rating_agg = rating_agg.merge(
        tot_fondo,
        on=["COLLECTION_DB_CODE", "TIM_REPORTING_DATE"],
        how="left"
    )

    rating_agg["PERC_EXPOSURE"] = (
        rating_agg["PESO"] / rating_agg["TOT_PESO"] * 100
    )

    rating_agg = rating_agg.drop(columns=["PESO", "TOT_PESO"])

    # -------------------------
    # DURATION
    # -------------------------
    duration_clean = (
        duration
        .groupby(["COLLECTION_DB_CODE", "TIM_REPORTING_DATE"], as_index=False)
        .agg({"DURATION": "mean"})
    )

    # -------------------------
    # MERGE FINALE
    # -------------------------
    finale = rating_agg.merge(
        duration_clean,
        on=["COLLECTION_DB_CODE", "TIM_REPORTING_DATE"],
        how="left"
    )

    storico.append(finale)


# ------------------------------------------------------------------------------
# CONCAT FINALE
# ------------------------------------------------------------------------------
if len(storico) == 0:
    raise ValueError("❌ Nessun dato elaborato: storico vuoto")

storico = pd.concat(storico, ignore_index=True)


# %%
# ------------------------------------------------------------------------------
# FORMATTAZIONI POWER BI
# ------------------------------------------------------------------------------

# Data vera (asse temporale)
storico["TIM_REPORTING_DATE"] = pd.to_datetime(
    storico["TIM_REPORTING_DATE"]
)

# rating senza %
storico["PERC_EXPOSURE"] = storico["PERC_EXPOSURE"] / 100


# Ordine rating
ordine_rating = {
    "AAA": 1,
    "AA": 2,
    "A": 3,
    "BBB": 4,
    "BB": 5,
    "B": 6,
    "CCC": 7,
    '<=CCC': 8,
    'D': 9,
    'NR': 10
}

storico["ORDINE_RATING"] = storico["COD_RATING"].map(ordine_rating)

# Rinomina colonne
storico = storico.rename(columns={
    "COLLECTION_DB_CODE": "FONDO",
    "COD_RATING": "RATING",
    "PESO": "EXPOSURE",
    "DURATION": "DURATION_YEARS"
})


# ------------------------------------------------------------------------------
# FORMATO WIDE
# ------------------------------------------------------------------------------

# Pivot per avere una colonna per ogni rating
storico_wide = storico.pivot_table(
    index=["FONDO", "TIM_REPORTING_DATE", "DURATION_YEARS"],
    columns="RATING",
    values="PERC_EXPOSURE",
    fill_value=0
).reset_index()


# %%
# --------------------------------------------------------------------------
# RATING MEDIO STORICO
# --------------------------------------------------------------------------

# Usa i dati long (prima del pivot)
rating_medio = storico.copy()

# Escludi NR
rating_medio = rating_medio[rating_medio["RATING"] != "NR"]

# Riponderazione a 1 per fondo/data
rating_medio["PESO_REBASE"] = (
    rating_medio
    .groupby(["FONDO", "TIM_REPORTING_DATE"])["PERC_EXPOSURE"]
    .transform("sum")
)

rating_medio["PESO_NORM"] = (
    rating_medio["PERC_EXPOSURE"] / rating_medio["PESO_REBASE"]
)

# Scala rating (stessa che usi tu)
scala_rating = {
    "AAA": 7,
    "AA": 6,
    "A": 5,
    "BBB": 4,
    "BB": 3,
    "B": 2,
    "<=CCC": 1,
    "D": 0
}

rating_medio["SCALA"] = rating_medio["RATING"].map(scala_rating)

# Passaggio ponderato
rating_medio["PASSAGGIO"] = (
    rating_medio["PESO_NORM"] * rating_medio["SCALA"]
)

# Rating medio numerico
rating_medio_num = (
    rating_medio
    .groupby(["FONDO", "TIM_REPORTING_DATE"], as_index=False)
    .agg({"PASSAGGIO": "sum"})
    .rename(columns={"PASSAGGIO": "RATING_MEDIO_NUM"})
)

# Arrotondamento (stesso stile)
rating_medio_num["RATING_MEDIO_NUM"] = (
    rating_medio_num["RATING_MEDIO_NUM"].round().astype(int)
)

# Inversione scala → rating
inv_scala_rating = {v: k for k, v in scala_rating.items()}

rating_medio_num["RATING_MEDIO"] = (
    rating_medio_num["RATING_MEDIO_NUM"].map(inv_scala_rating)
)

# Ordina per sicurezza
rating_medio_num = rating_medio_num.sort_values(
    ["FONDO", "TIM_REPORTING_DATE"]
)


# ------------------------------------------------------------------------------
# OUTPUT
# ------------------------------------------------------------------------------

nome_codice = "STORIA_DEI_FONDI"
nome_codice_di_origine = "Storico_rating_e_duration_MIF_MGF_MGE"
cartella_codice = os.getcwd()
os.makedirs(cartella_codice, exist_ok=True)
data_ora_run = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

info_run = pd.DataFrame({
    "Campo": [
        "Nome codice",
        "Nome codice di origine",
        "Famiglia",
        "Cartella codice",
        "Data e ora run",
        "Periodo inizio",
        "Periodo fine"
    ],
    "Valore": [
        nome_codice,
        nome_codice_di_origine,
        NOME_FAMIGLIA,
        cartella_codice,
        data_ora_run,
        f"{ANNO_INIZIO}-{MESE_INIZIO:02d}",
        f"{ANNO_FINE}-{MESE_FINE:02d}"
    ]
})


# %%
# -------------------------------------------------------------------------------
# AGGIUNGI DENOMINAZIONE FONDO
# -------------------------------------------------------------------------------

# Ricava anno/mese ultima data dei tracciati già importati
ultima_data = storico["TIM_REPORTING_DATE"].max()
anno_ultima = ultima_data.year % 100  # formato YY
mese_ultima = ultima_data.month

# Percorso file PTF corrispondente
# file_ptf_ultima = f"G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Portafoglio/PTF_FUNDLOOKTHROUGH {NOME_FAMIGLIA} {anno_ultima:02d} {mese_ultima:02d}.xlsx"
file_ptf_ultima = f"C:/Users/Lenovo/Desktop/Mediolanum/PTF_FUNDLOOKTHROUGH {NOME_FAMIGLIA} {anno_ultima:02d} {mese_ultima:02d}.xlsx"

if os.path.exists(file_ptf_ultima):
    ptf_ultima = pd.read_excel(file_ptf_ultima)
    ptf_ultima = ptf_ultima[["COLLECTION_DB_CODE", "DENOMINAZIONE"]].drop_duplicates()
else:
    ptf_ultima = pd.DataFrame()
    print(f"⚠️ File PTF non trovato per {anno_ultima}-{mese_ultima:02d}")

# ---- Merge con storico long ----
if not ptf_ultima.empty:
    storico = storico.merge(
        ptf_ultima,
        left_on="FONDO",
        right_on="COLLECTION_DB_CODE",
        how="left"
    ).drop(columns=["COLLECTION_DB_CODE"])
else:
    storico["DENOMINAZIONE"] = np.nan

# ---- Aggiorna storico wide includendo DENOMINAZIONE ----
storico_wide = storico.pivot_table(
    index=["FONDO", "DENOMINAZIONE", "TIM_REPORTING_DATE", "DURATION_YEARS"],
    columns="RATING",
    values="PERC_EXPOSURE",
    fill_value=0
).reset_index()

# ---- Aggiorna rating medio storico includendo DENOMINAZIONE ----
if not ptf_ultima.empty:
    rating_medio_num = rating_medio_num.merge(
        ptf_ultima,
        left_on="FONDO",
        right_on="COLLECTION_DB_CODE",
        how="left"
    ).drop(columns=["COLLECTION_DB_CODE"])
else:
    rating_medio_num["DENOMINAZIONE"] = np.nan



# -------------------------------------------------------------------
# SOVRASCRIVI DURATION PER MARZO E APRILE 2025
# -------------------------------------------------------------------
# file_duration_corretto = r"G:/Analisi e Performance Prodotti/Fact Sheet New 2020/YTM MGF/Duration_MIFL.xlsx"
file_duration_corretto = r"C:/Users/Lenovo/Desktop/Mediolanum/Duration_MIFL.xlsx"

if os.path.exists(file_duration_corretto):
    duration_corr = pd.read_excel(file_duration_corretto)
    
    # Assumiamo che abbia colonne: COLLECTION_DB_CODE, TIM_REPORTING_DATE, DURATION
    duration_corr["TIM_REPORTING_DATE"] = pd.to_datetime(duration_corr["TIM_REPORTING_DATE"])
    
    # Consideriamo solo anno e mese per l'aggancio
    duration_corr["ANNO_MESE"] = duration_corr["TIM_REPORTING_DATE"].dt.to_period("M")
    storico["ANNO_MESE"] = storico["TIM_REPORTING_DATE"].dt.to_period("M")
    
    # Merge per aggiungere la colonna DURATION dal file corretto
    storico = storico.merge(
        duration_corr[["COLLECTION_DB_CODE", "ANNO_MESE", "DURATION"]],
        left_on=["FONDO", "ANNO_MESE"],
        right_on=["COLLECTION_DB_CODE", "ANNO_MESE"],
        how="left"
    )
    
    # Sovrascrivi DURATION_YEARS solo se DURATION ha valore
    storico["DURATION_YEARS"] = storico["DURATION"].combine_first(storico["DURATION_YEARS"])
    
    # Pulisci colonne ausiliarie
    storico = storico.drop(columns=["COLLECTION_DB_CODE", "DURATION", "ANNO_MESE"])
    
    print("✅ DURATION aggiornata con file Duration_MIFL per i mesi presenti")
else:
    print(f"⚠️ File {file_duration_corretto} non trovato: duration non aggiornata")


# CHECK

# Filtra per DB2017
storico_db2017 = storico[storico["FONDO"] == "DB2017"].copy()  #

# Filtra per Marzo e Aprile 2025 usando Period
storico_mar_apr = storico_db2017[
    storico_db2017["TIM_REPORTING_DATE"].dt.to_period("M").isin([
        pd.Period("2025-03", freq="M"),
        pd.Period("2025-04", freq="M")
    ])
]

# Mostra solo colonne utili
print(storico_mar_apr[["FONDO", "DENOMINAZIONE", "TIM_REPORTING_DATE", "DURATION_YEARS"]])

# %%
# ------------------------------------------------------------------------------
# EXPORT EXCEL CON PIÙ SHEET
# ------------------------------------------------------------------------------
with pd.ExcelWriter(FILE_OUTPUT, engine="openpyxl") as writer:
    storico.to_excel(
        writer,
        sheet_name="STORICO_RATING_DURATION",
        index=False
    )
    info_run.to_excel(
        writer,
        sheet_name="INFO_RUN",
        index=False
    )
    storico_wide.to_excel(
        writer,
        sheet_name="FORMATO_WIDE",
        index=False
        )
    rating_medio_num.to_excel(
        writer,
        sheet_name="RATING_MEDIO_STORICO",
        index=False
    )
    

print("✅ File Excel creato con sheet dati + info run")
print(FILE_OUTPUT)


# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# =================================================
# =================================================

# CONFRONTO MIF MGF ED MGE ALL'ULTIMA RILEVAZIONE 

# 🌍💼 CALCOLO RATING OBBLIGAZIONARIO MIF / MGF / MGE

# ==========================================================

# ==========================================================
# 📥 1️⃣ CARICAMENTO TRACCIATI RATING + FAMIGLIA
# ==========================================================
anno_int = 2000 + int(anno)  # '26' -> 2026
mese_int = int(mese) # '02' -> 2

ANNO_FINE = anno_int
MESE_FINE = mese_int

anno_str = f"{ANNO_FINE % 100:02d}"
mese_str = f"{MESE_FINE:02d}"


# 📂 PATH

# BASE_PATH = "G:/Analisi e Performance Prodotti/Fact Sheet New 2020"
BASE_PATH = "C:/Users/Lenovo/Desktop/Mediolanum"


# 📊 OUTPUT

df_finale_list = []


# 🏦 LOOP FAMIGLIE

for famiglia in ["MIF", "MGF", "MGE"]:

    print(f"➡️ Elaboro {famiglia} - {ANNO_FINE}-{mese_str}")

    try:
        # ==========================================================
        # 📥 LOAD FILE
        # ==========================================================
        file_rating = f"{BASE_PATH}/Tracciati/Tracciato {famiglia} {anno_str} {mese_str}.xlsx"
        file_asset = f"{BASE_PATH}/Tracciati/Tracciato {famiglia} {anno_str} {mese_str}.xlsx"
        file_ptf = f"{BASE_PATH}/Portafoglio/PTF_FUNDLOOKTHROUGH {famiglia} {anno_str} {mese_str}.xlsx"

        if not (os.path.exists(file_rating) and os.path.exists(file_ptf)):
            print(f"⚠️ File mancanti per {famiglia}")
            continue

        rating = pd.read_excel(file_rating, sheet_name="RATING")
        asset = pd.read_excel(file_asset, sheet_name="ASSET_5D", header=1)
        ptf = pd.read_excel(file_ptf)

        # ==========================================================
        # 🏷️ AGGIUNTA FAMIGLIA
        # ==========================================================
        rating["FAMIGLIA"] = famiglia

        # ==========================================================
        # 🔗 MERGE OBB (da ASSET_5D)
        # ==========================================================
        asset_obb = asset[["COLLECTION_DB_CODE", "OBB"]]

        rating = rating.merge(
            asset_obb,
            on="COLLECTION_DB_CODE",
            how="left"
        )

        # ==========================================================
        # 🔗 MERGE NAV (da PTF)
        # ==========================================================
        ptf_nav = (
            ptf[["COLLECTION_DB_CODE", "NAV"]]
            .drop_duplicates(subset=["COLLECTION_DB_CODE"])
        )

        rating = rating.merge(
            ptf_nav,
            on="COLLECTION_DB_CODE",
            how="left"
        )

        # ==========================================================
        # 💰 CALCOLO HOLDING MARKET VALUE OBB
        # ==========================================================
        rating["NAV_OBB"] = rating["OBB"] * rating["NAV"]
        rating["HOLDING_MARKET_VALUE_OBB"] = rating["PESO"] * rating["NAV_OBB"]

        # ==========================================================
        # ⚖️ CALCOLO TOTALE_RIBASATO (per famiglia)
        # ==========================================================
        totale_famiglia = rating.groupby("FAMIGLIA")["HOLDING_MARKET_VALUE_OBB"].transform("sum")

        rating["TOTALE_RIBASATO"] = rating["HOLDING_MARKET_VALUE_OBB"] / totale_famiglia

        # ==========================================================
        # 📦 KEEP SOLO COLONNE UTILI
        # ==========================================================
        rating_final = rating[[
            "TIM_REPORTING_DATE",
            "COLLECTION_DB_CODE",
            "COD_RATING",
            "TIPO_RATING",
            "PESO",
            "FAMIGLIA",
            "OBB",
            "NAV",
            "HOLDING_MARKET_VALUE_OBB",
            "TOTALE_RIBASATO"
        ]]

        df_finale_list.append(rating_final)

    except Exception as e:
        print(f"❌ Errore su {famiglia}: {e}")
        continue

# %%
# ==========================================================
# 📦 DATAFRAME FINALE
# ==========================================================
df_finale = pd.concat(df_finale_list, ignore_index=True)

# ==========================================================
# ✅ CHECK
# ==========================================================
print("\n✅ Check TOTALE_RIBASATO per famiglia (deve fare 1):")
print(
    df_finale.groupby("FAMIGLIA")["TOTALE_RIBASATO"]
    .sum()
    .reset_index()
)

rating_totale = df_finale.copy()

# ------------------------------------------------
# Parametri file
# ------------------------------------------------
nome_codice = "STORIA_DEI_FONDI"
nome_codice_di_origine = "Storico_rating_e_duration_MIF_MGF_MGE"
# share_codice = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Codici - Storia dei fondi"
share_codice = r"C:\Users\Lenovo\Desktop\Mediolanum"
file_output = "confronto_rating_mif_mgf_mge.xlsx"
# share_output = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi"
share_output = r"C:\Users\Lenovo\Desktop\Mediolanum"

# Creiamo le cartelle se non esistono
os.makedirs(share_codice, exist_ok=True)
os.makedirs(share_output, exist_ok=True)

# Percorso completo del file di output
file_excel_output = os.path.join(share_output, file_output)

# ------------------------------------------------
# Data e ora esecuzione
# ------------------------------------------------
data_ora_run = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ------------------------------------------------
# Creiamo il dataframe info_run "sicuro"
# ------------------------------------------------
info_run = pd.DataFrame({
    "Nome codice": [nome_codice],
    "nome_codice_di_origine": [nome_codice_di_origine],
    "Percorso codice": [share_codice],
    "File di output": [file_output],
    "Share output": [share_output],
    "Data e ora esecuzione": [data_ora_run]
})

# ------------------------------------------------
# Salvataggio Excel con più sheet
# ------------------------------------------------
with pd.ExcelWriter(file_excel_output, engine="openpyxl") as writer:
    rating_totale.to_excel(writer, sheet_name="CONFRONTO_RATING", index=False)
    info_run.to_excel(writer, sheet_name="INFO_RUN", index=False)

print(f"✅ File Excel salvato correttamente: {file_excel_output}")


# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣


# ===========================================================
# ===========================================================
# STORICO RATING MEDIO OBB MIF MGF ED MGE
# ===========================================================
# ===========================================================


# PER IL RATING OBB DEI FONDI
# L'HOST_FUND_WEIGHTING_RATING è già calcolato sulla parte OBB del fondo
# Quindi abbiamo due possibilità:
    
    # PTF:
        # livello di fondo: 
            # facciamo una pivot sul PTF_MIF dove filtriamo per il COLLECTION_DB_CODE di interesse, il COD_MACRO_ASSET = OB
            # mettiamo poi in riga i COD_RATING_DB_ASSET
            # e due volte nei valori HOLDING_MARKET_VALUE: 
                # (1) somma dei valori HOLDING_MARKET_VALUE
                # (2) somma dei valori HOLDING_MARKET_VALUE ma mostra valori come % del totale complessivo 
            # => in questo modo abbiamo il NAV_OBB del fondo (somma di HOLDING_MARKET_VALUE) e poi il ptf 
        # livello di famiglia: andiamo a fare sostanzialmente la stessa cosa ma senza filtrare per COLLECTION_DB_CODE
            # quindi filtriamo solo per COD_MACRO_ASSET = OB
            # mettiamo poi in riga i COD_RATING_DB_ASSET
            # e due volte nei valori HOLDING_MARKET_VALUE:
                # (1) somma dei valori HOLDING_MARKET_VALUE
                # (2) somma dei valori HOLDING_MARKET_VALUE ma mostra valori come % del totale complessivo 


# USIAMO METODO TRACCIATO
    
    # TRACCIATO: a livello di tacciato se vogliamo è più preciso rispetto al ptf perché abbiamo anche il dettaglio dei fondi di cui facciamo noi il caricamento 
        # prendiamo in traccaito da ASSET_5D la diversificaizone nelle macro asset class (colonna D, riga 2 -.> intestazione OBB)
        # poi prendiamo da ptf il NAV del fondo 
        # andiamo a considerare la parte OB del fondo: HOLDING_MARKET_VALUE_OB =  OBB * NAV 
        # a questo punto basta fare una pivot di nuovo: 
            # COD_RATING nelle righe
            # somma di HOLDING_MARKET_VALUE_OBB
            # somma di HOLDING_MARKET_VALUE_OBB -> mostra valori come % del totale complessivo 


# per il rating medio: 
    # assegniamo un punteggio a partire da 1 (più basso) per la D fino alla AAA:
        # <=CCC	2
        # A 6 
        # AA	 7
        # AAA 8
        # B	3
        # BB 4
        # BBB 5
        # D	1
    # moltiplichiamo il valore per il peso -> sommiamo ed otteniamo un puntggio totale
    # in base ad una scala, assegnamo poi il rating medio a MIF per il ptf a quella data 
    # def rating_uale(val):
    #    if val > 7.5: return 'AAA'
    #    elif val > 6.5: return 'AA'
    #    elif val > 5.5: return 'A'
    #    elif val > 4.5: return 'BBB'
    #    elif val > 3.5: return 'BB'
    #    elif val > 2.5: return 'B'
    #    elif val > 1.5: return '<=CCC'
    #    else: return 'D'


# %%
# TODO

r"""

# QUI ANDIAMO A CREARE LO STORICO


import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings("ignore")

# ==========================================================
# 📅 PARAMETRI PERIODO
# ==========================================================
ANNO_INIZIO = 2021
MESE_INIZIO = 12

anno = '26'
mese = '02'

anno_int = 2000 + int(anno)
mese_int = int(mese)

ANNO_FINE = anno_int
MESE_FINE = mese_int

# Lista mesi (YYYY-MM)
periodi = []

current_year = ANNO_INIZIO
current_month = MESE_INIZIO

while (current_year < ANNO_FINE) or (current_year == ANNO_FINE and current_month <= MESE_FINE):
    periodi.append((current_year, current_month))
    
    current_month += 1
    if current_month > 12:
        current_month = 1
        current_year += 1


# ==========================================================
# 🧾 OUTPUT FINALE
# ==========================================================
risultati = []

# ==========================================================
# 🏦 LOOP FAMIGLIE
# ==========================================================
famiglie = ["MIF", "MGF", "MGE"]

for famiglia in famiglie:

    for year, month in periodi:

        anno_str = str(year)[2:]
        mese_str = f"{month:02d}"

        print(f"➡️ Sto elaborando: {famiglia} - {year}-{mese_str}")

        try:
            # ==========================================================
            # 📥 LOAD FILE
            # ==========================================================
            # base_path = 'G:/Analisi e Performance Prodotti/Fact Sheet New 2020'
            base_path = 'C:/Users/Lenovo/Desktop/Mediolanum'

            PTF = pd.read_excel(
                f'{base_path}/Portafoglio/PTF_FUNDLOOKTHROUGH {famiglia} {anno_str} {mese_str}.xlsx'
            )

            tracciato_rating = pd.read_excel(
                f'{base_path}/Tracciati/Tracciato {famiglia} {anno_str} {mese_str}.xlsx',
                sheet_name='RATING'
            )

            tracciato_asset = pd.read_excel(
                f'{base_path}/Tracciati/Tracciato {famiglia} {anno_str} {mese_str}.xlsx',
                sheet_name='ASSET_5D',
                header=1
            )

            # ==========================================================
            # 🔧 CLEAN
            # ==========================================================
            rating = tracciato_rating.copy()
            rating["FAMIGLIA"] = famiglia

            # ==========================================================
            # 🔗 MERGE OBB
            # ==========================================================
            asset_obb = tracciato_asset[[
                "COLLECTION_DB_CODE",
                "OBB"
            ]]

            rating = rating.merge(
                asset_obb,
                on="COLLECTION_DB_CODE",
                how="left"
            )

            # ==========================================================
            # 🏦 NAV
            # ==========================================================
            ptf_nav = (
                PTF[["COLLECTION_DB_CODE", "NAV"]]
                .drop_duplicates(subset=["COLLECTION_DB_CODE"])
            )

            rating = rating.merge(
                ptf_nav,
                on="COLLECTION_DB_CODE",
                how="left"
            )

            # ==========================================================
            # 💰 HOLDING VALUE
            # ==========================================================
            rating["NAV_OBB"] = rating["OBB"] * rating["NAV"]
            rating["HOLDING_MARKET_VALUE_OBB"] = rating["PESO"] * rating["NAV_OBB"]

            # ==========================================================
            # 📊 PIVOT
            # ==========================================================
            pivot = (
                rating[rating["COD_RATING"] != "NR"]
                .groupby(["TIM_REPORTING_DATE", "FAMIGLIA", "COD_RATING"])["HOLDING_MARKET_VALUE_OBB"]
                .sum()
                .reset_index()
            )

            total = pivot["HOLDING_MARKET_VALUE_OBB"].sum()
            pivot["PERCENTUALE"] = pivot["HOLDING_MARKET_VALUE_OBB"] / total

            # ==========================================================
            # ⭐ RATING SCORE
            # ==========================================================
            rating_map = {
                "D": 1,
                "<=CCC": 2,
                "B": 3,
                "BB": 4,
                "BBB": 5,
                "A": 6,
                "AA": 7,
                "AAA": 8
            }

            pivot["RATING_SCORE"] = pivot["COD_RATING"].map(rating_map)
            pivot["passaggio_intermedio"] = pivot["RATING_SCORE"] * pivot["PERCENTUALE"]

            totale_score = pivot["passaggio_intermedio"].sum()

            # ==========================================================
            # 🎯 RATING MEDIO
            # ==========================================================
            def rating_uale(val):
                if val > 7.5: return 'AAA'
                elif val > 6.5: return 'AA'
                elif val > 5.5: return 'A'
                elif val > 4.5: return 'BBB'
                elif val > 3.5: return 'BB'
                elif val > 2.5: return 'B'
                elif val > 1.5: return '<=CCC'
                else: return 'D'

            rating_medio = rating_uale(totale_score)

            # ==========================================================
            # 🧾 OUTPUT RIGA
            # ==========================================================
            risultati.append({
                "TIM_REPORTING_DATE": f"{year}-{mese_str}-01",
                "FAMIGLIA": famiglia,
                "RATING_MEDIO": rating_medio,
                "SCORE_MEDIO": totale_score
            })

        except Exception as e:
            print(f"Errore su {famiglia} {year}-{mese_str}: {e}")
            continue


# ==========================================================
# 📦 DATAFRAME FINALE
# ==========================================================
df_finale = pd.DataFrame(risultati)

df_finale = df_finale.sort_values(["TIM_REPORTING_DATE", "FAMIGLIA"]).reset_index(drop=True)


output_path = r"C:\Users\alessandra.morigi\OneDrive - Banca Mediolanum SPA\Desktop\rating_medio.xlsx"

df_finale.to_excel(output_path, index=False)

print("File salvato in:", output_path)


# --- Percorso di salvataggio ---
# PERCORSO_FILE = r'G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Storico_rating_MEDIO_OBB_MIF_MGF_MGE.xlsx'
PERCORSO_FILE = r'C:\Users\Lenovo\Desktop\Mediolanum\Storico_rating_MEDIO_OBB_MIF_MGF_MGE.xlsx'

# --- Scrittura su Excel con più fogli ---
with pd.ExcelWriter(PERCORSO_FILE, engine='openpyxl') as writer:
    
    # foglio 1: rating per famiglia
    storico_famiglia.to_excel(writer, sheet_name='Rating_Famiglia', index=False)
    
    # foglio 4: info dati
    info = pd.DataFrame({
        'Codice': ['STORIA_DEI_FONDI'],
        'Codice di origine': ['Rating_medio'],
        'Nome_File': ['Storico_rating_MEDIO_OBB_MIF_MGF_MGE'],
        'Descrizione': ['Storico rating medio OBBLIGAZIONARIO per MIF, MGF e MGE.'],
        # 'Fonte': ['G:\\Analisi e Performance Prodotti\\PowerBi MAA\\Storia dei fondi\\Codici - Storia dei fondi']
        'Fonte': ['C:\\Users\\Lenovo\\Desktop\\Mediolanum']
    })
    info.to_excel(writer, sheet_name='Info_Dati', index=False)

print(f"✅ File Excel salvato correttamente in: {PERCORSO_FILE}")


"""


# %%
# ==========================================================
# ⚙️ PARAMETRI
# ==========================================================
anno = anno
mese = mese

OVERWRITE = True  # 👉 True = sovrascrive se esiste, False = aggiunge solo se manca

anno_int = 2000 + int(anno)
mese_int = int(mese)

periodi = [(anno_int, mese_int)]

famiglie = ["MIF", "MGF", "MGE"]

# ==========================================================
# 📂 PERCORSO FILE STORICO
# ==========================================================
# PERCORSO_FILE = r'G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Storico_rating_MEDIO_OBB_MIF_MGF_MGE.xlsx'
PERCORSO_FILE = r'C:\Users\Lenovo\Desktop\Mediolanum\Storico_rating_MEDIO_OBB_MIF_MGF_MGE.xlsx'

# ==========================================================
# 📥 CARICAMENTO STORICO ESISTENTE
# ==========================================================
try:
    storico_esistente = pd.read_excel(PERCORSO_FILE, sheet_name='Rating_Famiglia')
except:
    storico_esistente = pd.DataFrame(columns=[
        "TIM_REPORTING_DATE", "FAMIGLIA", "RATING_MEDIO", "SCORE_MEDIO"
    ])

# ==========================================================
# 🧾 OUTPUT NUOVO
# ==========================================================
risultati = []

# %%
# ==========================================================
# 🔁 LOOP
# ==========================================================
for famiglia in famiglie:

    for year, month in periodi:

        anno_str = str(year)[2:]
        mese_str = f"{month:02d}"

        print(f"➡️ Sto elaborando: {famiglia} - {year}-{mese_str}")

        try:
            # ==========================================================
            # 📥 LOAD FILE
            # ==========================================================
            # base_path = 'G:/Analisi e Performance Prodotti/Fact Sheet New 2020'
            base_path = 'C:/Users/Lenovo/Desktop/Mediolanum'

            PTF = pd.read_excel(
                f'{base_path}/Portafoglio/PTF_FUNDLOOKTHROUGH {famiglia} {anno_str} {mese_str}.xlsx'
            )

            tracciato_rating = pd.read_excel(
                f'{base_path}/Tracciati/Tracciato {famiglia} {anno_str} {mese_str}.xlsx',
                sheet_name='RATING'
            )

            tracciato_asset = pd.read_excel(
                f'{base_path}/Tracciati/Tracciato {famiglia} {anno_str} {mese_str}.xlsx',
                sheet_name='ASSET_5D',
                header=1
            )

            # ==========================================================
            # 🔧 CLEAN
            # ==========================================================
            rating = tracciato_rating.copy()
            rating["FAMIGLIA"] = famiglia

            # ==========================================================
            # 🔗 MERGE OBB
            # ==========================================================
            asset_obb = tracciato_asset[[
                "COLLECTION_DB_CODE",
                "OBB"
            ]]

            rating = rating.merge(
                asset_obb,
                on="COLLECTION_DB_CODE",
                how="left"
            )

            # ==========================================================
            # 🏦 NAV
            # ==========================================================
            ptf_nav = (
                PTF[["COLLECTION_DB_CODE", "NAV"]]
                .drop_duplicates(subset=["COLLECTION_DB_CODE"])
            )

            rating = rating.merge(
                ptf_nav,
                on="COLLECTION_DB_CODE",
                how="left"
            )

            # ==========================================================
            # 💰 HOLDING VALUE
            # ==========================================================
            rating["NAV_OBB"] = rating["OBB"] * rating["NAV"]
            rating["HOLDING_MARKET_VALUE_OBB"] = rating["PESO"] * rating["NAV_OBB"]

            # ==========================================================
            # 📊 PIVOT
            # ==========================================================
            pivot = (
                rating[rating["COD_RATING"] != "NR"]
                .groupby(["TIM_REPORTING_DATE", "FAMIGLIA", "COD_RATING"])["HOLDING_MARKET_VALUE_OBB"]
                .sum()
                .reset_index()
            )

            total = pivot["HOLDING_MARKET_VALUE_OBB"].sum()
            pivot["PERCENTUALE"] = pivot["HOLDING_MARKET_VALUE_OBB"] / total

            # ==========================================================
            # ⭐ RATING SCORE
            # ==========================================================
            rating_map = {
                "D": 1,
                "<=CCC": 2,
                "B": 3,
                "BB": 4,
                "BBB": 5,
                "A": 6,
                "AA": 7,
                "AAA": 8
            }

            pivot["RATING_SCORE"] = pivot["COD_RATING"].map(rating_map)
            pivot["passaggio_intermedio"] = pivot["RATING_SCORE"] * pivot["PERCENTUALE"]

            totale_score = pivot["passaggio_intermedio"].sum()

            # ==========================================================
            # 🎯 RATING MEDIO
            # ==========================================================
            def rating_uale(val):
                if val > 7.5: return 'AAA'
                elif val > 6.5: return 'AA'
                elif val > 5.5: return 'A'
                elif val > 4.5: return 'BBB'
                elif val > 3.5: return 'BB'
                elif val > 2.5: return 'B'
                elif val > 1.5: return '<=CCC'
                else: return 'D'

            rating_medio = rating_uale(totale_score)

            risultati.append({
                "TIM_REPORTING_DATE": f"{year}-{mese_str}-01",
                "FAMIGLIA": famiglia,
                "RATING_MEDIO": rating_medio,
                "SCORE_MEDIO": totale_score
            })

        except Exception as e:
            print(f"Errore su {famiglia} {year}-{mese_str}: {e}")
            continue

# %%
# ==========================================================
# 📦 DATAFRAME NUOVO
# ==========================================================
df_finale = pd.DataFrame(risultati)

# ==========================================================
# 🔄 AGGIORNAMENTO STORICO
# ==========================================================
df_finale["TIM_REPORTING_DATE"] = pd.to_datetime(df_finale["TIM_REPORTING_DATE"])
storico_esistente["TIM_REPORTING_DATE"] = pd.to_datetime(storico_esistente["TIM_REPORTING_DATE"])

if OVERWRITE:
    storico_esistente = storico_esistente[
        ~storico_esistente.set_index(["TIM_REPORTING_DATE", "FAMIGLIA"]).index.isin(
            df_finale.set_index(["TIM_REPORTING_DATE", "FAMIGLIA"]).index
        )
    ]
    storico_aggiornato = pd.concat([storico_esistente, df_finale], ignore_index=True)
else:
    nuovi = df_finale[
        ~df_finale.set_index(["TIM_REPORTING_DATE", "FAMIGLIA"]).index.isin(
            storico_esistente.set_index(["TIM_REPORTING_DATE", "FAMIGLIA"]).index
        )
    ]
    storico_aggiornato = pd.concat([storico_esistente, nuovi], ignore_index=True)

# ==========================================================
# 📊 ORDINAMENTO
# ==========================================================
storico_aggiornato = storico_aggiornato.sort_values(
    ["TIM_REPORTING_DATE", "FAMIGLIA"]
).reset_index(drop=True)

# ==========================================================
# 💾 SALVATAGGIO
# ==========================================================
with pd.ExcelWriter(PERCORSO_FILE, engine='openpyxl') as writer:

    storico_aggiornato.to_excel(writer, sheet_name='Rating_Famiglia', index=False)

    info = pd.DataFrame({
        'Codice': ['STORIA_DEI_FONDI'],
        'Codice di origine': ['Rating_medio'],
        'Nome_File': ['Storico_rating_MEDIO_OBB_MIF_MGF_MGE'],
        'Descrizione': ['Storico rating medio OBBLIGAZIONARIO per MIF, MGF e MGE.'],
        # 'Fonte': ['G:\\Analisi e Performance Prodotti\\PowerBi MAA\\Storia dei fondi\\Codici - Storia dei fondi']
        'Fonte': ['C:\\Users\\Lenovo\\Desktop\\Mediolanum']
    })

    info.to_excel(writer, sheet_name='Info_Dati', index=False)

print(f"✅ File aggiornato correttamente in: {PERCORSO_FILE}")
print(f"📊 Righe totali storico: {len(storico_aggiornato)}")
    



# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==============================================================
# ==============================================================

# RATING MIF diviso per Region/Paese ex FoHF

# ==============================================================
# ==============================================================


# 📥 IMPORT FUNZIONE REGIONI

from ptf_mif_diviso_per_region import ptf_mif_region_fn

TRACCIATO_MIF_ASSET, PTF_MIF, totale = ptf_mif_region_fn(
    anno=anno,
    mese=mese,
    giorno=giorno,
    giorno_lavorativo=giorno_lavorativo
)

# ==========================================================
# 🧹 CLEAN PTF
# ==========================================================
PTF_MIF = PTF_MIF[PTF_MIF['COLLECTION_DB_CODE'] != 'DBFoHF']

# ==========================================================
# 📥 LOAD TRACCIATI
# ==========================================================
# BASE_PATH = "G:/Analisi e Performance Prodotti/Fact Sheet New 2020"
BASE_PATH = "C:/Users/Lenovo/Desktop/Mediolanum"

tracciato_rating = pd.read_excel(
    f"{BASE_PATH}/Tracciati/Tracciato MIF {anno} {mese}.xlsx",
    sheet_name="RATING"
)

tracciato_asset = pd.read_excel(
    f"{BASE_PATH}/Tracciati/Tracciato MIF {anno} {mese}.xlsx",
    sheet_name="ASSET_5D",
    header=1
)

rating = tracciato_rating.copy()
rating["FAMIGLIA"] = "MIF"

# ==========================================================
# 🔗 MERGE OBB
# ==========================================================
rating = rating.merge(
    tracciato_asset[["COLLECTION_DB_CODE", "OBB"]],
    on="COLLECTION_DB_CODE",
    how="left"
)

# ==========================================================
# 🔗 PREPARO NAV DA PTF_MIF
# ==========================================================
nav_region = (
    PTF_MIF[[
        "COLLECTION_DB_CODE",
        "TOTALE_NAV_ITALIA",
        "TOTALE_NAV_SPAGNA",
        "TOTALE_NAV_GERMANIA",
        "TOTALE_NAV_ESTERO"
    ]]
    .drop_duplicates()
)

# ==========================================================
# 🔗 MERGE NAV REGIONI
# ==========================================================
rating = rating.merge(
    nav_region,
    on="COLLECTION_DB_CODE",
    how="left"
)

# %%
# ==========================================================
# 💰 NAV OBB PER REGIONE
# ==========================================================
rating["NAV_OBB_ITA"] = rating["TOTALE_NAV_ITALIA"] * rating["OBB"]
rating["NAV_OBB_SPA"] = rating["TOTALE_NAV_SPAGNA"] * rating["OBB"]
rating["NAV_OBB_GER"] = rating["TOTALE_NAV_GERMANIA"] * rating["OBB"]
rating["NAV_OBB_ESTERO"] = rating["TOTALE_NAV_ESTERO"] * rating["OBB"]

# ==========================================================
# 💰 MV_OBB (PESO * NAV_OBB)
# ==========================================================
rating["MV_OBB_ITA"] = rating["PESO"] * rating["NAV_OBB_ITA"]
rating["MV_OBB_SPA"] = rating["PESO"] * rating["NAV_OBB_SPA"]
rating["MV_OBB_GER"] = rating["PESO"] * rating["NAV_OBB_GER"]
rating["MV_OBB_ESTERO"] = rating["PESO"] * rating["NAV_OBB_ESTERO"]

# ==========================================================
# 🔄 RESHAPE LONG (CON TUTTI I PASSAGGI)
# ==========================================================
df_long = pd.concat([

    rating.assign(
        REGIONE="MIF ITALIA",
        TOTALE_NAV_REGION=rating["TOTALE_NAV_ITALIA"],
        TOTALE_NAV_OBB_REGION=rating["NAV_OBB_ITA"],
        MV_OBB=rating["MV_OBB_ITA"]
    ),

    rating.assign(
        REGIONE="MIF SPAGNA",
        TOTALE_NAV_REGION=rating["TOTALE_NAV_SPAGNA"],
        TOTALE_NAV_OBB_REGION=rating["NAV_OBB_SPA"],
        MV_OBB=rating["MV_OBB_SPA"]
    ),

    rating.assign(
        REGIONE="MIF GERMANIA",
        TOTALE_NAV_REGION=rating["TOTALE_NAV_GERMANIA"],
        TOTALE_NAV_OBB_REGION=rating["NAV_OBB_GER"],
        MV_OBB=rating["MV_OBB_GER"]
    ),

    rating.assign(
        REGIONE="MIF TERZI ESTERO",
        TOTALE_NAV_REGION=rating["TOTALE_NAV_ESTERO"],
        TOTALE_NAV_OBB_REGION=rating["NAV_OBB_ESTERO"],
        MV_OBB=rating["MV_OBB_ESTERO"]
    )

], ignore_index=True)

# ==========================================================
# ⚖️ TOTALE_RIBASATO
# ==========================================================
df_long["TOTALE_RIBASATO"] = (
    df_long["MV_OBB"] /
    df_long.groupby("REGIONE")["MV_OBB"].transform("sum")
)

# ==========================================================
# 📦 OUTPUT FINALE COMPLETO
# ==========================================================
df_finale = df_long[[
    "TIM_REPORTING_DATE",
    "COLLECTION_DB_CODE",
    "COD_RATING",
    "TIPO_RATING",
    "PESO",
    "FAMIGLIA",
    "REGIONE",
    "OBB",
    "TOTALE_NAV_REGION",
    "TOTALE_NAV_OBB_REGION",
    "MV_OBB",
    "TOTALE_RIBASATO"
]]

# %%
# ==========================================================
# ✅ CHECK
# ==========================================================
print("\n✅ Check per regione (devono fare 1):")
print(
    df_finale.groupby("REGIONE")["TOTALE_RIBASATO"]
    .sum()
    .reset_index()
)


# ---- Info ----
info = {
    "Nome codice": ["STORIA_DEI_FONDI"],
    "Nome codice di origine": ["Rating_MIF_x_Region_ex_FoHF"],
    # "Percorso": [r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Codici - Storia dei fondi"],
    "Percorso": [r"C:\Users\Lenovo\Desktop\Mediolanum"],
    "Data stampa": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
}
df_info = pd.DataFrame(info)

# ---- Percorso file dinamico ----
# output_path = f"G:/Analisi e Performance Prodotti/PowerBi MAA/Storia dei fondi/Rating_MIF_x_Region_ex_FoHF.xlsx"
output_path = f"C:/Users/Lenovo/Desktop/Mediolanum/Rating_MIF_x_Region_ex_FoHF.xlsx"

# ---- Salvataggio ----
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df_finale.to_excel(writer, sheet_name="Dati", index=False)
    df_info.to_excel(writer, sheet_name="Info", index=False)

print(f"File salvato in: {output_path}")



# TODO

# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==============================================================
# ==============================================================

# VALUTA Confronto MIF, MGF ed MGE + MIF diviso per Region/paese

# ==============================================================
# ==============================================================


# ⚠️⚠️⚠️ IMPORTANTE ⚠️⚠️⚠️
# ABBIAMO ESCLUSO IL FOHF
# SE VOGLIAMO CONSIDERARLO BASTA COMMENTARE LE RIGHE DI CODICE: 
    # PTF_MIF = PTF_MIF[PTF_MIF['COLLECTION_DB_CODE'] != 'DBFoHF']


#------------------------  MIF_ITA -----------------------------------------------

from ptf_mif_diviso_per_region_ISIN import ptf_mif_region_fn
TRACCIATO_MIF_ASSET, PTF_MIF, totale = ptf_mif_region_fn (anno = anno, mese = mese, giorno = giorno, giorno_lavorativo = giorno_lavorativo)


# Escludo FOHF
PTF_MIF = PTF_MIF[PTF_MIF['COLLECTION_DB_CODE'] != 'DBFoHF']

# carichiamo tracciato MIF valuta

# tracciato_mif_valuta = pd.read_excel(f'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Tracciati/Tracciato MIF {anno} {mese}.xlsx', sheet_name = "VALUTA")
tracciato_mif_valuta = pd.read_excel(f'C:/Users/Lenovo/Desktop/Mediolanum/Tracciato MIF {anno} {mese}.xlsx', sheet_name = "VALUTA")

# Estraiamo una mappa con CODICE_FONDO_DETT e NAV_ITALIA_TOTALE


# dobbiamo utilizzare il codice fondo dett perché in realtà nello sheet valuta nonostante l'intestazione sia collection db code
# viene utilizzo il codocie fondo dett
# nel ptf_mif_diviso_per_region_ISIN abbiamo il nav italia per collection db code, per codice isin e per codice fondo dett

print(PTF_MIF['TOTALE_NAV_ITALIA'].sum())

mappa_nav_italia = PTF_MIF[['CODICE_ISIN', 'CODICE_FONDO_DETT', 'TOTALE_NAV_ITALIA', 'COLLECTION_DB_CODE']]

print(mappa_nav_italia['TOTALE_NAV_ITALIA'].sum())

# sommiamo TOTALE_NAV_ITALIA per CODICE_FONDO_DETT
somma_nav_italia = mappa_nav_italia.groupby('CODICE_FONDO_DETT', as_index=False).agg({
    'TOTALE_NAV_ITALIA': 'sum',
    'COLLECTION_DB_CODE': 'first'  # o 'max', se preferisci
})

somma_nav_italia = somma_nav_italia.groupby('CODICE_FONDO_DETT', as_index=False).agg({
    'TOTALE_NAV_ITALIA': 'sum',
    'COLLECTION_DB_CODE': 'first'  # o 'max', se preferisci
})


# facciamo un check: fammi la somma della colonna TOTALE_NAV_ITALIA
print(mappa_nav_italia['TOTALE_NAV_ITALIA'].sum())


# aggiungiamo nel tracciato mif della valuta la riga del fof

#valore_nav_dbfof = mappa_nav_italia.loc[mappa_nav_italia["CODICE_FONDO_DETT"] == "DBFOF", "TOTALE_NAV_ITALIA"].values[0]
"""
# Creiamo la nuova riga come un dizionario
new_row = pd.DataFrame({
'COLLECTION_DB_CODE':['DBFOF'],
#'CODICE_FONDO_DETT': ['DBFOF'],
'CURRENCY': ['EUR'],                
'DES_CURRENCY': ['Euro'],     
'VALUTA': ['Euro'],           
'PESO': [1],                
#'TOTALE_NAV_ITALIA': [valore_nav_dbfof]                   
})


# Concatenare la nuova riga con il dataframe PTF_MIF
tracciato_mif_valuta = pd.concat([tracciato_mif_valuta, new_row], ignore_index=True)
"""
# nel tracciato della valuta dei mif c'è un problema/errore per i Gamax
# per tutti i fondi viene utilizzato il CODICE_FONDO_DETT, quindi distinzione se H o no. 
# per il gamax invece viene usato il COLLECTION_DB_CODE, che per i GAMAX è diverso da CODICE_FONDO_DETT. 
# quindi: creiamo una nuova colonna "db_tracciato_valuta" nel df somma_nav_italia  che è un mix tra COLLECTION_DB_CODE e CODICE_FONDO_DETT
    # -> se COLLECTION_DB_CODE = CODICE_FONDO_DETT => db_tracciato_valuta = CODICE_FONDO_DETT
    # -> se COLLECTION_DB_CODE <> CODICE_FONDO_DETT => db_tracciato_valuta = COLLECTION_DB_CODE

# Mappa delle eccezioni 
eccezioni = {
    'DB41330': 'LU4133',
    'DB41340': 'LU4134',
    'DB41360': 'LU4136',
    'DB5606': 'LU5606',
    'DB7554': 'LU7554'
}


# %%
def get_db_tracciato_valuta(row):
    return eccezioni.get(row['CODICE_FONDO_DETT'], row['CODICE_FONDO_DETT'])

somma_nav_italia['db_tracciato_valuta'] = somma_nav_italia.apply(get_db_tracciato_valuta, axis=1)

# rinominiamo COLLECTION_DB_CODE in db_tracciato_valuta

tracciato_mif_valuta = tracciato_mif_valuta.rename(columns={'COLLECTION_DB_CODE': 'db_tracciato_valuta'})

# ora aggiungiamo il TOTALE_NAV_ITALIA al tracciato mif valuta perché è da tracciato che prendiamo le diversificaizoni

tracciato_mif_valuta = tracciato_mif_valuta.merge(
    somma_nav_italia[['db_tracciato_valuta', 'TOTALE_NAV_ITALIA']],
    on='db_tracciato_valuta',
    how='left'
)

# se TOTALE_NAV_ITALIA = nan => 0
# questo perché sono presenti in valuta due codici DB6214H e DB6575H che non esistono nel principale
# probabilmente sono le classi hedged dei due fondi che sono state chiuse ma che sono rimaste in valuta

tracciato_mif_valuta['TOTALE_NAV_ITALIA'] = tracciato_mif_valuta['TOTALE_NAV_ITALIA'].fillna(0)

# raggruppiamo le valute

tracciato_mif_valuta['CURRENCY_aggregato'] = tracciato_mif_valuta['CURRENCY'].apply(
    lambda x: x if x in ['EUR', 'USD', 'JPY'] else 'Altre valute'
)

# calcoliamo la somma dei nav italia
totale_nav_italia = somma_nav_italia['TOTALE_NAV_ITALIA'].sum()
print("Totale NAV Italia (somma dei valori univoci):", totale_nav_italia)

# ora possiamo eseguire il ribilanciamento sul NAV ITALIA
tracciato_mif_valuta['TOTALE_Ribasato_Italia'] = tracciato_mif_valuta['PESO'] * tracciato_mif_valuta['TOTALE_NAV_ITALIA'] / totale_nav_italia

#Verifichiamo che la somma di Totale Ribasato sia 1
print("Somma della colonna TOTALE_Ribasato_Italia:", tracciato_mif_valuta['TOTALE_Ribasato_Italia'].sum())

#Se la somma di Totale Ribasato non torna ad 1, normalizza i valori per far tornare il risultato ad 1
tracciato_mif_valuta['TOTALE_Ribasato_Italia'] /= tracciato_mif_valuta['TOTALE_Ribasato_Italia'].sum() if tracciato_mif_valuta['TOTALE_Ribasato_Italia'].sum() != 1 else 1
print("Somma della colonna TOTALE_Ribasato_Italia:", tracciato_mif_valuta['TOTALE_Ribasato_Italia'].sum())


riga_valuta_mif_ita = tracciato_mif_valuta.pivot_table(
    index=["CURRENCY", "CURRENCY_aggregato"],  # le dimensioni su cui raggruppi
    values="TOTALE_Ribasato_Italia",           # la colonna da aggregare
    aggfunc="sum",                             # funzione di aggregazione
    fill_value=0                               # opzionale: sostituisce i NaN con 0
).reset_index()

riga_valuta_mif_ita['data'] = data

riga_valuta_mif_ita['famiglia'] = 'MIF ITALIA'

print('riga_valuta_mif_ita: ', riga_valuta_mif_ita['TOTALE_Ribasato_Italia'].sum())

riga_valuta_mif_ita['NAV_FAMIGLIA'] = totale_nav_italia

riga_valuta_mif_ita = riga_valuta_mif_ita.rename(columns={'TOTALE_Ribasato_Italia': 'TOTALE_Ribasato',
                                                          })


#------------------------  MIF_SPA -----------------------------------------------

#from ptf_mif_diviso_per_region_ISIN import ptf_mif_region_fn
#TRACCIATO_MIF_ASSET, PTF_MIF, totale = ptf_mif_region_fn (anno = anno, mese = mese, giorno = giorno, giorno_lavorativo = giorno_lavorativo)

print(PTF_MIF['TOTALE_NAV_SPAGNA'].sum())

mappa_nav_spa = PTF_MIF[['CODICE_ISIN', 'CODICE_FONDO_DETT', 'TOTALE_NAV_SPAGNA', 'COLLECTION_DB_CODE']]

print(mappa_nav_spa['TOTALE_NAV_SPAGNA'].sum())

# sommiamo TOTALE_NAV_ITALIA per CODICE_FONDO_DETT
somma_nav_spa = mappa_nav_spa.groupby('CODICE_FONDO_DETT', as_index=False).agg({
    'TOTALE_NAV_SPAGNA': 'sum',
    'COLLECTION_DB_CODE': 'first'  # o 'max', se preferisci
})

somma_nav_spa = somma_nav_spa.groupby('CODICE_FONDO_DETT', as_index=False).agg({
    'TOTALE_NAV_SPAGNA': 'sum',
    'COLLECTION_DB_CODE': 'first'  # o 'max', se preferisci
})


# facciamo un check: fammi la somma della colonna TOTALE_NAV_ITALIA
print(mappa_nav_spa['TOTALE_NAV_SPAGNA'].sum())


# aggiungiamo nel tracciato mif della valuta la riga del fof

#valore_nav_dbfof = mappa_nav_italia.loc[mappa_nav_italia["CODICE_FONDO_DETT"] == "DBFOF", "TOTALE_NAV_ITALIA"].values[0]
"""
# Creiamo la nuova riga come un dizionario
new_row = pd.DataFrame({
'COLLECTION_DB_CODE':['DBFOF'],
#'CODICE_FONDO_DETT': ['DBFOF'],
'CURRENCY': ['EUR'],                
'DES_CURRENCY': ['Euro'],     
'VALUTA': ['Euro'],           
'PESO': [1],                
#'TOTALE_NAV_ITALIA': [valore_nav_dbfof]                   
})


# Concatenare la nuova riga con il dataframe PTF_MIF
tracciato_mif_valuta = pd.concat([tracciato_mif_valuta, new_row], ignore_index=True)
"""
# nel tracciato della valuta dei mif c'è un problema/errore per i Gamax
# per tutti i fondi viene utilizzato il CODICE_FONDO_DETT, quindi distinzione se H o no. 
# per il gamax invece viene usato il COLLECTION_DB_CODE, che per i GAMAX è diverso da CODICE_FONDO_DETT. 
# quindi: creiamo una nuova colonna "db_tracciato_valuta" nel df somma_nav_italia  che è un mix tra COLLECTION_DB_CODE e CODICE_FONDO_DETT
    # -> se COLLECTION_DB_CODE = CODICE_FONDO_DETT => db_tracciato_valuta = CODICE_FONDO_DETT
    # -> se COLLECTION_DB_CODE <> CODICE_FONDO_DETT => db_tracciato_valuta = COLLECTION_DB_CODE

# Mappa delle eccezioni 
eccezioni = {
    'DB41330': 'LU4133',
    'DB41340': 'LU4134',
    'DB41360': 'LU4136',
    'DB5606': 'LU5606',
    'DB7554': 'LU7554'
}


# %%
def get_db_tracciato_valuta(row):
    return eccezioni.get(row['CODICE_FONDO_DETT'], row['CODICE_FONDO_DETT'])

somma_nav_spa['db_tracciato_valuta'] = somma_nav_spa.apply(get_db_tracciato_valuta, axis=1)

# rinominiamo COLLECTION_DB_CODE in db_tracciato_valuta

tracciato_mif_valuta = tracciato_mif_valuta.rename(columns={'COLLECTION_DB_CODE': 'db_tracciato_valuta'})

# ora aggiungiamo il TOTALE_NAV_ITALIA al tracciato mif valuta perché è da tracciato che prendiamo le diversificaizoni

tracciato_mif_valuta = tracciato_mif_valuta.merge(
    somma_nav_spa[['db_tracciato_valuta', 'TOTALE_NAV_SPAGNA']],
    on='db_tracciato_valuta',
    how='left'
)

# se TOTALE_NAV_ITALIA = nan => 0
# questo perché sono presenti in valuta due codici DB6214H e DB6575H che non esistono nel principale
# probabilmente sono le classi hedged dei due fondi che sono state chiuse ma che sono rimaste in valuta

tracciato_mif_valuta['TOTALE_NAV_SPAGNA'] = tracciato_mif_valuta['TOTALE_NAV_SPAGNA'].fillna(0)

# raggruppiamo le valute

tracciato_mif_valuta['CURRENCY_aggregato'] = tracciato_mif_valuta['CURRENCY'].apply(
    lambda x: x if x in ['EUR', 'USD', 'JPY'] else 'Altre valute'
)

# calcoliamo la somma dei nav italia
totale_nav_spa = somma_nav_spa['TOTALE_NAV_SPAGNA'].sum()
print("TOTALE_NAV_SPAGNA (somma dei valori univoci):", totale_nav_spa)

# ora possiamo eseguire il ribilanciamento sul NAV ITALIA
tracciato_mif_valuta['TOTALE_Ribasato_spa'] = tracciato_mif_valuta['PESO'] * tracciato_mif_valuta['TOTALE_NAV_SPAGNA'] / totale_nav_spa

#Verifichiamo che la somma di Totale Ribasato sia 1
print("Somma della colonna TOTALE_Ribasato_spa:", tracciato_mif_valuta['TOTALE_Ribasato_spa'].sum())

#Se la somma di Totale Ribasato non torna ad 1, normalizza i valori per far tornare il risultato ad 1
tracciato_mif_valuta['TOTALE_Ribasato_spa'] /= tracciato_mif_valuta['TOTALE_Ribasato_spa'].sum() if tracciato_mif_valuta['TOTALE_Ribasato_spa'].sum() != 1 else 1
print("Somma della colonna TOTALE_Ribasato_spa:", tracciato_mif_valuta['TOTALE_Ribasato_spa'].sum())


riga_valuta_mif_spa = tracciato_mif_valuta.pivot_table(
    index=["CURRENCY", "CURRENCY_aggregato"],  # le dimensioni su cui raggruppi
    values="TOTALE_Ribasato_spa",           # la colonna da aggregare
    aggfunc="sum",                             # funzione di aggregazione
    fill_value=0                               # opzionale: sostituisce i NaN con 0
).reset_index()

riga_valuta_mif_spa['data'] = data

riga_valuta_mif_spa['famiglia'] = 'MIF SPAGNA'

print('riga_valuta_mif_ita: ', riga_valuta_mif_spa['TOTALE_Ribasato_spa'].sum())

riga_valuta_mif_spa['NAV_FAMIGLIA'] = totale_nav_spa

riga_valuta_mif_spa = riga_valuta_mif_spa.rename(columns={'TOTALE_Ribasato_spa': 'TOTALE_Ribasato',
                                                          })



#------------------------  MIF_GER -----------------------------------------------


#from ptf_mif_diviso_per_region_ISIN import ptf_mif_region_fn
#TRACCIATO_MIF_ASSET, PTF_MIF, totale = ptf_mif_region_fn (anno = anno, mese = mese, giorno = giorno, giorno_lavorativo = giorno_lavorativo)

print(PTF_MIF['TOTALE_NAV_GERMANIA'].sum())

mappa_nav_ger = PTF_MIF[['CODICE_ISIN', 'CODICE_FONDO_DETT', 'TOTALE_NAV_GERMANIA', 'COLLECTION_DB_CODE']]

print(mappa_nav_ger['TOTALE_NAV_GERMANIA'].sum())

# sommiamo TOTALE_NAV_ITALIA per CODICE_FONDO_DETT
somma_nav_ger = mappa_nav_ger.groupby('CODICE_FONDO_DETT', as_index=False).agg({
    'TOTALE_NAV_GERMANIA': 'sum',
    'COLLECTION_DB_CODE': 'first'  # o 'max', se preferisci
})

somma_nav_ger = somma_nav_ger.groupby('CODICE_FONDO_DETT', as_index=False).agg({
    'TOTALE_NAV_GERMANIA': 'sum',
    'COLLECTION_DB_CODE': 'first'  # o 'max', se preferisci
})


# facciamo un check: fammi la somma della colonna TOTALE_NAV_ITALIA
print(mappa_nav_ger['TOTALE_NAV_GERMANIA'].sum())


# aggiungiamo nel tracciato mif della valuta la riga del fof

#valore_nav_dbfof = mappa_nav_italia.loc[mappa_nav_italia["CODICE_FONDO_DETT"] == "DBFOF", "TOTALE_NAV_ITALIA"].values[0]
"""
# Creiamo la nuova riga come un dizionario
new_row = pd.DataFrame({
'COLLECTION_DB_CODE':['DBFOF'],
#'CODICE_FONDO_DETT': ['DBFOF'],
'CURRENCY': ['EUR'],                
'DES_CURRENCY': ['Euro'],     
'VALUTA': ['Euro'],           
'PESO': [1],                
#'TOTALE_NAV_ITALIA': [valore_nav_dbfof]                   
})


# Concatenare la nuova riga con il dataframe PTF_MIF
tracciato_mif_valuta = pd.concat([tracciato_mif_valuta, new_row], ignore_index=True)
"""
# nel tracciato della valuta dei mif c'è un problema/errore per i Gamax
# per tutti i fondi viene utilizzato il CODICE_FONDO_DETT, quindi distinzione se H o no. 
# per il gamax invece viene usato il COLLECTION_DB_CODE, che per i GAMAX è diverso da CODICE_FONDO_DETT. 
# quindi: creiamo una nuova colonna "db_tracciato_valuta" nel df somma_nav_italia  che è un mix tra COLLECTION_DB_CODE e CODICE_FONDO_DETT
    # -> se COLLECTION_DB_CODE = CODICE_FONDO_DETT => db_tracciato_valuta = CODICE_FONDO_DETT
    # -> se COLLECTION_DB_CODE <> CODICE_FONDO_DETT => db_tracciato_valuta = COLLECTION_DB_CODE

# Mappa delle eccezioni 
eccezioni = {
    'DB41330': 'LU4133',
    'DB41340': 'LU4134',
    'DB41360': 'LU4136',
    'DB5606': 'LU5606',
    'DB7554': 'LU7554'
}


# %%
def get_db_tracciato_valuta(row):
    return eccezioni.get(row['CODICE_FONDO_DETT'], row['CODICE_FONDO_DETT'])

somma_nav_ger['db_tracciato_valuta'] = somma_nav_ger.apply(get_db_tracciato_valuta, axis=1)

# rinominiamo COLLECTION_DB_CODE in db_tracciato_valuta

tracciato_mif_valuta = tracciato_mif_valuta.rename(columns={'COLLECTION_DB_CODE': 'db_tracciato_valuta'})

# ora aggiungiamo il TOTALE_NAV_ITALIA al tracciato mif valuta perché è da tracciato che prendiamo le diversificaizoni

tracciato_mif_valuta = tracciato_mif_valuta.merge(
    somma_nav_ger[['db_tracciato_valuta', 'TOTALE_NAV_GERMANIA']],
    on='db_tracciato_valuta',
    how='left'
)

# se TOTALE_NAV_ITALIA = nan => 0
# questo perché sono presenti in valuta due codici DB6214H e DB6575H che non esistono nel principale
# probabilmente sono le classi hedged dei due fondi che sono state chiuse ma che sono rimaste in valuta

tracciato_mif_valuta['TOTALE_NAV_GERMANIA'] = tracciato_mif_valuta['TOTALE_NAV_GERMANIA'].fillna(0)

# raggruppiamo le valute

tracciato_mif_valuta['CURRENCY_aggregato'] = tracciato_mif_valuta['CURRENCY'].apply(
    lambda x: x if x in ['EUR', 'USD', 'JPY'] else 'Altre valute'
)

# calcoliamo la somma dei nav italia
totale_nav_ger = somma_nav_ger['TOTALE_NAV_GERMANIA'].sum()
print("totale_nav_ger (somma dei valori univoci):", totale_nav_ger)

# ora possiamo eseguire il ribilanciamento sul NAV ITALIA
tracciato_mif_valuta['TOTALE_Ribasato_ger'] = tracciato_mif_valuta['PESO'] * tracciato_mif_valuta['TOTALE_NAV_GERMANIA'] / totale_nav_ger

#Verifichiamo che la somma di Totale Ribasato sia 1
print("Somma della colonna TOTALE_Ribasato_ger:", tracciato_mif_valuta['TOTALE_Ribasato_ger'].sum())

#Se la somma di Totale Ribasato non torna ad 1, normalizza i valori per far tornare il risultato ad 1
tracciato_mif_valuta['TOTALE_Ribasato_ger'] /= tracciato_mif_valuta['TOTALE_Ribasato_ger'].sum() if tracciato_mif_valuta['TOTALE_Ribasato_ger'].sum() != 1 else 1
print("Somma della colonna TOTALE_Ribasato_ger:", tracciato_mif_valuta['TOTALE_Ribasato_ger'].sum())


riga_valuta_mif_ger = tracciato_mif_valuta.pivot_table(
    index=["CURRENCY", "CURRENCY_aggregato"],  # le dimensioni su cui raggruppi
    values="TOTALE_Ribasato_ger",           # la colonna da aggregare
    aggfunc="sum",                             # funzione di aggregazione
    fill_value=0                               # opzionale: sostituisce i NaN con 0
).reset_index()

riga_valuta_mif_ger['data'] = data

riga_valuta_mif_ger['famiglia'] = 'MIF GERMANIA'

print('riga_valuta_mif_ger: ', riga_valuta_mif_ger['TOTALE_Ribasato_ger'].sum())

riga_valuta_mif_ger['NAV_FAMIGLIA'] = totale_nav_ger

riga_valuta_mif_ger = riga_valuta_mif_ger.rename(columns={'TOTALE_Ribasato_ger': 'TOTALE_Ribasato',
                                                          })



#------------------------  MIF_ESTERO -----------------------------------------------

#from ptf_mif_diviso_per_region_ISIN import ptf_mif_region_fn
#TRACCIATO_MIF_ASSET, PTF_MIF, totale = ptf_mif_region_fn (anno = anno, mese = mese, giorno = giorno, giorno_lavorativo = giorno_lavorativo)

print(PTF_MIF['TOTALE_NAV_ESTERO'].sum())

mappa_nav_estero = PTF_MIF[['CODICE_ISIN', 'CODICE_FONDO_DETT', 'TOTALE_NAV_ESTERO', 'COLLECTION_DB_CODE']]


# sommiamo TOTALE_NAV_ITALIA per CODICE_FONDO_DETT
somma_nav_estero = mappa_nav_estero.groupby('CODICE_FONDO_DETT', as_index=False).agg({
    'TOTALE_NAV_ESTERO': 'sum',
    'COLLECTION_DB_CODE': 'first'  # o 'max', se preferisci
})

# facciamo un check: fammi la somma della colonna TOTALE_NAV_ITALIA
print(mappa_nav_estero['TOTALE_NAV_ESTERO'].sum())

# Mappa delle eccezioni 
eccezioni = {
    'DB41330': 'LU4133',
    'DB41340': 'LU4134',
    'DB41360': 'LU4136',
    'DB5606': 'LU5606',
    'DB7554': 'LU7554'
}


# %%
def get_db_tracciato_valuta(row):
    return eccezioni.get(row['CODICE_FONDO_DETT'], row['CODICE_FONDO_DETT'])

somma_nav_estero['db_tracciato_valuta'] = somma_nav_estero.apply(get_db_tracciato_valuta, axis=1)

# rinominiamo COLLECTION_DB_CODE in db_tracciato_valuta

tracciato_mif_valuta = tracciato_mif_valuta.rename(columns={'COLLECTION_DB_CODE': 'db_tracciato_valuta'})

# ora aggiungiamo il TOTALE_NAV_ITALIA al tracciato mif valuta perché è da tracciato che prendiamo le diversificaizoni

tracciato_mif_valuta = tracciato_mif_valuta.merge(
    somma_nav_estero[['db_tracciato_valuta', 'TOTALE_NAV_ESTERO']],
    on='db_tracciato_valuta',
    how='left'
)

# se TOTALE_NAV_ITALIA = nan => 0
# questo perché sono presenti in valuta due codici DB6214H e DB6575H che non esistono nel principale
# probabilmente sono le classi hedged dei due fondi che sono state chiuse ma che sono rimaste in valuta

tracciato_mif_valuta['TOTALE_NAV_ESTERO'] = tracciato_mif_valuta['TOTALE_NAV_ESTERO'].fillna(0)

# raggruppiamo le valute

tracciato_mif_valuta['CURRENCY_aggregato'] = tracciato_mif_valuta['CURRENCY'].apply(
    lambda x: x if x in ['EUR', 'USD', 'JPY'] else 'Altre valute'
)

# calcoliamo la somma dei nav italia
totale_nav_estero = somma_nav_estero['TOTALE_NAV_ESTERO'].sum()
print("TOTALE_NAV_ESTERO (somma dei valori univoci):", totale_nav_estero)

# ora possiamo eseguire il ribilanciamento sul NAV ITALIA
tracciato_mif_valuta['TOTALE_Ribasato_estero'] = tracciato_mif_valuta['PESO'] * tracciato_mif_valuta['TOTALE_NAV_ESTERO'] / totale_nav_estero

#Verifichiamo che la somma di Totale Ribasato sia 1
print("Somma della colonna TOTALE_Ribasato_estero:", tracciato_mif_valuta['TOTALE_Ribasato_estero'].sum())

#Se la somma di Totale Ribasato non torna ad 1, normalizza i valori per far tornare il risultato ad 1
tracciato_mif_valuta['TOTALE_Ribasato_estero'] /= tracciato_mif_valuta['TOTALE_Ribasato_estero'].sum() if tracciato_mif_valuta['TOTALE_Ribasato_estero'].sum() != 1 else 1
print("Somma della colonna TOTALE_Ribasato_estero:", tracciato_mif_valuta['TOTALE_Ribasato_estero'].sum())


riga_valuta_mif_estero = tracciato_mif_valuta.pivot_table(
    index=["CURRENCY", "CURRENCY_aggregato"],  # le dimensioni su cui raggruppi
    values="TOTALE_Ribasato_estero",           # la colonna da aggregare
    aggfunc="sum",                             # funzione di aggregazione
    fill_value=0                               # opzionale: sostituisce i NaN con 0
).reset_index()

riga_valuta_mif_estero['data'] = data

riga_valuta_mif_estero['famiglia'] = 'MIF TERZI ESTERO'

print('riga_valuta_mif_estero: ', riga_valuta_mif_estero['TOTALE_Ribasato_estero'].sum())

riga_valuta_mif_estero['NAV_FAMIGLIA'] = totale_nav_estero

riga_valuta_mif_estero = riga_valuta_mif_estero.rename(columns={'TOTALE_Ribasato_estero': 'TOTALE_Ribasato',
                                                          })




#--------------------------------------------------------

# Concatenazione dei vari DataFrame
valuta_captive_mif = pd.concat([riga_valuta_mif_spa, 
                                  riga_valuta_mif_ita, 
                                  riga_valuta_mif_ger,
                                  riga_valuta_mif_estero
                                 ], ignore_index=True)



sum_unique_nav_famiglia = valuta_captive_mif['NAV_FAMIGLIA'].drop_duplicates().sum()

# Mostra il risultato
print('NAV valuta_captive_mif:', sum_unique_nav_famiglia)

valuta_captive_mif['nav_captive_mif'] = sum_unique_nav_famiglia

# Rinomina la colonna TOTALE_Ribasato in TOTALE_Ribasato_Old
valuta_captive_mif = valuta_captive_mif.rename(columns={'TOTALE_Ribasato': 'TOTALE_Ribasato_Old'})

# Calcola la nuova colonna TOTALE_Ribasato_New
valuta_captive_mif['TOTALE_Ribasato_New'] = valuta_captive_mif['TOTALE_Ribasato_Old'] * valuta_captive_mif['NAV_FAMIGLIA'] / valuta_captive_mif['nav_captive_mif']

#check
somma_totale_ribasato_new = valuta_captive_mif['TOTALE_Ribasato_New'].sum()
print("Somma di TOTALE_Ribasato_New:", somma_totale_ribasato_new)
#Se la somma di Totale Ribasato non torna ad 1, normalizza i valori per far tornare il risultato ad 1
valuta_captive_mif['TOTALE_Ribasato_New'] /= valuta_captive_mif['TOTALE_Ribasato_New'].sum() if valuta_captive_mif['TOTALE_Ribasato_New'].sum() != 1 else 1
print("Somma normalizzata della colonna TOTALE_Ribasato_New:", valuta_captive_mif['TOTALE_Ribasato_New'].sum())

pivot_valuta_captive_italia = (
    valuta_captive_mif
    .groupby(['data', 'CURRENCY_aggregato'], as_index=False)
    .agg({
        'TOTALE_Ribasato_New': 'sum',
        'nav_captive_mif': 'first'  # oppure 'mean' se cambia
    })
)


# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣


# facciamo anche il confronto tra MIF totale, MGF ed MGE

# quindi MIF siamo a posto 
# carichiamo valuta MGF

# MGF -> non possiamo usare quella del comitato perché ricomprende anche i tre comparti del previgest
    # ed include anche F. chiuso Private Market e F. Real Estate
    # utilizziamo la valuta MGF da tracciato 

# lo stesso per MGE 

import os
import glob

# base_tracciati = 'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Tracciati'
base_tracciati = 'C:/Users/Lenovo/Desktop/Mediolanum'

dati_valuta = []

files = glob.glob(os.path.join(base_tracciati, 'Tracciato *.xlsx'))

for file in files:
    nome_file = os.path.basename(file)
    
    # filtro solo MGF e MGE per anno/mese
    if (
        nome_file.startswith(f'Tracciato MGF {anno} {mese}') or
        nome_file.startswith(f'Tracciato MGE {anno} {mese}')
    ):
        try:
            df = pd.read_excel(file, sheet_name='VALUTA')
            
            # aggiungo info famiglia
            if 'MGF' in nome_file:
                df['FAMIGLIA'] = 'MGF'
            elif 'MGE' in nome_file:
                df['FAMIGLIA'] = 'MGE'
            
            dati_valuta.append(df)
            print(f"✅ Caricato: {nome_file}")
        
        except Exception as e:
            print(f"❌ Errore su {nome_file}: {e}")

# concat finale
valuta_totale = pd.concat(dati_valuta, ignore_index=True)

print("📊 Righe totali:", len(valuta_totale))

# valute di riferimento

valuta_totale['CURRENCY_aggregato'] = valuta_totale['CURRENCY'].apply(
    lambda x: x if x in ['EUR', 'USD', 'JPY'] else 'Altre valute'
)

# prendiamo i ptf per recuperare i NAV dei COLLECTION_DB_CODE

# base_portafogli = 'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Portafoglio/'
base_portafogli = 'C:/Users/Lenovo/Desktop/Mediolanum'

dfs = []

for fam in ['MGF', 'MGE']:
    
    file = os.path.join(
        base_portafogli,
        f'PTF_FUNDLOOKTHROUGH {fam} {anno} {mese}.xlsx'
    )
    
    df = pd.read_excel(file, usecols=['COLLECTION_DB_CODE', 'NAV'])
    
    dfs.append(df)

# unisco i due file
ptf_all = pd.concat(dfs, ignore_index=True)

# prendo NAV univoco per fondo
ptf_nav = (
    ptf_all
    .groupby('COLLECTION_DB_CODE', as_index=False)
    .agg({'NAV': 'first'})   
)

valuta_totale = valuta_totale.merge(
    ptf_nav,
    on='COLLECTION_DB_CODE',
    how='left'
)


# ribasiamo il peso sulla famiglia 

nav_fondo_famiglia = ptf_nav.merge(
    valuta_totale[['COLLECTION_DB_CODE', 'FAMIGLIA']].drop_duplicates(),
    on='COLLECTION_DB_CODE',
    how='left'
)

nav_famiglia = (
    nav_fondo_famiglia
    .groupby('FAMIGLIA', as_index=False)
    .agg({'NAV': 'sum'})
    .rename(columns={'NAV': 'NAV_FAMIGLIA'})
)

valuta_totale = valuta_totale.merge(
    nav_famiglia,
    on='FAMIGLIA',
    how='left'
)


valuta_totale['PESO_VALUTA_FAMIGLIA'] = (
    valuta_totale['PESO'] * valuta_totale['NAV'] / valuta_totale['NAV_FAMIGLIA']
)

valuta_famiglia = (
    valuta_totale
    .groupby(['FAMIGLIA', 'CURRENCY_aggregato'], as_index=False)
    .agg({'PESO_VALUTA_FAMIGLIA': 'sum'})
)

valuta_famiglia.groupby('FAMIGLIA')['PESO_VALUTA_FAMIGLIA'].sum()

valuta_famiglia['PESO_VALUTA_FAMIGLIA_NORM'] = (
    valuta_famiglia['PESO_VALUTA_FAMIGLIA'] /
    valuta_famiglia.groupby('FAMIGLIA')['PESO_VALUTA_FAMIGLIA'].transform('sum')
)

valuta_famiglia.groupby('FAMIGLIA')['PESO_VALUTA_FAMIGLIA_NORM'].sum()

# aggiungiamo la valutadi di MIF a valuta_totale con FAMIGLIA = MIF

# %%
# ============================================================
# 🟢 PREPARAZIONE MIF
# ============================================================
df_mif = pivot_valuta_captive_italia.copy()

df_mif = df_mif.rename(columns={
    'TOTALE_Ribasato_New': 'PESO'
})

df_mif['FAMIGLIA'] = 'MIF'

df_mif = df_mif[['data', 'FAMIGLIA', 'CURRENCY_aggregato', 'PESO']]


# ============================================================
# 🔵 PREPARAZIONE MGF + MGE
# ============================================================
df_fam = valuta_famiglia.copy()

df_fam = df_fam.rename(columns={
    'PESO_VALUTA_FAMIGLIA_NORM': 'PESO'
})

# 👉 aggiungiamo la data (stessa logica di MIF)
df_fam['data'] = data

df_fam = df_fam[['data', 'FAMIGLIA', 'CURRENCY_aggregato', 'PESO']]

# ============================================================
# 🔗 CONCAT FINALE
# ============================================================
valuta_totale = pd.concat([df_mif, df_fam], ignore_index=True)

# ============================================================
# 🧪 CHECK
# ============================================================
check = valuta_totale.groupby(['FAMIGLIA'])['PESO'].sum()
print("Check somma pesi per famiglia:\n", check)

# 👉 sostituzione MGF -> SMFI
valuta_totale['FAMIGLIA'] = valuta_totale['FAMIGLIA'].replace({'MGF': 'SMFI'})


# PERCORSO_FILE_1 = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Region\Nuovi_dati_25_01\Region_Comitato\Valuta\valuta_captive_mif.xlsx"
PERCORSO_FILE_1 = r"C:\Users\Lenovo\Desktop\Mediolanum\valuta_captive_mif.xlsx"

# ============================================================
# 📄 INFO DATI
# ============================================================
info = pd.DataFrame({
    'Nome codice': ['STORIA_DEI_FONDI'],
    'Nome codice di origine': ['valuta_mif_totale_e_per_paese'],
    'Nome_File': ['valuta_captive_mif'],
    'Descrizione': [
        'Dati valuta per MIF diviso per region (escluso FOHF) e per MIF, MGF (non uguale alla valuta comitato perché include previgest, fondi chiusi e private markets) e MGE.'
    ],
    # 'Percorso codice': [r'G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Region\Nuovi_dati_25_01\Region_Comitato\Valuta']
    'Percorso codice': [r'C:\Users\Lenovo\Desktop\Mediolanum']
})

# ============================================================
# 💾 EXPORT EXCEL
# ============================================================
with pd.ExcelWriter(PERCORSO_FILE_1, engine='openpyxl') as writer:
    valuta_captive_mif.to_excel(writer, sheet_name='Valuta_Captive_MIF', index=False)
    pivot_valuta_captive_italia.to_excel(writer, sheet_name='Pivot', index=False)
    valuta_totale.to_excel(writer, sheet_name='valuta_mif_mgf_mge', index=False)
    info.to_excel(writer, sheet_name='info', index=False)

print("✅ File Excel creato correttamente")




# %%
# 🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣🟣

# ==============================================================
# ==============================================================

# GEO MIF Totale diviso OBB e AZ

# ==============================================================
# ==============================================================


# NEI QUARTERLY TRA I COLLECTION_DB_CODE CI SONO ANCHE LE CLASSI H
# quindi per esempio DB2018 e DB2018H -> quindi per i NAV non possiamo usare ptf mif 
# quindi come facciamo per la valuta, recuperiamo i NAV per fondo per classa

# ci serve il tracciato per aggiungere ai file di Pirozzi i collection db code
# file_path_tracciato_mif = f'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Tracciati/Tracciato MIF {anno} {mese}.xlsx'
file_path_tracciato_mif = f'C:/Users/Lenovo/Desktop/Mediolanum/Tracciato MIF {anno} {mese}.xlsx'
sheet_name = 'PRINCIPALE'
tracciato_mif = pd.read_excel(file_path_tracciato_mif, sheet_name = sheet_name, header=0)

# PTF MIF

# file_path_ptf_mif = f'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Portafoglio/PTF_FUNDLOOKTHROUGH MIF {anno} {mese}.xlsx'
file_path_ptf_mif = f'C:/Users/Lenovo/Desktop/Mediolanum/PTF_FUNDLOOKTHROUGH MIF {anno} {mese}.xlsx'
#sheet_name = f'ptf {region}'  # Il nome del foglio diventerà 'ptf spagna'
PTF_MIF = pd.read_excel(file_path_ptf_mif, header=0)

# del ptf mif teniamo solo COLLECTION_DB_CODE e NAV, univoci 
# da qui andiamo ad aggiungere la colonna del CODICE_ISIN da tracciato 

PTF_MIF = PTF_MIF[['COLLECTION_DB_CODE', 'NAV']].drop_duplicates()

tracciato_mif = tracciato_mif.rename(columns={'NAV': 'NAV_x_ISIN'}) # rinominiamo NAV in NAV_x_ISIN

df_merged = PTF_MIF.merge(
    tracciato_mif[['COLLECTION_DB_CODE', 'CODICE_ISIN', 'CODICE_FONDO_DETT', 'NAV_x_ISIN']],
    on='COLLECTION_DB_CODE',
    how='left'
)
   
# aggiungiamo una colonna di check tra tracciato e ptf
# quindi nuova colonna somma_NAV_x_ISIN_x_COLLECTION_DB_CODE somma la colonna NAV_x_ISIN per COLLECTION_DB_CODE
# quindi nuova colonna: se somma_NAV_x_ISIN_x_COLLECTION_DB_CODE = NAV -> True o False

df_merged['somma_NAV_x_ISIN'] = df_merged.groupby('COLLECTION_DB_CODE')['NAV_x_ISIN'].transform('sum')
df_merged['check_NAV'] = df_merged['somma_NAV_x_ISIN'] == df_merged['NAV']

# riportiamo qui la lista di COLLECTION_DB_CODE per cui check_NAV = False

db_nav_diversi_tracciato_ptf = df_merged.loc[~df_merged['check_NAV'], 'COLLECTION_DB_CODE'].unique()

# stampiamo quali sono questi COLLECTION_DB_CODE

print(db_nav_diversi_tracciato_ptf)

# andiamo ad aggiungere una colonna con la differenza tra NAV e somma_NAV_x_ISIN

# ora la differenza la convertiamo in un numero con due cifre decimali

df_merged['differenza'] = (df_merged['NAV'] - df_merged['somma_NAV_x_ISIN']).round(2)

# creiamo una nuova colonna di NAV_corretta, per cui:
# dove differenza = +- 0 => NAV_corretta = NAV_x_ISIN
# per quei db per cui differenza <> +- 0, andiamo ad aggiungeer/togliere la differenza sul NAV_x_ISIN in base a COLLECTION_DB_CODE ma per cui CODICE_ISIN ha NAV_x_ISIN più grande

df_merged['NAV_corretta'] = df_merged['NAV_x_ISIN']

# Calcolo del NAV_x_ISIN massimo per ciascun COLLECTION_DB_CODE

max_nav_per_db = df_merged.groupby('COLLECTION_DB_CODE')['NAV_x_ISIN'].transform('max')

mask_corretti = (df_merged['differenza'].abs() > 0.00) & (df_merged['NAV_x_ISIN'] == max_nav_per_db)

# Funzione per correggere il NAV
def correggi_nav(row):
    return row['NAV_x_ISIN'] + row['differenza']  # somma anche se negativa

# Applica la correzione solo alle righe selezionate
df_merged.loc[mask_corretti, 'NAV_corretta'] = df_merged.loc[mask_corretti].apply(correggi_nav, axis=1)

# CHECK 3 – somma NAV_corretta per COLLECTION_DB_CODE deve essere uguale a NAV
df_merged['somma_NAV_corretta'] = df_merged.groupby('COLLECTION_DB_CODE')['NAV_corretta'].transform('sum')
df_merged['check_NAV_3'] = df_merged['somma_NAV_corretta'] == df_merged['NAV']

# CHECK 4 – differenza tra NAV e somma_NAV_corretta arrotondata
df_merged['differenza_4'] = (df_merged['NAV'] - df_merged['somma_NAV_corretta']).round(2)
df_merged['check_NAV_4'] = df_merged['differenza_4'].apply(lambda x: 'OK' if x == 0 else 'KO')

# %%
# ALERT BLOCCANTE SE CI SONO KO
ko_presenti = df_merged[df_merged['check_NAV_4'] == 'KO']

if not ko_presenti.empty:
    codici_ko = ko_presenti['COLLECTION_DB_CODE'].unique()
    raise ValueError(f"❌ CI SONO FONDI CON UNA DIFFERENZA SIGNIFICATIVA CHE DEVONO ESSERE VERIFICATI.\n"
                     f"COLLECTION_DB_CODE coinvolti: {', '.join(codici_ko)}")
    

PTF_MIF = df_merged[[
    'COLLECTION_DB_CODE',
    'NAV',
    'CODICE_ISIN',
    'CODICE_FONDO_DETT',
    'NAV_corretta'
]]

# dove NAV_corretta non c'è elimina riga 

# Elimina righe dove NAV_corretta è null
PTF_MIF = PTF_MIF[PTF_MIF["NAV_corretta"].notna()]


# importiamo le div già pronte dai quarterly

# quarterly_mif_geo_obb = pd.read_excel(f'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Dati quarterly MIF/Template - Quarterly Factsheet {anno} {mese}.xlsx', sheet_name = 'COUNTRY BOND')
quarterly_mif_geo_obb = pd.read_excel(f'C:/Users/Lenovo/Desktop/Mediolanum/Template - Quarterly Factsheet {anno} {mese}.xlsx', sheet_name = 'COUNTRY BOND')
# quarterly_mif_geo_az = pd.read_excel(f'G:/Analisi e Performance Prodotti/Fact Sheet New 2020/Dati quarterly MIF/Template - Quarterly Factsheet {anno} {mese}.xlsx', sheet_name = 'COUNTRY EQUITY')
quarterly_mif_geo_az = pd.read_excel(f'C:/Users/Lenovo/Desktop/Mediolanum/Template - Quarterly Factsheet {anno} {mese}.xlsx', sheet_name = 'COUNTRY EQUITY')

quarterly_mif_geo_obb = quarterly_mif_geo_obb.rename(columns={"FUND CODE": "COLLECTION_DB_CODE"})
quarterly_mif_geo_az = quarterly_mif_geo_az.rename(columns={"FUND CODE": "COLLECTION_DB_CODE"})


# facciamo poi la somma di NAV_corretta per CODICE_FONDO_DETT in PTF_MIF
# così abbiamo il totale nav per db con e senza H 
# poi facciamo il merge

NAV_per_fondo_dett = (
    PTF_MIF
    .groupby("CODICE_FONDO_DETT", as_index=False)["NAV_corretta"]
    .sum()
    .rename(columns={"NAV_corretta": "NAV_FONDO_DETT"})
    .rename(columns={"CODICE_FONDO_DETT": "COLLECTION_DB_CODE"})
)



quarterly_mif_geo_obb = quarterly_mif_geo_obb.merge(
    NAV_per_fondo_dett,
    on="COLLECTION_DB_CODE",
    how="left"
)

quarterly_mif_geo_az = quarterly_mif_geo_az.merge(
    NAV_per_fondo_dett,
    on="COLLECTION_DB_CODE",
    how="left"
)

# ora calcoliamo i NAV_per_fondo_dett univoci per COLLECTION_DB_CODE nei quarterly come NAV_MIF_OBB e NAV_MIF_AZ

NAV_MIF_OBB = (
    quarterly_mif_geo_obb[["COLLECTION_DB_CODE", "NAV_FONDO_DETT"]]
    .drop_duplicates()
    ["NAV_FONDO_DETT"]
    .sum()
)

NAV_MIF_AZ = (
    quarterly_mif_geo_az[["COLLECTION_DB_CODE", "NAV_FONDO_DETT"]]
    .drop_duplicates()
    ["NAV_FONDO_DETT"]
    .sum()
)

quarterly_mif_geo_obb["NAV_MIF_OBB"] = NAV_MIF_OBB
quarterly_mif_geo_az["NAV_MIF_AZ"] = NAV_MIF_AZ

# TOTALE_Ribasato = WEIGHT * NAV_FONDO_DETT / NAV_MIF_OBB o NAV_MIF_AZ

quarterly_mif_geo_obb["TOTALE_Ribasato"] = quarterly_mif_geo_obb['WEIGHT'] * quarterly_mif_geo_obb['NAV_FONDO_DETT'] / quarterly_mif_geo_obb['NAV_MIF_OBB']
quarterly_mif_geo_az["TOTALE_Ribasato"] = quarterly_mif_geo_az['WEIGHT'] * quarterly_mif_geo_az['NAV_FONDO_DETT'] / quarterly_mif_geo_az['NAV_MIF_AZ']

print(quarterly_mif_geo_obb["TOTALE_Ribasato"].sum())
print(quarterly_mif_geo_az["TOTALE_Ribasato"].sum())

# aggiungiamo la colonna TIPO: quarterly_mif_geo_obb -> OBB e quarterly_mif_geo_az -> AZ
# poi concateniamo in un unico df 
# i tre file (due separati) e quello concatenato devono essere salvati nello stesso excel

quarterly_mif_geo_obb["TIPO"] = "OBB"
quarterly_mif_geo_az["TIPO"] = "AZ"

# 👉 rinomina in NAV_MIF_tot
quarterly_mif_geo_obb = quarterly_mif_geo_obb.rename(columns={"NAV_MIF_OBB": "NAV_MIF_tot"})
quarterly_mif_geo_az = quarterly_mif_geo_az.rename(columns={"NAV_MIF_AZ": "NAV_MIF_tot"})


quarterly_mif_geo_totale = pd.concat(
    [quarterly_mif_geo_obb, quarterly_mif_geo_az],
    ignore_index=True
)

print(quarterly_mif_geo_totale["TOTALE_Ribasato"].sum())

# %%
# ============================================================
# 📄 INFO DATI
# ============================================================
info = pd.DataFrame({
    'Nome codice': ['STORIA_DEI_FONDI'],
    'Nome_File': ['Geo MIF Totale divisa OBB e AZ'],
    'Descrizione': [
        'Partendo dai quarterly, abbiamo creato due ptf MIF, AZ e OBB, per la geo agganciandoci i NAV dei fondi MIF per dettaglio fondo'
    ],
    # 'Percorso codice': [r'G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Region\Nuovi_dati_25_01\Region_Comitato\Valuta']
    'Percorso codice': [r'C:\Users\Lenovo\Desktop\Mediolanum']
})

# ============================================================
# 💾 EXPORT EXCEL
# ============================================================
# percorso_output = r"G:\Analisi e Performance Prodotti\PowerBi MAA\Storia dei fondi\Geo MIF Totale divisa OBB e AZ.xlsx"
percorso_output = r"C:\Users\Lenovo\Desktop\Mediolanum\Geo MIF Totale divisa OBB e AZ.xlsx"

with pd.ExcelWriter(percorso_output, engine="openpyxl") as writer:
    
    quarterly_mif_geo_obb.to_excel(
        writer,
        sheet_name="OBB",
        index=False
    )
    
    quarterly_mif_geo_az.to_excel(
        writer,
        sheet_name="AZ",
        index=False
    )
    
    quarterly_mif_geo_totale.to_excel(
        writer,
        sheet_name="TOTALE",
        index=False
    )
    
    info.to_excel(
        writer,
        sheet_name="INFO",
        index=False
    )

print(f"✅ File salvato: {percorso_output}")



# ============================================================
# 🟢 CREAZIONE STORICO QUARTERLY MIF GEO (da 2021 fino a oggi)
# ============================================================


# TODO
























