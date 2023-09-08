import pandas as pd
import numpy as np
from pathlib import Path


DEMANDES_FILE = 'demandes-v2.xlsx'

new_demandes = pd.read_excel(Path(__file__).parent.parent / "data" / "DM-FCZ-2023-08-05-07-24.xlsx")

# Incorporate data
demandes = pd.read_excel(DEMANDES_FILE, sheet_name='All').dropna(how='all')
informations = pd.read_excel(DEMANDES_FILE, sheet_name='Information').dropna(how='all')
annulations = pd.read_excel(DEMANDES_FILE, sheet_name='Annulation').dropna(how='all')
erreurs = pd.read_excel(DEMANDES_FILE, sheet_name='Erreur').dropna(how='all')
activations = pd.read_excel(DEMANDES_FILE, sheet_name='Activation').dropna(how='all')
creations = pd.read_excel(DEMANDES_FILE, sheet_name='Creation').dropna(how='all')
modifications = pd.read_excel(DEMANDES_FILE, sheet_name='Modification').dropna(how='all')
extractions = pd.read_excel(DEMANDES_FILE, sheet_name='Extraction').dropna(how='all')
alls = pd.concat([informations, annulations, erreurs, activations, creations, modifications, extractions], join='outer', ignore_index=True)

merged = pd.merge(alls, new_demandes, how='outer', on=['Thread'], suffixes=('_x', '_y'))
# merged.to_excel("merged.xlsx")
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
merged['Temps_x'].fillna(0, inplace=True)
merged['Temps_y'].fillna(0, inplace=True)
def update_column(row, x, y):
    if pd.isna(row[x]) or row[x] == '':
        return row[y]
    return row[x]
for column in ["Temps", "From", "To", "Subject", "Date"]:
    merged[column] = np.nan
    merged[column] = merged.apply(update_column, args=(column + "_x", column + "_y"), axis=1)
merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
final_merged = merged[["Thread", "Type", "Objet", "Specifique", "Action", "Status", "Application", "Temps", "From", "Demandeur", "Traitant", "To", "Subject", "Date"]]
final_merged.to_excel("DM-TT-2023-08-05.xlsx")
