import pandas as pd
import numpy as np
from pathlib import Path


# DEMANDES_FILE = 'DM-FCZ-2023-09-17-08-05.xlsx'
old_demandes = pd.read_excel('final_merged.xlsx')
new_demandes = pd.read_excel("DM-EMAILS-2023-09-24-09-16.xlsx")

# Incorporate data
# demandes = pd.read_excel(DEMANDES_FILE, sheet_name='All').dropna(how='all')
# informations = pd.read_excel(DEMANDES_FILE, sheet_name='Information').dropna(how='all')
# annulations = pd.read_excel(DEMANDES_FILE, sheet_name='Annulation').dropna(how='all')
# erreurs = pd.read_excel(DEMANDES_FILE, sheet_name='Erreur').dropna(how='all')
# activations = pd.read_excel(DEMANDES_FILE, sheet_name='Activation').dropna(how='all')
# creations = pd.read_excel(DEMANDES_FILE, sheet_name='Creation').dropna(how='all')
# modifications = pd.read_excel(DEMANDES_FILE, sheet_name='Modification').dropna(how='all')
# extractions = pd.read_excel(DEMANDES_FILE, sheet_name='Extraction').dropna(how='all')
# alls = pd.concat([informations, annulations, erreurs, activations, creations, modifications, extractions], join='outer', ignore_index=True)

merged = pd.merge(old_demandes, new_demandes, how='outer', on=['Thread'], suffixes=('_x', '_y'))
# merged.to_excel("merged.xlsx")
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
# merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
def update_column(row, x, y):
    if pd.isna(row[x]) or row[x] == '':
        return row[y]
    return row[x]
for column in ["From", "To", "Subject", "Date"]:
    merged[column] = np.nan
    merged[column] = merged.apply(update_column, args=(column + "_x", column + "_y"), axis=1)
merged['Temps_x'].fillna(0, inplace=True)
merged['Temps_y'].fillna(0, inplace=True)
merged['Temps'] = merged['Temps_x'] + merged['Temps_y']
final_merged = merged[["Thread", "Type", "Objet", "Specifique", "Action", "Status", "Application", "Temps", "From", "Demandeur", "Traitant", "To", "Subject", "Date"]]
# final_merged.to_excel("DM-TT-2023-09-17.xlsx")
