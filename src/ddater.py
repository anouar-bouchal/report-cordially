import pandas as pd
from pathlib import Path


EMAILS_FILE = Path(__file__).parent.parent / "data" / "FCZ-TT.xlsx"
DEMANDES_FILE = Path(__file__).parent.parent / "demandes-v1.xlsx"

emails = pd.read_excel(EMAILS_FILE)[['Thread', 'Date']]
aggregations = {
    'Date': lambda x: pd.to_datetime(x).min(),
} 
dated_demandes = emails.groupby("Thread").aggregate(aggregations)
# Convert 'Temps' column to datetime
dated_demandes['Date'] = pd.to_datetime(dated_demandes['Date'], utc=True).dt.tz_localize(None)
# dated_demandes['Date'].dt.tz_localize(None)

# Extract day and time components
dated_demandes['Annee'] = dated_demandes['Date'].dt.strftime('%Y')
dated_demandes['Mois'] = dated_demandes['Date'].dt.strftime('%m')
dated_demandes['Jour'] = dated_demandes['Date'].dt.strftime('%d')
dated_demandes['Hour'] = dated_demandes['Date'].dt.strftime('%H:%M')

demandes = pd.read_excel(DEMANDES_FILE, sheet_name='All').dropna(how='all')
informations = pd.read_excel(DEMANDES_FILE, sheet_name='Information').dropna(how='all')
annulations = pd.read_excel(DEMANDES_FILE, sheet_name='Annulation').dropna(how='all')
erreurs = pd.read_excel(DEMANDES_FILE, sheet_name='Erreur').dropna(how='all')
activations = pd.read_excel(DEMANDES_FILE, sheet_name='Activation').dropna(how='all')
creations = pd.read_excel(DEMANDES_FILE, sheet_name='Creation').dropna(how='all')
modifications = pd.read_excel(DEMANDES_FILE, sheet_name='Modification').dropna(how='all')
extractions = pd.read_excel(DEMANDES_FILE, sheet_name='Extraction').dropna(how='all')

dated_demandes = pd.merge(demandes, dated_demandes, left_on='Thread', right_on='Thread', how='left', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
dated_informations = pd.merge(informations, dated_demandes, left_on='Thread', right_on='Thread', how='left', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
dated_annulations = pd.merge(annulations, dated_demandes, left_on='Thread', right_on='Thread', how='left', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
dated_erreurs = pd.merge(erreurs, dated_demandes, left_on='Thread', right_on='Thread', how='left', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
dated_activations = pd.merge(activations, dated_demandes, left_on='Thread', right_on='Thread', how='left', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
dated_creations = pd.merge(creations, dated_demandes, left_on='Thread', right_on='Thread', how='left', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
dated_modifications = pd.merge(modifications, dated_demandes, left_on='Thread', right_on='Thread', how='left', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
dated_extractions = pd.merge(extractions, dated_demandes, left_on='Thread', right_on='Thread', how='left', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')


# writer = pd.ExcelWriter("demandes-v2.xlsx")
# dated_demandes.to_excel(writer, sheet_name='All')
# dated_informations.to_excel(writer, sheet_name='Information')
# dated_annulations.to_excel(writer, sheet_name='Annulation')
# dated_erreurs.to_excel(writer, sheet_name='Erreur')
# dated_activations.to_excel(writer, sheet_name='Activation')
# dated_creations.to_excel(writer, sheet_name='Creation')
# dated_modifications.to_excel(writer, sheet_name='Modification')
# dated_extractions.to_excel(writer, sheet_name='Extraction')
# writer.close()
print(dated_demandes)
