import pandas as pd
from pathlib import Path
import numpy as np


EMAILS_FILE = "FCZ-2023-08-05-07-24.xlsx"
DEMANDES_FILE = "DM-" + EMAILS_FILE
EMAILS_PATH = Path(__file__).parent.parent / "data" / EMAILS_FILE

emails = pd.read_excel(EMAILS_PATH)[["Thread", "Date", "From", "To", "Subject"]].astype(str)
emails.dropna(how='all')
emails['Temps'] = emails['Date']

def set_and_join_words(values):
    words = ' '.join(values).split()  # Split into words and join
    unique_words = ' '.join(set(words))  # Apply set and join
    return unique_words

aggregations = {
    'Date': lambda x: pd.to_datetime(x).min(),
    'Temps': lambda x: (pd.to_datetime(x).max() - pd.to_datetime(x).min()).total_seconds() / 60 / 60,
    'From': set_and_join_words,
    'To': set_and_join_words,
    'Subject': set_and_join_words,
}

demandes = emails.groupby('Thread').agg(aggregations).reset_index()
demandes['Date'] = demandes['Date'].astype(str)
# demandes.to_excel(Path(__file__).parent.parent / "data" / DEMANDES_FILE)

