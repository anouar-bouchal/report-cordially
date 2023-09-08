import pandas as pd
from pathlib import Path


EMAILS_PATH = Path(__file__).parent.parent / "data" / "FCZ-TT.xlsx"
emails = pd.read_excel(EMAILS_PATH)
emails.dropna(how="all")

# emails['Date'] = pd.to_datetime(emails['Date'], utc=True)
demandes = emails[["Thread", "Date", "From", "Subject", "General", "Specifique"]].astype(str)

def set_and_join_words(values):
    words = ' '.join(values).split()  # Split into words and join
    unique_words = ' '.join(set(words))  # Apply set and join
    return unique_words


aggregations = {
        'Date': lambda x: (pd.to_datetime(x).max() - pd.to_datetime(x).min()).total_seconds() / 60,
        'From': set_and_join_words,
        'Subject': set_and_join_words,
        'General': set_and_join_words,
        'Specifique': set_and_join_words
        }

analytic_demandes = demandes.groupby('Thread').agg(aggregations).reset_index()

analytic_demandes.to_excel("analytic_demandes.xlsx")
