import pandas as pd
import numpy as np
import re
import sys
from pathlib import Path

EMAILS_PATH = Path(__file__).parent.parent
EMAILS_FILE = EMAILS_PATH / "demandes.xlsx"

def emails_finder(line):
    emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', line)
    return " | ".join(set(emails))


if __name__ == "__main__":
    xls = pd.ExcelFile(EMAILS_FILE)
    writer = pd.ExcelWriter("demandes-v1.xlsx")
    for sheet in xls.sheet_names:
        if sheet == "Equipe":
            continue
        demandes = pd.read_excel(EMAILS_FILE, sheet_name=sheet)
        demandes["Expediteurs"] = np.nan
        demandes["Destinataires"] = np.nan

        demandes['From'] = demandes['From'].astype(str)
        demandes['To'] = demandes['To'].astype(str)
        demandes['Expediteurs'] = demandes['From'].apply(emails_finder)
        demandes['Destinataires'] = demandes['To'].apply(emails_finder)
        print(demandes)
        demandes.to_excel(writer, sheet_name=sheet)
    writer.close()




   # emails_file = 'FCZ-24-30-07-2023'
    # emails = pd.read_csv(emails_file + '.csv')
    # demandes = emails[:]
    # emails.to_excel(emails_file + '.xlsx')

    # demandes = pd.read_excel(emails_file)
    # print(demandes)
    # senders = demandes.From.unique()
    # senders = pd.DataFrame(senders)
    # senders.to_excel("Sendersa.xlsx")

    # tos = list(demandes.To.unique())
    # receivers = [re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', receiver) for receiver in tos]
    # receivers = set([item for sublist in receivers for item in sublist])
    # print(len(receivers))
    # receivers = pd.DataFrame(receivers)
    # receivers.to_excel('Receiversa.xlsx')


# main()
