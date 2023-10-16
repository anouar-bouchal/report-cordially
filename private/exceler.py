import pandas as pd
from pathlib import Path


emails = pd.read_csv(Path(__file__).parent / 'FCZ-2023-08-05-07-24.csv')
emails.to_excel(Path(__file__).parent / 'FCZ-2023-08-05-07-24.xlsx')

