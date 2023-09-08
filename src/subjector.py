import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from difflib import SequenceMatcher


EMAILS_FILE = 'FCZ-24-30-07-2023.csv'

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

emails = pd.read_csv(EMAILS_FILE)
subjects = emails.Subject.unique()
subjects = [subject for subject in subjects if str(subject) != 'nan']
scores = {}
for i in subjects:
    scores[i] = {}
    for j in subjects:
        scores[i][j] = similar(i, j)
# print(scores)

d = scores
df = pd.DataFrame.from_dict(d, orient='index')
df.to_excel('similarity.xlsx')
# Create the heatmap using seaborn
# plt.figure(figsize=(14, 10))
# sns.heatmap(df, cmap="crest")
# plt.title('Heatmap of the Dictionary')
# plt.show()
# plt.savefig('similarity.png')

