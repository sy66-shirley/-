import pandas as pd
import re

file_path = 'C:\\Users\\16262\\Downloads\\同花顺/CR+Agent调查问卷.xlsx'
df = pd.read_excel(file_path,header=0)

df = df.iloc[0:, 2:]
word_counts = {}


for i, column in enumerate(df.columns):
    if i == 16:  
        continue
    column_word_count = {}
    for word in df[column].dropna().astype(str):
        word = re.sub(r'[，\s/、。,]+', ' ', word)
        words = word.split()  
        for w in words:
            w = w.strip()
            if w in column_word_count:
                column_word_count[w] += 1
            else:
                column_word_count[w] = 1
    word_counts[column] = column_word_count


for column, counts in word_counts.items():
    print(f"\n'{column}'：")
    for word, count in counts.items():
        print(f"{word}: {count}")
 
 
