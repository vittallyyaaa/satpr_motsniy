import pandas as pd

file_path = r'D:\lab_tpr\lab1\task1.xlsx'
data = pd.read_excel(file_path)

weights = data.iloc[-1, 1:].values
alternatives = data.iloc[:-1, 1:].values
candidates = data.iloc[:-1, 0].values

scores = alternatives * weights
total_scores = scores.sum(axis=1)

for candidate, total_score in zip(candidates, total_scores):
    print(f'{candidate}: {total_score:.2f}')

best_alternative_index = total_scores.argmax()
best_lawyer = candidates[best_alternative_index]
print(f'\nНайбільш підходящий адвокат: {best_lawyer}')
