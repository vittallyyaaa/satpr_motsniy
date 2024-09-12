import pandas as pd

file_path = r'D:\lab_tpr\lab1\task2.xlsx'
data = pd.read_excel(file_path)

weights = data.iloc[-1, 1:].values.astype(float)
alternatives = data.iloc[:-1, 1:].apply(pd.to_numeric, errors='coerce').values
candidates = data.iloc[:-1, 0].values

print("Ваги:")
print(weights)
print("\nАльтернативи:")
print(alternatives)
print("\nКандидати:")
print(candidates)


def normalize(column, is_maximization):
    min_val = column.min()
    max_val = column.max()
    if max_val == min_val:
        return column
    if is_maximization:
        return (column - min_val) / (max_val - min_val)
    else:
        return (max_val - column) / (max_val - min_val)


is_maximization = [True, False, True, True, True]

normalized_alternatives = pd.DataFrame()
for i in range(alternatives.shape[1]):
    normalized_alternatives[i] = normalize(alternatives[:, i], is_maximization[i])

print("\nНормалізовані альтернативи:")
print(normalized_alternatives)

scores = normalized_alternatives.dot(weights)
total_scores = scores.values

print("\nФункції корисності:")
for candidate, total_score in zip(candidates, total_scores):
    print(f'{candidate}: {total_score:.2f}')

best_alternative_index = total_scores.argmax()
best_lawyer = candidates[best_alternative_index]
print(f'\nНайбільш підходящий адвокат: {best_lawyer}')
