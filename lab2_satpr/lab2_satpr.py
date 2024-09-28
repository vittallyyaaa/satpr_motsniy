import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# Зчитуємо матриці з Excel
def read_matrix(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name, header=None)
    return df.to_numpy()


# Обчислюэмо Wi
def calculate_wi(matrix):
    wi = np.prod(matrix, axis=1) ** (1 / matrix.shape[1])  # Степінь 1/кількість стовпців
    return wi


# Нормалізуємо Wi
def normalize_wi(wi_values):
    wi_sum = sum(wi_values)
    wnorm_values = wi_values / wi_sum  # Нормалізуємо значення W_i
    return wnorm_values


# Нормалізуємо матрицю альтернатив
def normalize_matrix(matrix):
    col_sums = np.sum(matrix, axis=0)
    normalized_matrix = matrix / col_sums
    priorities = np.mean(normalized_matrix, axis=1)
    return normalized_matrix, priorities


# Обчислюємо пріоритетів альтернатив
def calculate_final_priorities(criteria_priorities, file_path, num_alternatives=3, num_criteria=9):
    final_priorities = np.zeros(num_alternatives)

    for i in range(1, num_criteria + 1):
        sheet_name = f'K{i}'
        alternative_matrix = read_matrix(file_path, sheet_name)

        _, priorities = normalize_matrix(alternative_matrix)
        final_priorities += criteria_priorities[i - 1] * priorities

    return final_priorities


# Виводим гарну матрицю
def print_matrix_with_sums(matrix, wi_values, wnorm_values, labels, col_sums):
    headers = [f"E{i + 1}" for i in range(matrix.shape[1])] + ["Wi", "Wнорм"]
    header_row = "{:>5}".format("") + " ".join(f"{header:>8}" for header in headers)
    print(header_row)

    for i, row in enumerate(matrix):
        formatted_row = " ".join(f"{x:>8.3f}" for x in row)
        print(f"{labels[i]:<5} {formatted_row} {wi_values[i]:>8.3f} {wnorm_values[i]:>8.3f}")

    formatted_sums = " ".join(f"{x:>8.3f}" for x in col_sums)
    print(f"Σ     {formatted_sums} {sum(wi_values):>8.3f} {sum(wnorm_values):>8.3f}")


# Візуалізація через плот
def visualize_final_priorities(final_priorities):
    labels = ['A1', 'A2', 'A3']
    colors = ['green', 'orange', 'blue']
    plt.bar(labels, final_priorities, color=colors)

    plt.xlabel('Альтернативи')
    plt.ylabel('Пріоритети')
    plt.title('Пріоритети альтернатив')
    plt.show()


# Основна програма
def main():
    file_path = 'xls/data.xlsx'

    # Зчитуємо матрицю критеріїв
    criteria_matrix = read_matrix(file_path, 'criteria_matrix')

    # Обчислюємо W_i та Wнорм
    wi_values = calculate_wi(criteria_matrix)
    wnorm_values = normalize_wi(wi_values)

    # Підписи для рядків
    row_labels = [f"E{i + 1}" for i in range(len(criteria_matrix))]

    # Виводимо матрицю критеріїв
    print_matrix_with_sums(criteria_matrix, wi_values, wnorm_values, row_labels, np.sum(criteria_matrix, axis=0))

    # Обчислюємо пріоритети альтернатив
    final_priorities = calculate_final_priorities(wnorm_values, file_path)

    # Виводимо остаточні пріоритети альтернатив
    print("\nОстаточні пріоритети альтернатив:")
    for i, priority in enumerate(final_priorities):
        print(f"a{i + 1}: {priority:.3f}")

    # Плот
    visualize_final_priorities(final_priorities)

    best_alternative_index = np.argmax(final_priorities)
    best_alternative = f"A{best_alternative_index + 1}"

    # Виводимо найбільш підходящу альтернативу
    print(f"\n{'=' * 30}")
    print(f"Найбільш підходяща альтернатива: {best_alternative}")
    print(f"{'=' * 30}")


if __name__ == '__main__':
    main()
