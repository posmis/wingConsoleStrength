import pandas as pd
import math as mt
import numpy as np

# Получение исходных данных
def extract_excel_to_map(file_path, sheet_name):
    # Загружаем Excel файл с помощью pandas
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # Извлекаем данные начиная с 4-й строки
    params = df.iloc[3:, [0, 1]]  # Используем срез начиная с индекса 3 (4-я строка)
    # Преобразуем первый столбец в ключи, второй в значения
    data_map = pd.Series(params[1].values, index=params[0]).to_dict()

    # Извлечение массивов площадей и координат стрингеров
    count = int(data_map['m_all']) + 4
    Fi = df.iloc[4:count, 6].values
    yi_ = df.iloc[4:count, 7].values
    data_map['Fi'] = Fi
    data_map['yi_'] = yi_

    return data_map

def calc_const_params(data):
    data['sigma_kr'] = pow(mt.pi, 2) * data['E'] * data['Jxx'] / (pow(data['l'], 2) * data['F'])
    data['sigmakr'] = np.full(10, data['sigma_kr'])
    data['sigmakr'] = np.append(data['sigmakr'], [data['sigma_t']] * 10)

def calc_params(const_data, data):
    data['Fi_pr'] = data['fi'] * const_data['Fi']
    Fi_pr_yi_ = data['Fi_pr'] * const_data['yi_']
    data['yt'] = sum(Fi_pr_yi_) / sum(data['Fi_pr'])
    data['yi'] = const_data['yi_'] - data['yt']
    Fi_pr_yi2 = data['Fi_pr'] * data['yi'] * data['yi']
    data['Jx_pr'] = sum(Fi_pr_yi2)
    data['sigmai'] = -data['fi'] * const_data['Mx'] / data['Jx_pr'] * data['yi']

    for i in range(const_data['m_all']):
        if const_data['sigmakr'][i] > abs(data['sigmai'][i]):
            data['fi'][i] = 1
        else:
            data['fi'][i] =const_data['sigmakr'][i] / abs(data['sigmai'][i])

def check_convergence(old_data, new_data, tolerance):
    error = np.abs(new_data - old_data) / np.abs(old_data)
    return np.all(error < tolerance)

def calc_iter(const_data, tolerance=0.01, max_iterations=100):
    iteration = 0
    converged = False

    data = dict(fi=np.full(const_data['m_all'], float(1)))

    while not converged and iteration < max_iterations:
        # Сохраняем старые данные для сравнения
        old_data = data['fi'].copy()
        calc_params(const_data, data)
        converged = check_convergence(old_data, data['fi'], tolerance)
        iteration += 1

    if converged:
        # print(old_data)
        print(f"Процесс сошелся за {iteration} итераций.")
    else:
        print(f"Максимальное количество итераций {max_iterations} достигнуто.")

    return data

# Главная функция программы
def main():
    file_path = "data.xlsx"
    sheet_name = "data"
    const_data = extract_excel_to_map(file_path, sheet_name)
    calc_const_params(const_data)
    data = calc_iter(const_data, const_data['Epsilon'], 100000)
    print(f"Mx: {const_data['Mx']} | Epsilon: {const_data['Epsilon']} | yt: {data['yt']}")
    print(data['fi'])


# Вызов основной функции
if __name__ == "__main__":
    main()