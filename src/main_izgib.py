import sys
import pandas as pd
import math as mt
import numpy as np

def read_file_path():
    if len(sys.argv) < 2:
        print("Пожалуйста, укажите путь к Excel-файлу.")
        sys.exit(1)
    file_path = sys.argv[1]
    return file_path

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
    si = df.iloc[4:count, 12].values
    ri = df.iloc[4:count, 13].values
    data_map['Fi'] = Fi
    data_map['yi_'] = yi_
    data_map['si'] = si
    data_map['ri'] = ri

    return data_map

def calc_const_params_izgib(data):
    data['sigma_kr'] = pow(mt.pi, 2) * data['E'] * data['Jxx'] / (pow(data['l'], 2) * data['F'])
    data['sigmakr'] = np.full(10, data['sigma_kr'])
    data['sigmakr'] = np.append(data['sigmakr'], [data['sigma_t']] * 10)

def calc_params_izgib(const_data, data):
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

def check_convergence_izgib(old_data, new_data, tolerance):
    error = np.abs(new_data - old_data) / np.abs(old_data)
    return np.all(error < tolerance)

def calc_iter_izgib(const_data, tolerance=0.01, max_iterations=100):
    iteration = 0
    converged = False

    data = dict(fi=np.full(const_data['m_all'], float(1)))

    while not converged and iteration < max_iterations:
        # Сохраняем старые данные для сравнения
        old_data = data['fi'].copy()
        calc_params_izgib(const_data, data)
        converged = check_convergence_izgib(old_data, data['fi'], tolerance)
        iteration += 1

    if converged:
        # print(old_data)
        print(f"Процесс сошелся за {iteration} итераций.")
    else:
        print(f"Максимальное количество итераций {max_iterations} достигнуто.")

    return data


def calc_params_sdvig(const_data, data):
    Fi_pr_yi_ = data['Fi_pr'] * data['yi']
    data['Sx_pr'] = [sum(Fi_pr_yi_[:i+1]) for i in range(const_data['m_all'])]
    data['q_oi'] = [const_data['Qy'] * sx / data['Jx_pr'] for sx in data['Sx_pr']]
    data['a_mn'] = [si / (const_data['G'] * const_data['delta']) for si in const_data['si']]
    qoi_si = data['q_oi'] * const_data['si']
    data['a_m0'] = [value / (const_data['G'] * const_data['delta']) for value in qoi_si]

    data['a11'] = sum(data['a_mn'])
    data['a12'] = sum(data['a_mn'][4:15])
    data['a21'] = data['a12']
    data['a22'] = sum(data['a_mn'][4:15]) + data['a_mn'][const_data['m_all'] - 1]
    data['a10'] = sum(data['a_m0'])
    data['a20'] = sum(data['a_m0'][4:15])

    data['qI'] = (data['a12'] * data['a20'] / data['a22'] - data['a10']) / (data['a11'] - data['a12'] * data['a21'] / data['a22'])
    data['qII'] = - (data['a21'] * data['qI'] + data['a20']) / data['a22']
    data['qi'] = []
    for i in range(0, 4):
        data['qi'].append(data['q_oi'][i] + data['qI'])
    for i in range(4, 15):
        data['qi'].append(data['q_oi'][i] + data['qI'] + data['qII'])
    for i in range(15, 20):
        data['qi'].append(data['q_oi'][i] + data['qI'])


def calc_params_centr_z(const_data, data):
    si_ri = const_data['si'] * const_data['ri']
    data['wi'] = [value / 2 for value in si_ri]
    qoi_wi = [q * w for q, w in zip(data['q_oi'], data['wi'])]
    data['Mq'] = 2 * (data['qI'] * const_data['w1'] + data['qII'] * const_data['w2'] + sum(qoi_wi))
    data['x_z'] = data['Mq'] / const_data['Qy']

def calc_params_kr(const_data, data):
    data['Mkr'] = const_data['Mz'] + const_data['Qy'] * data['x_z']

def results_output(const_data, data):
    print("|Расчет на изгиб|")
    print(f"Mx: {const_data['Mx']} | Epsilon: {const_data['Epsilon']} | yt: {data['yt']}")
    for index, value in enumerate(data['fi']):
        print(f"Стрингер {index + 1:<{10}}| {value}")

    print("|Расчет на сдвиг|")
    print(f"a11: {data['a11']} | a12=a21: {data['a12']} | a22: {data['a22']}")
    print(f"a10: {data['a10']} | a20: {data['a20']}")
    print(f"qI: {data['qI']} | qII: {data['qII']}")
    for index, value in enumerate(data['qi']):
        if index < const_data['m_all'] - 1:
            print(f"Панель {index + 2:<{10}}| {value}")
        else:
            print(f"Панель {1:<{10}}| {value}")

    print("|Расчет. Центр жесткости|")
    print(f"x_z: {data['x_z']}")
    for index, value in enumerate(data['wi']):
        if index < const_data['m_all'] - 1:
            print(f"Панель {index + 2:<{10}}| {value}")
        else:
            print(f"Панель {1:<{10}}| {value}")

    print("|Расчет. Кручение|")
    print(f"Mkr: {data['Mkr']}")


# Главная функция программы
def main():
    file_path = read_file_path()
    sheet_name = "data"
    const_data = extract_excel_to_map(file_path, sheet_name)
    calc_const_params_izgib(const_data)
    data = calc_iter_izgib(const_data, const_data['Epsilon'], 100000)
    calc_params_sdvig(const_data, data)
    calc_params_centr_z(const_data, data)
    calc_params_kr(const_data, data)

    results_output(const_data, data)

# Вызов основной функции
if __name__ == "__main__":
    main()