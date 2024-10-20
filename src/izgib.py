import math as mt
import numpy as np
import xlwings as xw


# Получение исходных данных из активного Excel файла
def extract_excel_to_map(sheet):
    # Извлекаем данные начиная с 4-й строки
    params = sheet.range("A4:B100").value  # Получаем диапазон данных

    # Преобразуем первый столбец в ключи, второй в значения (без пустых строк)
    data_map = {row[0]: row[1] for row in params if row[0] is not None}

    # Извлечение массивов площадей и координат стрингеров
    count = int(data_map['m_all']) + 4
    Fi = sheet.range(f"G5:G{count}").value
    yi_ = sheet.range(f"H5:H{count}").value
    data_map['Fi'] = np.array(Fi, dtype=float)
    data_map['yi_'] = np.array(yi_, dtype=float)

    return data_map


# Расчёт постоянных параметров
def calc_const_params(data):
    data['sigma_kr'] = pow(mt.pi, 2) * data['E'] * data['Jxx'] / (pow(data['l'], 2) * data['F'])
    data['sigmakr'] = np.full(10, data['sigma_kr'])
    data['sigmakr'] = np.append(data['sigmakr'], [data['sigma_t']] * 10)


# Расчёт переменных параметров
def calc_params(const_data, data):
    data['Fi_pr'] = data['fi'] * const_data['Fi']
    Fi_pr_yi_ = data['Fi_pr'] * const_data['yi_']
    data['yt'] = sum(Fi_pr_yi_) / sum(data['Fi_pr'])
    data['yi'] = const_data['yi_'] - data['yt']
    Fi_pr_yi2 = data['Fi_pr'] * data['yi'] * data['yi']
    data['Jx_pr'] = sum(Fi_pr_yi2)
    data['sigmai'] = -data['fi'] * const_data['Mx'] / data['Jx_pr'] * data['yi']

    for i in range(int(const_data['m_all'])):
        if const_data['sigmakr'][i] > abs(data['sigmai'][i]):
            data['fi'][i] = 1
        else:
            data['fi'][i] = const_data['sigmakr'][i] / abs(data['sigmai'][i])


# Проверка сходимости
def check_convergence(old_data, new_data, tolerance):
    error = np.abs(new_data - old_data) / np.abs(old_data)
    return np.all(error < tolerance)


# Итерационный процесс
def calc_iter(const_data, tolerance=0.01, max_iterations=100):
    iteration = 0
    converged = False

    data = dict(fi=np.full(int(const_data['m_all']), float(1)))

    while not converged and iteration < max_iterations:
        # Сохраняем старые данные для сравнения
        old_data = data['fi'].copy()
        calc_params(const_data, data)
        converged = check_convergence(old_data, data['fi'], tolerance)
        iteration += 1

    if converged:
        print(f"Процесс сошелся за {iteration} итераций.")
    else:
        print(f"Максимальное количество итераций {max_iterations} достигнуто.")

    return data


# Основная функция для запуска в Excel
def main():
    wb = xw.Book.caller()  # Открываем активную книгу
    sheet1 = wb.sheets['data']  # Активный лист
    sheet2 = wb.sheets['results']  # Лист для записи результатов

    # Извлекаем константные данные из таблицы
    const_data = extract_excel_to_map(sheet1)

    # Вычисляем постоянные параметры
    calc_const_params(const_data)

    # Выполняем итерационные расчеты
    data = calc_iter(const_data, const_data['Epsilon'], 100000)

    # Запись результатов обратно в Excel
    # sheet2.range("B2").value = f"Mx: {const_data['Mx']}, yt: {data['yt']}"
    #sheet2.range("B2:B21").value = data['fi']  # Пример записи fi в колонку C
    sheet2.range("B2").options(transpose=True).value = data['fi']


#Эта функция запускается по нажатию кнопки в Excel
if __name__ == "__main__":
    xw.Book("data.xlsx").set_mock_caller()  # Имитирует вызов из Excel для тестов
    main()