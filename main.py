from statistics import get_statistics
from vacancy import get_vacancies

def stat_vac():
    """Выбор типа анализа данных из csv-файла
    """
    main_input_request = "Выберите тип вывода: "
    main_input_info = input(main_input_request)
    if main_input_info != "Вакансии" and main_input_info != "Вакансии":
        print("Введён неправильный тип вывода")
        return
    if main_input_info == "Навык":
        get_vacancies()
    else:
        get_statistics()

if __name__ == '__main__':
    stat_vac()
