import whois
import datetime
import pandas as pd


def load_domains_from_file(filename="domains.txt"):
    """
    Загружает список доменов из текстового файла.
    Каждая непустая строка считается доменом.
    """
    with open(filename, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]


def process_domain(domain, drop_delay=30):
    """
    Выполняет whois-запрос для домена и извлекает данные:
      - Дата регистрации: самая ранняя дата регистрации (creation_date)
      - Дата окончания: expiration_date
      - Регистрация закончится через: разница между expiration_date и текущей датой (в днях)
      - Возраст домена: количество дней с даты регистрации до текущей даты
      - Когда освободится дроп: expiration_date + drop_delay (если прошло, то "уже свободен", иначе количество дней)
      - Пометка: "можно купить" если домен свободен (expiration_date < now) или если не удалось получить whois,
          иначе "занят".
    """
    result = {
        "Домен": domain,
        "Дата регистрации": "",
        "Дата окончания": "",
        "Регистрация закончится через": "",
        "Возраст домена": "",
        "Когда освободится дроп": "",
        "Пометка": ""
    }

    try:
        w = whois.whois(domain)
    except Exception as e:
        print(f"Ошибка whois для {domain}: {e}")
        result["Пометка"] = "можно купить"
        return result

    # Обработка creation_date: если значение - список, выбираем минимальную (самую раннюю дату)
    creation_date = w.creation_date
    if isinstance(creation_date, list):
        try:
            creation_date = min([d if isinstance(d, datetime.datetime) else pd.to_datetime(d)
                                 for d in creation_date if d])
        except Exception as e:
            print(f"Ошибка при выборе самой ранней даты регистрации для {domain}: {e}")
            creation_date = None
    elif creation_date and not isinstance(creation_date, datetime.datetime):
        try:
            creation_date = pd.to_datetime(creation_date)
        except Exception as e:
            print(f"Ошибка при преобразовании даты регистрации для {domain}: {e}")
            creation_date = None

    # Обработка expiration_date: если значение - список, берем первый элемент
    expiration_date = w.expiration_date
    if isinstance(expiration_date, list):
        expiration_date = expiration_date[0]
    elif expiration_date and not isinstance(expiration_date, datetime.datetime):
        try:
            expiration_date = pd.to_datetime(expiration_date)
        except Exception as e:
            print(f"Ошибка при преобразовании даты окончания для {domain}: {e}")
            expiration_date = None

    now = datetime.datetime.now()

    # Дата регистрации и возраст домена (в днях)
    if creation_date:
        result["Дата регистрации"] = creation_date.strftime("%Y-%m-%d")
        age_days = (now - creation_date).days
        result["Возраст домена"] = age_days
    else:
        result["Дата регистрации"] = "Неизвестно"
        result["Возраст домена"] = "Неизвестно"

    # Обработка даты окончания регистрации и "Регистрация закончится через"
    if expiration_date:
        result["Дата окончания"] = expiration_date.strftime("%Y-%m-%d")
        delta_days = (expiration_date - now).days
        if delta_days >= 0:
            result["Регистрация закончится через"] = f"через {delta_days} дней"
        else:
            result["Регистрация закончится через"] = f"{-delta_days} дней назад"
    else:
        result["Дата окончания"] = "Неизвестно"
        result["Регистрация закончится через"] = "Неизвестно"

    # "Когда освободится дроп" = expiration_date + drop_delay (если expiration_date есть)
    if expiration_date:
        drop_release_date = expiration_date + datetime.timedelta(days=drop_delay)
        delta_drop = (drop_release_date - now).days
        if delta_drop > 0:
            result["Когда освободится дроп"] = f"через {delta_drop} дней"
        else:
            result["Когда освободится дроп"] = "уже свободен"
    else:
        result["Когда освободится дроп"] = "Неизвестно"

    # Пометка: если expiration_date есть и она меньше текущей даты, домен считается свободным
    if expiration_date and expiration_date < now:
        result["Пометка"] = "можно купить"
    else:
        result["Пометка"] = "занят"

    return result


if __name__ == "__main__":
    domains = load_domains_from_file("domains.txt")
    results = []
    for domain in domains:
        print(f"Обрабатываем {domain}...")
        res = process_domain(domain)
        results.append(res)

    df = pd.DataFrame(results)
    output_file = "whois_results.xlsx"
    try:
        df.to_excel(output_file, index=False)
        print(f"Результаты сохранены в {output_file}.")
    except PermissionError as pe:
        print(f"Ошибка сохранения файла {output_file}: {pe}. Проверьте, не открыт ли файл в Excel.")
