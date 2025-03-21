Ниже приведено подробное описание данного скрипта с объяснением его назначения, зависимостей, входных и выходных данных, а также подробными комментариями по каждой функции:

---

## Общее назначение скрипта

Этот скрипт предназначен для автоматического получения информации о доменах с помощью запроса WHOIS. Для каждого домена из входного файла скрипт извлекает следующие данные:

1. **Дата регистрации** – самая ранняя дата регистрации домена (если домен регистрировался несколько раз, выбирается минимальная дата).
2. **Дата окончания регистрации** – дата, до которой домен зарегистрирован (если она известна).
3. **Регистрация закончится через** – разница (в днях) между датой окончания регистрации и текущей датой. Если регистрация уже истекла, указывается, сколько дней назад.
4. **Возраст домена** – количество дней от даты первой регистрации до текущей даты.
5. **Когда освободится дроп** – рассчитывается как дата окончания регистрации плюс заданное число дней (по умолчанию 30). Если эта дата уже прошла, выводится «уже свободен», иначе – сколько дней осталось.
6. **Пометка** – если дата окончания регистрации уже в прошлом или если не удалось выполнить whois-запрос, домен считается свободным (выводится «можно купить»), иначе – домен считается занятым (выводится «занят»).

Результаты сохраняются в Excel‑файл, где каждая строка содержит информацию для одного домена.

---

## Зависимости и установка

Скрипт использует следующие библиотеки:

- **python-whois** – для выполнения whois-запросов.  
  Установить можно командой:  
  ```
  pip install python-whois
  ```

- **pandas** – для работы с данными и экспорта в Excel.  
  Установить можно командой:  
  ```
  pip install pandas
  ```

- **datetime** – стандартный модуль Python для работы с датами и временем.

---

## Структура кода

### 1. Функция `load_domains_from_file`

```python
def load_domains_from_file(filename="domains.txt"):
    """
    Загружает список доменов из текстового файла.
    Каждая непустая строка считается доменом.
    """
    with open(filename, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]
```

**Назначение:**  
Эта функция читает файл (по умолчанию `domains.txt`) и возвращает список доменов. Каждая строка, не являющаяся пустой, считается доменом.

**Как использовать:**  
Поместите список доменов (по одному домену в строке) в файл `domains.txt` в той же директории, что и скрипт.

---

### 2. Функция `process_domain`

```python
def process_domain(domain, drop_delay=30):
    """
    Выполняет whois-запрос для домена и извлекает данные:
      - Дата регистрации: самая ранняя дата регистрации (creation_date)
      - Дата окончания: expiration_date
      - Регистрация закончится через: разница между expiration_date и текущей датой (в днях)
      - Возраст домена: количество дней с даты регистрации до текущей даты
      - Когда освободится дроп: expiration_date + drop_delay (если прошло, то "уже свободен", иначе количество дней)
      - Пометка: "можно купить", если домен свободен (expiration_date < текущей даты) или если не удалось выполнить whois-запрос, иначе "занят".
    Параметры:
      domain (str): домен для проверки.
      drop_delay (int): количество дней после expiration_date, по истечении которых домен считается свободным для покупки.
    Возвращает:
      Словарь с ключами:
        "Домен", "Дата регистрации", "Дата окончания",
        "Регистрация закончится через", "Возраст домена",
        "Когда освободится дроп", "Пометка".
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
```

**Обработка запроса WHOIS:**  
В блоке `try` выполняется whois-запрос к домену. Если запрос завершается с ошибкой, скрипт выводит сообщение об ошибке и помечает домен как «можно купить».

**Обработка `creation_date`:**  
Если поле `creation_date` возвращается в виде списка (что бывает, если домен регистрировался несколько раз), выбирается самая ранняя дата с помощью функции `min()`. Если значение – не объект типа `datetime`, оно пытается преобразоваться через `pd.to_datetime`.

**Обработка `expiration_date`:**  
Если поле `expiration_date` возвращается как список, выбирается первый элемент. При необходимости также производится преобразование в `datetime`.

**Расчет возраста домена:**  
Возраст домена теперь рассчитывается как количество дней между текущей датой и датой первой регистрации. Результат сохраняется в столбце «Возраст домена».

**Расчет времени до окончания регистрации:**  
Разница между датой окончания регистрации и текущей датой выводится в столбце «Регистрация закончится через». Если регистрация истекла, выводится сколько дней назад.

**Расчет "Когда освободится дроп":**  
Добавляется период (по умолчанию 30 дней) к дате окончания регистрации. Если полученная дата уже прошла, выводится «уже свободен», иначе – сколько дней осталось до освобождения дропа.

**Пометка:**  
Если дата окончания регистрации уже в прошлом, домен считается свободным ("можно купить"), иначе – "занят".

---

### 3. Основной блок

```python
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
```

**Основной блок выполняет следующие действия:**

- **Загрузка доменов:**  
  Функция `load_domains_from_file` считывает список доменов из файла `domains.txt`.

- **Обработка каждого домена:**  
  Для каждого домена вызывается функция `process_domain`, которая выполняет whois-запрос и обрабатывает данные. Результаты собираются в список.

- **Создание DataFrame:**  
  Список результатов преобразуется в DataFrame с помощью библиотеки pandas.

- **Экспорт в Excel:**  
  DataFrame записывается в Excel‑файл `whois_results.xlsx`. Если файл открыт в Excel, возникает ошибка PermissionError, и выводится сообщение об этом.

---

## Инструкции по использованию

1. **Установите зависимости:**  
   Выполните в командной строке:
   ```bash
   pip install python-whois pandas
   ```

2. **Подготовьте файл с доменами:**  
   Создайте файл `domains.txt` в той же папке, где находится скрипт. Каждая строка должна содержать один домен (например, `example.com`).

3. **Запустите скрипт:**  
   Запустите скрипт (например, через IDE или командную строку). Скрипт обработает каждый домен, выведет в консоль ход обработки и создаст файл `whois_results.xlsx` с итоговыми данными.

4. **Проверьте Excel-файл:**  
   Откройте созданный файл `whois_results.xlsx` в Excel. Таблица будет содержать следующие столбцы:
   - **Домен**
   - **Дата регистрации** – первая (самая ранняя) дата регистрации.
   - **Дата окончания** – дата, до которой домен зарегистрирован.
   - **Регистрация закончится через** – сколько дней осталось до окончания регистрации или сколько дней прошло с момента её окончания.
   - **Возраст домена** – общее количество дней с момента первой регистрации до текущей даты.
   - **Когда освободится дроп** – разница между (expiration_date + drop_delay) и текущей датой; если дата уже прошла, выводится "уже свободен".
   - **Пометка** – "можно купить", если домен свободен (т.е. срок регистрации истёк или не удалось получить данные WHOIS), иначе "занят".

---

## Возможные доработки

- **Более точное вычисление возраста:**  
  Если потребуется учитывать месяцы и дни для более точного расчета, можно использовать другие методы вычисления разницы между датами.

- **Обработка нескольких whois-серверов:**  
  Если стандартный whois-запрос не возвращает нужные данные, можно добавить альтернативные источники.

- **Дополнительные проверки:**  
  Можно добавить проверки на корректность получаемых дат, логирование ошибок в отдельный файл, а также добавить поддержку командной строки для передачи имени файла с доменами.

---

Это подробное описание должно помочь вам понять, как работает скрипт, какие данные он использует, и как его можно настроить или доработать в будущем. Если будут вопросы или понадобится дополнительная функциональность, дайте знать!
