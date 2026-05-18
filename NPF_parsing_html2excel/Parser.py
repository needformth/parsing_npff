from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import requests
from datetime import datetime
import re

class HtmlToExcel:
    def __init__(self,
                 excel_path: str='result.xlsx', 
                 sber_url: str='https://npfsberbanka.ru/about/information-to-be-disclosed/results/',
                 aviapolis_url: str='https://www.npf-aviapolis.ru/index.files/Page025.html'
                 ):
        self.excel_path = excel_path
        self.sber_url = sber_url
        self.aviapolis_url = aviapolis_url

    # ===== Сбербанк =====
    sber_mapping = {
        # -- Пенсионные резервы (НПО, ПДС) --
        "пенсионныерезервынаконецмесяцавсеговтомчисле": "Пенсрез_",
        "страховойрезерв": "Пенсрез_страховой резерв",
        "финансовыйрезультатотразмещенияпенсионныхрезервовнераспределеннаяприбыльнаконецотчетногопериода": "Пенсрез_финрез",
        "резервыпокрытияпенсионныхобязательствподоговорамдолгосрочныхсбережений": "Пенсрез_ДДС",
        "количествовкладчиковюридическихлиц": "Пенсрез_вкладчики_ЮЛ",
        "количествовкладчиковфизическихлицподоговорамнпо": "Пенсрез_вкладчики_ФЛ",
        "количествоучастниковфондаподействующимдоговорамнаконецотчетногопериода": "Пенсрез_участники",
        "количествоучастниковфондаподоговорамдолгосрочныхсбереженийнаконецотчетногопериода": "Пенсрез_участники ДДС",
        "количествоучастниковфондаполучающихрегулярныевыплатыподоговорамнпонаконецотчетногопериода": "Пенсрез_участники пенсионеры",
        "количествоучастниковполучающихрегулярныевыплатыподоговорамдолгосрочныхсбереженийнаконецотчетногопериода": "Пенсрез_участники пенсионеры ДДС",
        "количествовкладчиковфизическихлиц": "Пенсрез_вкладчики_итог",
        "количествоучастников": "Пенсрез_участники",
        "количествоучастниковполучающихнегосударственнуюпенсию": "Пенсрез_участники_пенсионеры",

        # -- Пенсионные накопления (ОПС) --
        "суммасредствпенсионныхнакопленийнаконецмесяцавсеговтомчисле": "Пенснак_",
        "средствапенсионныхнакопленийсформированныевпользузастрахованныхлицкоторымназначенасрочнаяпенсионнаявыплата": "Пенснак_срочная выплата",
        "выплатнойрезервсредствпенсионныхнакоплений": "Пенснак_выплатной резерв",
        "резервпообязательномупенсионномустрахованию": "Пенснак_резерв по ОПС",
        "финансовыйрезультатотинвестированиясредствпенсионныхнакопленийнераспределеннаяприбыльнаконецотчетногопериода": "Пенснак_финрез",
        "количествозастрахованныхлицподействующимдоговорамобобязательномпенсионномстраховании": "Пенснак_ЗЛ",
        "количествозастрахованныхлицосуществляющихформированиесвоихпенсионныхнакопленийвфонде": "Пенснак_ЗЛ"
    }
    # !!! На сайте сентябрь 2021 года представлен датой 31.09.2021 (по состоянию на 14.05.2026)
    sber_dates = ['31.03.2026', '28.02.2026', '31.01.2026', '31.12.2025', '30.11.2025', '31.10.2025', '30.09.2025', '31.08.2025',
                          '31.07.2025', '30.06.2025', '31.05.2025', '30.04.2025', '31.03.2025', '28.02.2025', '31.01.2025', '31.12.2024',
                          '30.11.2024', '31.10.2024', '30.09.2024', '31.08.2024', '31.07.2024', '30.06.2024', '31.05.2024', '30.04.2024',
                          '31.03.2024', '29.02.2024', '31.01.2024', '31.12.2023', '30.11.2023', '31.10.2023', '30.09.2023', '31.08.2023',
                          '31.07.2023', '30.06.2023', '31.05.2023', '30.04.2023', '31.03.2023', '28.02.2023', '31.01.2023', '31.12.2022',
                          '30.11.2022', '31.10.2022', '30.09.2022', '31.08.2022', '31.07.2022', '30.06.2022', '31.05.2022', '30.04.2022',
                          '31.03.2022', '28.02.2022', '31.01.2022', '31.12.2021', '30.11.2021', '31.10.2021', '30.09.2021', '31.08.2021',
                          '31.07.2021', '30.06.2021', '31.05.2021', '30.04.2021', '31.03.2021', '28.02.2021', '31.01.2021', '31.12.2020', 
                          '30.11.2020', '31.10.2020', '31.09.2020', '31.08.2020', '31.07.2020', '30.06.2020', '31.05.2020', '30.04.2020', 
                          '31.03.2020', '29.02.2020', '31.01.2020', '31.12.2019', '30.11.2019', '31.10.2019', '30.09.2019']

    @staticmethod
    def clean_key(dict_key: str) -> str:
        """
        Метод для очистки названия показателя от лишних знаков, пробелов и приведения к нижнему регистру

        dict_ket: str – название показателя, искомое в словаре
        """
        result_key = ''
        for char in dict_key:
            if char.isalpha():
                result_key += char.lower()
        return result_key

    def extract_html_sber(self, date: str=None, max_attempts: int=5) -> list:
        """
        Функция, извлекающая html-верстку таблицы за определенную дату

        date: str – Дата в формате дд.мм.ГГГ, Подается последняя дата заданного месяца.
        !!! На сайте сентябрь 2021 года представлен датой 31.09.2021 (по состоянию на 14.05.2026)
        max_attempts: int – максимальное количество попыток обращения переключения даты 
        """
        def fix_date_for_header(date_str):
            """Превращает '31.09.2021' в '30.09.2021', если день не существует в этом месяце."""
            try:
                datetime.strptime(date_str, '%d.%m.%Y')
                return date_str  # дата и так корректна
            except ValueError:
                # Пробуем уменьшать день, пока не получим валидную дату
                day, month, year = map(int, date_str.split('.'))
                for d in range(day, 0, -1):
                    try:
                        corrected = datetime(year, month, d)
                        return corrected.strftime('%d.%m.%Y')
                    except ValueError:
                        continue
                return date_str  # на всякий случай, если не удалось
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            )
            page = context.new_page()
            page.goto(self.sber_url, wait_until='domcontentloaded', timeout=60000)

            # Ждём появления селектора даты
            page.wait_for_selector('div[class*="DateSelector_dateSelector"] .ant-select', timeout=10000)

            # ---------- ВЫБОР ДАТЫ ----------
            if date is not None:
                for attempt in range(max_attempts):
                    print(f"{attempt + 1} попытка выбора даты (Sber)")
                    # 1. Открываем селектор
                    page.click('div[class*="DateSelector_dateSelector"] .ant-select-selector')
                    page.wait_for_selector('.ant-select-dropdown:not(.ant-select-dropdown-hidden)', timeout=10000)

                    dropdown = page.locator('.ant-select-dropdown:not(.ant-select-dropdown-hidden)')
                    scroll_container = dropdown.locator('.rc-virtual-list-holder')

                    # 2. Прокрутка и выбор опции
                    if scroll_container.count():
                        # Прыжок вниз для загрузки старых дат
                        scroll_container.evaluate('el => el.scrollTop = el.scrollHeight')
                        page.wait_for_timeout(500)

                        target = dropdown.locator('.ant-select-item-option', has_text=date)
                        if target.count() and target.first.is_visible():
                            target.first.click(force=True)
                        else:
                            # Запасной вариант – через React onChange
                            page.keyboard.press('Escape')
                            page.wait_for_timeout(200)
                            page.evaluate('''(date) => {
                                const selectEl = document.querySelector('div[class*="DateSelector_dateSelector"] .ant-select');
                                if (!selectEl) return;
                                const fiberKey = Object.keys(selectEl).find(k => k.startsWith('__reactFiber$'));
                                if (!fiberKey) return;
                                let fiber = selectEl[fiberKey];
                                while (fiber) {
                                    if (fiber.memoizedProps && typeof fiber.memoizedProps.onChange === 'function') {
                                        fiber.memoizedProps.onChange(date, { label: date, value: date });
                                        break;
                                    }
                                    fiber = fiber.return;
                                }
                            }''', date)

                    # 3. Ждём, пока селектор покажет нужную дату
                    try:
                        page.wait_for_selector(f'.ant-select-selection-item[title="{date}"]', timeout=3000)
                    except:
                        pass

                    # 4. Главная проверка: заголовок таблицы с нужной датой (с коррекцией для 31 сентября и т.п.)
                    corrected_date = fix_date_for_header(date)
                    try:
                        page.wait_for_selector(f'th:has-text("Показатели на {corrected_date}")', timeout=5000)
                        break  # успех, выходим из цикла
                    except:
                        if attempt == max_attempts - 1:
                            raise Exception(f"Не удалось выбрать дату {date} после {max_attempts} попыток")
                        page.keyboard.press('Escape')
                        page.wait_for_timeout(1000)

            # ---------- ОЖИДАНИЕ ЗАГРУЗКИ ДАННЫХ ----------
            # (после успешного выбора даты или если дата не указана)
            page.wait_for_selector('text=Загрузка...', state='hidden', timeout=10000)

            # ---------- СБОР ТАБЛИЦЫ НПО / ПДС ----------
            # Дожидаемся появления таблицы (заголовок "Показатели на")
            page.wait_for_selector('th:has-text("Показатели на")', timeout=10000)
            html_npo = page.content()
            soup_npo = BeautifulSoup(html_npo, 'lxml')
            table_npo = soup_npo.find('th', string=lambda t: t and 'Показатели на' in t)
            if table_npo:
                table_npo = table_npo.find_parent('table')

            # ---------- ПЕРЕКЛЮЧЕНИЕ НА ОПС ----------
            ops_button = page.locator('#OperationalFundInfo_headContent').get_by_role('button', name='ОПС')
            ops_button.click()
            # Ждём появления уникального показателя ОПС
            page.wait_for_selector('text=Сумма средств пенсионных накоплений', timeout=10000)
            # Убеждаемся, что индикатор загрузки скрыт
            page.wait_for_selector('text=Загрузка...', state='hidden', timeout=10000)
            # Забираем HTML
            html_ops = page.content()
            soup_ops = BeautifulSoup(html_ops, 'lxml')
            table_ops = soup_ops.find('th', string=lambda t: t and 'Показатели на' in t)
            if table_ops:
                table_ops = table_ops.find_parent('table')

            browser.close()

        return [table_npo, table_ops]
    
    @staticmethod
    def split_cell_by_br(cell: BeautifulSoup) -> list[str]:
        """Возвращает список непустых строк, на которые ячейка разбита <br/> (на любом уровне)."""
        parts = []
        current = []

        # Обход ребенка (+ рекурсия, если в ячейке любой другой тег)
        def traverse(node):
            nonlocal current
            # Если у нас непустой текст, то мы добавляем его в текущий лист
            if isinstance(node, str):
                text = node.strip()
                if text:
                    current.append(text)
            # Если оказался не текст а разделитель br, то, если текущий лист непуст, добавляем объединенную строку в части parts
            elif node.name == 'br':
                if current:
                    parts.append(' '.join(current))
                    current = []
            else:
                # любой другой тег (div, b, span и т.д.) — рекурсивно входим внутрь
                for child in node.children:
                    traverse(child)
            # В итоге функция рекурсивно добавляет элементы в parts

        # Разбираем всех детей ячейки
        for child in cell.children:
            # Применяем к ребенку функцию обхода traverse()
            traverse(child)
        # Так как child не всегда заканчивается br, то в конце обязательно проверяем непустоту current и добавляем сами
        if current:
            parts.append(' '.join(current))

        return parts
    
    def get_df_sber(self, all_tables: list[BeautifulSoup]) -> pd.DataFrame:
        '''
        Метод необходим для того, чтобы превратить html верстку в датафрейм
        
        all_tables: list – список таблиц в формате beautifulsoup
        '''
        # Подфункция для парсинга таблицы и сохранения ее содержимого в список data
        def parse_one_table(table):
            nonlocal data
            rows = table.find_all('tr')[1:] 
            for row in rows:
                cells = row.find_all('td')
                if not cells:
                    continue
                
                # Формируем листы из показателей, значений и единиц измерения (на случай если в строчке несколько показателей)
                indicators = HtmlToExcel.split_cell_by_br(cells[0])
                values     = HtmlToExcel.split_cell_by_br(cells[1])
                units      = HtmlToExcel.split_cell_by_br(cells[2])

                # zip автоматически сопоставит части: из первой строки получим 3 тройки,
                # из остальных — по одной (т.к. там нет <br/>)
                for ind, val, unit in zip(indicators, values, units):
                    cleaned_ind = HtmlToExcel.clean_key(ind)
                    if cleaned_ind in HtmlToExcel.sber_mapping:
                        try:
                            data.append([ind, HtmlToExcel.sber_mapping[cleaned_ind], int(val.replace(' ', '')), unit])
                        except ValueError:
                            data.append([ind, HtmlToExcel.sber_mapping[cleaned_ind], np.nan, unit])
                    else:
                        print(f'!WARNING! Показатель {ind} с сайта не имеет аналога в словаре. Необходима проверка')
                        try:
                            data.append([ind, '', int(val.replace(' ', '')), unit])
                        except ValueError: # Если val - не число (напр. "-"), то вставляем пустое значение
                            data.append([ind, '', np.nan, unit])

        # Применяем к данным
        table1 = all_tables[0] # 1 таблица
        table2 = all_tables[1] # 2 таблица
        data = []
        parse_one_table(table1) # Парсим первую таблицу
        parse_one_table(table2) # Парсим вторую таблицу
        df = pd.DataFrame(data, columns=['Показатель', 'ID', 'Значение', 'Единица измерения'])
        return df

    def extract_data_from_sber(self, date: str=None, to_excel: bool=False, max_attempts: int=5) -> pd.DataFrame:
        '''
        Docstring for extract_data_from_sber

    
        date: str – дата в формате день.месяц.Год(20xx)
        to_excel: bool – выгрузить ли данные в эксель. По умолчанию True
        max_attempts: int – максимальное количество попыток обращения к дате
        '''
        # Извлекаем все таблицы в формате bs4 HTML
        all_tables = self.extract_html_sber(date, max_attempts)
        df = self.get_df_sber(all_tables)
        if to_excel:
            try:
                with pd.ExcelWriter(self.excel_path, 'openpyxl', mode='a', if_sheet_exists='replace') as excel_file:
                    df.to_excel(excel_file, sheet_name='Sber', index=False)
                    print(f'Извлечение данных с {self.sber_url} на лист Sber произошло')
            except FileNotFoundError:
                with pd.ExcelWriter(self.excel_path, 'openpyxl', mode='w') as excel_file:
                    df.to_excel(excel_file, sheet_name='Sber', index=False)
                    print(f'Извлечение данных с {self.sber_url} на лист Sber произошло')
        return df
    
    # ===== Aviapolis =====

    aviapolis_dates = ['31.03.2026', '28.02.2026', '31.01.2026', '31.12.2025', '30.11.2025', '31.10.2025', '30.09.2025', '31.08.2025',
                        '31.07.2025', '30.06.2025', '31.05.2025', '30.04.2025', '31.03.2025', '28.02.2025', '31.01.2025', '31.12.2024',
                        '30.11.2024', '31.10.2024', '30.09.2024', '31.08.2024', '31.07.2024', '30.06.2024', '31.05.2024', '30.04.2024',
                        '31.03.2024', '29.02.2024', '31.01.2024', '31.12.2023', '30.11.2023', '31.10.2023', '30.09.2023', '31.08.2023',
                        '31.07.2023', '30.06.2023', '31.05.2023', '30.04.2023', '31.03.2023', '28.02.2023', '31.01.2023', '31.12.2022',
                        '30.11.2022', '31.10.2022', '30.09.2022', '31.08.2022', '31.07.2022', '30.06.2022', '31.05.2022', '30.04.2022',
                        '31.03.2022', '28.02.2022', '31.01.2022', '31.12.2021', '30.11.2021', '31.10.2021', '30.09.2021', '31.08.2021',
                        '31.07.2021', '30.06.2021', '31.05.2021', '30.04.2021', '31.03.2021', '28.02.2021', '31.01.2021', '31.12.2020', 
                        '30.11.2020', '31.10.2020', '30.09.2020', '31.08.2020', '31.07.2020', '30.06.2020', '31.05.2020', '30.04.2020', 
                        '31.03.2020', '29.02.2020', '31.01.2020', '31.12.2019', '30.11.2019', '31.10.2019', '30.09.2019']

    @staticmethod
    def parse_aviapolis_table(html):
        soup = BeautifulSoup(html, 'lxml')
        # Ищем td с классом style45
        content_td = soup.find('td', class_='style45')
        if not content_td:
            raise ValueError("Не найдена ячейка с таблицей показателей")
        table = content_td.find('table')
        rows = table.find_all('tr')

        # Ищем строку с датами – она содержит текст типа "31.01.2026"
        date_row = None
        for row in rows:
            cells = row.find_all('td')
            if len(cells) >= 3:
                # Проверяем первую ячейку (не заголовок "№ пп")
                if cells[0].get_text(strip=True) == '' or cells[0].get_text(strip=True) == '№ пп':
                    continue
                # Проверяем, что в третьей ячейке есть дата вида дд.мм.гггг
                if re.match(r'\d{2}\.\d{2}\.\d{4}', cells[2].get_text(strip=True)):
                    date_row = row
                    break
        if not date_row:
            # Если не нашли – пробуем вторую строку как обычно
            date_row = rows[1] if len(rows) > 1 else None
            if not date_row:
                raise ValueError("Не найдена строка с датами")
        
        # Извлекаем даты, начиная с индекса 2 (первые два – служебные)
        date_cells = date_row.find_all('td')
        dates = [cell.get_text(strip=True) for cell in date_cells if re.match(r'\d{2}\.\d{2}\.\d{4}', cell.get_text(strip=True))]
        data = []
        for row in rows:
            cols = row.find_all('td')
            if len(cols) < 3:
                continue
            # Проверяем, что первый столбец – число (№ пп)
            first_text = cols[0].get_text(strip=True)
            if not first_text.isdigit():
                continue  # пропускаем заголовки и строки без номера
            
            indicator = cols[1].get_text(strip=True)
            values = [col.get_text(strip=True) for col in cols[2:]]
            # Обрезаем values до длины dates
            values = values[:len(dates)]
            for d, v in zip(dates, values):
                data.append([indicator, d, v])
        
        df = pd.DataFrame(data, columns=['Показатель', 'Дата', 'Значение'])
        df['Значение'] = df['Значение'].str.replace(' ', '', regex=False).replace('-', None)
        return df


    def extract_data_from_aviapolis(self, date_str: str=None, to_excel: bool=False) -> pd.DataFrame:
        def collect_year_links(html, base_url):
            soup = BeautifulSoup(html, 'lxml')
            links = {}
            for a in soup.find_all('a', href=True):
                text = a.get_text(strip=True)
                if 'Показатели деятельности за' in text:
                    year = text.split()[-2].rstrip('г.')  # "2025"
                    href = a['href']
                    full_url = requests.compat.urljoin(base_url, href)
                    links[int(year)] = full_url
            return links
        
        if date_str:
            target_date = datetime.strptime(date_str, '%d.%m.%Y')
            target_year = target_date.year
        else:
            target_year = None

        # Сначала получаем текущую страницу (базовый URL) и собираем ссылки
        current_html = requests.get(self.aviapolis_url).content
        year_links = collect_year_links(current_html, self.aviapolis_url)

        if target_year and target_year != datetime.now().year:
            # Нужна страница за другой год
            if target_year not in year_links:
                raise ValueError(f'Данные за {target_year} год не найдены')
            url = year_links[target_year]
        else:
            url = self.aviapolis_url

        html = requests.get(url).content
        df_all = HtmlToExcel.parse_aviapolis_table(html)

        # Фильтр по дате, если указана
        if date_str:
            df = df_all[df_all['Дата'] == date_str].copy()
        else:
            df = df_all.copy()
        
        # Выносим единицы измерения в отдельное поле
        df['Показатель'] = df['Показатель'].str.replace('(в соответствии с п. 1.2.2. Указания Банка России от 18.06.2019 № 5175-У)', '')
        df['Единица измерения'] = df['Показатель'].map(lambda indicator: str.split(indicator, ', ')[-1])
        df['Показатель'] = df['Показатель'].map(lambda indicator: ' '.join(str.split(indicator, ', ')[:-1]))

        if to_excel:
            try:
                with pd.ExcelWriter(self.excel_path, 'openpyxl', mode='a', if_sheet_exists='replace') as excel_file:
                    df.to_excel(excel_file, sheet_name='Aviapolis', index=False)
                    print(f'Извлечение данных с {self.aviapolis_url} на лист Aviapolis произошло')
            except FileNotFoundError:
                with pd.ExcelWriter(self.excel_path, 'openpyxl', mode='w') as excel_file:
                    df.to_excel(excel_file, sheet_name='Aviapolis', index=False)
                    print(f'Извлечение данных с {self.aviapolis_url} на лист Aviapolis произошло')
        

        return df[['Показатель', 'Значение', 'Дата', 'Единица измерения']]
        


