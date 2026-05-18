from Parser import HtmlToExcel

extractor = HtmlToExcel(
    excel_path='NPF_parsing_html2excel/result.xlsx',
    # sber_url ='https://npfsberbanka.ru/about/information-to-be-disclosed/results/',
    # aviapolis_url ='https://www.npf-aviapolis.ru/index.files/Page025.html'
)

print('Пример дат для сбера')
print(extractor.sber_dates[:10])
print()

print('Пример дат для Авиаполиса')
print(extractor.aviapolis_dates[:10])

date_to_parse_sber = extractor.sber_dates[4]
date_to_parse_aviapolis = extractor.aviapolis_dates[4]

df_sber = extractor.extract_data_from_sber(
    date=date_to_parse_sber,
    to_excel=True
)
df_aviapolis = extractor.extract_data_from_aviapolis(
    date_str=date_to_parse_aviapolis,
    to_excel=True
)

print(f'Показатели сбера на {date_to_parse_sber}')
print(df_sber)
print()
print(f'Показатели Авиаполиса на {date_to_parse_aviapolis}')
print(df_aviapolis)
