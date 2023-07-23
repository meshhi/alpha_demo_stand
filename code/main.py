import pandas as pd
import re
import numpy as np
import random

df = pd.read_excel('../files/Раздел 1- Население.xlsx', skiprows=5, skipfooter=2, sheet_name='1.1.')
df_2 = pd.read_excel('../files/Раздел 2 - Труд.xlsx', skiprows=6, skipfooter=2, sheet_name='2.1.')
df_3 = pd.read_excel('../files/Раздел 3 - Уровень жизни населения.xlsx', skiprows=6, skipfooter=5, sheet_name='3.2.')
df_4 = pd.read_excel('../files/Раздел 4 - Образование.xlsx', skiprows=5, skipfooter=2, sheet_name='4.12.')
df_5 = pd.read_excel('../files/Раздел 5 - Здравоохранение.xlsx', skiprows=6, skipfooter=2, sheet_name='5.4.1.')
df_6 = pd.read_excel('../files/Раздел 1- Население.xlsx', skiprows=5, skipfooter=2, sheet_name='1.5.')
df_7 = pd.read_excel('../files/Раздел 1- Население.xlsx', skiprows=6, skipfooter=3, sheet_name='1.6.2.')
df_8 = pd.read_excel('../files/Раздел 1- Население.xlsx', skiprows=5, skipfooter=2, sheet_name='1.9.')
df_9 = pd.read_excel('../files/Раздел 1- Население.xlsx', skiprows=5, skipfooter=2, sheet_name='1.10.')
df_10 = pd.read_excel('../files/Раздел 1- Население.xlsx', skiprows=5, skipfooter=2, sheet_name='1.3.')
df_11 = pd.read_excel('../files/Раздел 1- Население.xlsx', skiprows=5, skipfooter=2, sheet_name='1.4.')
df_12 = pd.read_excel('../files/Раздел 2 - Труд.xlsx', skiprows=6, skipfooter=2, sheet_name='2.9.')
df_13 = pd.read_excel('../files/Раздел 2 - Труд.xlsx', skiprows=7, skipfooter=2, sheet_name='2.10.1.')
regions = {}
not_allow_regions = {
  'в том числе:' : True,
}

def transform_region_name(region_name):
  region_name = region_name.strip()
  region_name = region_name.replace('\n', '')
  region_name = region_name.replace('–', '-')
  while region_name.find('  ') != -1:
    region_name = region_name.replace('  ', ' ')
  if region_name == 'Республика Северная Осетия -Алания' or region_name == 'Республика Северная Осетия-Алания':
    region_name = 'Республика Северная Осетия - Алания'
  if region_name == 'Ханты-Мансийский автономный округ-Югра':
    region_name = 'Ханты-Мансийский автономный округ - Югра'
  if region_name == 'Южныйфедеральный округ':
    region_name = 'Южный федеральный округ'
    
    
  return region_name

def ex_structure(row):
    region_name = row['Unnamed: 0']
    region_name = transform_region_name(region_name)

    if region_name not in regions and region_name not in not_allow_regions:
      # print(region_name + '\n')
      regions[region_name] = {} 
      for key in row.index:
        if key not in regions[region_name] and key != 'Unnamed: 0':
          regions[region_name][key] = {}
        else:
          pass
    else:
      pass
    return row

def ex_population(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Численность населения'] = row[key]
  return row

def ex_work(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Численность рабочей силы'] = row[key]
  return row

def ex_incomes(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Среднедушевые денежные доходы населения'] = row[key]
  return row

def ex_high_education(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Выпуск квалифицированных рабочих и кадров'] = row[key]
  return row

def ex_doctors(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Численность врачей всех специальностей'] = row[key]
  return row

def ex_men_women_count(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Соотношение мужчин и женщин'] = row[key]
  return row

def ex_can_work_age(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Население в трудоспособном возрасте'] = row[key]
  return row

def ex_birth(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Общие коэффициенты рождаемости'] = row[key]
  return row

def ex_death(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Общие коэффициенты смертности'] = row[key]
  return row

def ex_city_population(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Удельный вес городского населения'] = row[key]
  return row

def ex_country_population(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Удельный вес сельского населения'] = row[key]
  return row

def ex_unemploye(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Численность незанятых граждан'] = row[key]
  return row

def ex_unemploye_lvl(row):
  region_name = row['Unnamed: 0']
  region_name = transform_region_name(region_name)
  if region_name in not_allow_regions:
    return row
  for key in row.index:
    if key != 'Unnamed: 0':
      regions[region_name][key]['Уровень безработицы'] = row[key]
  return row

# update result structure
df = df.apply(ex_structure, axis=1)
df_2 = df_2.apply(ex_structure, axis=1)
df_3 = df_3.apply(ex_structure, axis=1)
df_4 = df_4.apply(ex_structure, axis=1)
df_5 = df_5.apply(ex_structure, axis=1)
df_6 = df_6.apply(ex_structure, axis=1)
df_7 = df_7.apply(ex_structure, axis=1)
df_8 = df_8.apply(ex_structure, axis=1)
df_9 = df_9.apply(ex_structure, axis=1)
df_10 = df_10.apply(ex_structure, axis=1)
df_11 = df_11.apply(ex_structure, axis=1)
df_12 = df_12.apply(ex_structure, axis=1)
df_13 = df_13.apply(ex_structure, axis=1)

#  extract metrics
df = df.apply(ex_population, axis=1)
df_2 = df_2.apply(ex_work, axis=1)
df_3 = df_3.apply(ex_incomes, axis=1)
df_4 = df_4.apply(ex_high_education, axis=1)
df_5 = df_5.apply(ex_doctors, axis=1)
df_6 = df_6.apply(ex_men_women_count, axis=1)
df_7 = df_7.apply(ex_can_work_age, axis=1)
df_8 = df_8.apply(ex_birth, axis=1)
df_9 = df_9.apply(ex_death, axis=1)
df_10 = df_10.apply(ex_city_population, axis=1)
df_11 = df_11.apply(ex_country_population, axis=1)
df_12 = df_12.apply(ex_unemploye, axis=1)
df_13 = df_13.apply(ex_unemploye_lvl, axis=1)

# print(regions)
# print(df_3)
regions_hierarchy = []
count_reg_id = 0
for region in regions:
  count_reg_id = count_reg_id + 1
  new_row = {'ID':count_reg_id, 'PARENT_ID': 1, 'Регион':region, }
  new_row_list = list(new_row.values())
  regions_hierarchy.append(new_row_list)
regions_df = {}
counter_regions = 0
for item in regions_hierarchy:
  counter_regions = counter_regions + 1
  regions_df[counter_regions] = item

df_regions = pd.DataFrame.from_dict(regions_df, orient='index', columns=['ID', 'PARENT_ID', 'Регион'])

def region_transform(row):
  # округ - федерация
  if row["Регион"] == "Центральный федеральный округ" or row["Регион"] == "Северо-Западный федеральный округ" or row["Регион"] == "Южный федеральный округ" or row["Регион"] == "Северо-Кавказский федеральный округ" or row["Регион"] == "Приволжский федеральный округ" or row["Регион"] == "Уральский федеральный округ" or row["Регион"] == "Сибирский федеральный округ" or row["Регион"] == "Дальневосточный федеральный округ":
    rslt_df = df_regions[df_regions['Регион'] == "Российская Федерация"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']
  # регион - округ
  # Центральный федеральный округ
  if row["Регион"] == "Белгородская область" or row["Регион"] == "Брянская область" or row["Регион"] == "Владимирская область" or row["Регион"] == "Воронежская область" or row["Регион"] == "Ивановская область"or row["Регион"] == "Калужская область" or row["Регион"] == "Костромская область"or row["Регион"] == "Курская область" or row["Регион"] == "Липецкая область"or row["Регион"] == "Московская область" or row["Регион"] == "Орловская область"or row["Регион"] == "Рязанская область" or row["Регион"] == "Смоленская область"or row["Регион"] == "Тамбовская область" or row["Регион"] == "Тверская область"or row["Регион"] == "Тульская область" or row["Регион"] == "Ярославская область"or row["Регион"] == "г. Москва":
    rslt_df = df_regions[df_regions['Регион'] == "Центральный федеральный округ"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']

  # Северо-Западный федеральный округ
  if row["Регион"] == "Республика Карелия" or row["Регион"] == "Республика Коми" or row["Регион"] == "Архангельская область" or row["Регион"] == "Ненецкий автономный округ" or row["Регион"] == "Вологодская область"or row["Регион"] == "Калининградская область" or row["Регион"] == "Ленинградская область"or row["Регион"] == "Мурманская область" or row["Регион"] == "Новгородская область"or row["Регион"] == "Псковская область" or row["Регион"] == "г. Санкт-Петербург":
    rslt_df = df_regions[df_regions['Регион'] == "Северо-Западный федеральный округ"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']

  # Южный федеральный округ
  if row["Регион"] == "Республика Адыгея" or row["Регион"] == "Республика Калмыкия" or row["Регион"] == "Республика Крым" or row["Регион"] == "Краснодарский край" or row["Регион"] == "Астраханская область"or row["Регион"] == "Волгоградская область" or row["Регион"] == "Ростовская область"or row["Регион"] == "г. Севастополь":
    rslt_df = df_regions[df_regions['Регион'] == "Южный федеральный округ"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']

  # Северо-Кавказский федеральный округ
  if row["Регион"] == "Ставропольский край" or row["Регион"] == "Чеченская Республика" or row["Регион"] == "Республика Северная Осетия - Алания" or row["Регион"] == "Карачаево-Черкесская Республика" or row["Регион"] == "Кабардино-Балкарская Республика"or row["Регион"] == "Республика Ингушетия" or row["Регион"] == "Республика Дагестан":
    rslt_df = df_regions[df_regions['Регион'] == "Северо-Кавказский федеральный округ"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']

  # Приволжский федеральный округ
  if row["Регион"] == "Республика Башкортостан" or row["Регион"] == "Ульяновская область" or row["Регион"] == "Саратовская область" or row["Регион"] == "Самарская область" or row["Регион"] == "Пензенская область"or row["Регион"] == "Оренбургская область" or row["Регион"] == "Нижегородская область" or row["Регион"] == "Кировская область" or row["Регион"] == "Пермский край" or row["Регион"] == "Чувашская Республика" or row["Регион"] == "Удмуртская Республика" or row["Регион"] == "Республика Татарстан" or row["Регион"] == "Республика Мордовия" or row["Регион"] == "Республика Марий Эл":
    rslt_df = df_regions[df_regions['Регион'] == "Приволжский федеральный округ"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']

  # Уральский федеральный округ
  if row["Регион"] == "Курганская область" or row["Регион"] == "Тюменская область без автономных округов" or row["Регион"] == "Челябинская область" or row["Регион"] == "Ямало-Ненецкий автономный округ" or row["Регион"] == "Ханты-Мансийский автономный округ - Югра"or row["Регион"] == "Тюменская область" or row["Регион"] == "Свердловская область":
    rslt_df = df_regions[df_regions['Регион'] == "Уральский федеральный округ"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']

  # Сибирский федеральный округ
  if row["Регион"] == "Республика Алтай" or row["Регион"] == "Республика Тыва" or row["Регион"] == "Республика Хакасия" or row["Регион"] == "Алтайский край" or row["Регион"] == "Красноярский край"or row["Регион"] == "Иркутская область" or row["Регион"] == "Кемеровская область" or row["Регион"] == "Новосибирская область" or row["Регион"] == "Омская область" or row["Регион"] == "Томская область":
    rslt_df = df_regions[df_regions['Регион'] == "Сибирский федеральный округ"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']

  # Дальневосточный федеральный округ
  if row["Регион"] == "Республика Бурятия" or row["Регион"] == "Чукотский автономный округ" or row["Регион"] == "Еврейская автономная область" or row["Регион"] == "Сахалинская область" or row["Регион"] == "Магаданская область"or row["Регион"] == "Амурская область" or row["Регион"] == "Хабаровский край" or row["Регион"] == "Приморский край" or row["Регион"] == "Камчатский край" or row["Регион"] == "Забайкальский край" or row["Регион"] == "Республика Саха (Якутия)":
    rslt_df = df_regions[df_regions['Регион'] == "Дальневосточный федеральный округ"]
    row["PARENT_ID"] = rslt_df.iloc[0]['ID']

  return row
df_regions = df_regions.apply(lambda row: region_transform(row), axis=1)
df_regions.to_excel("./regions.xlsx") 

data_list = []
for region in regions:
  for year in regions[region]:
    new_row = {'Регион':region, 'Год':year, 'Численность населения':regions[region][year]['Численность населения'], 'Численность рабочей силы':regions[region][year]['Численность рабочей силы'], 'Среднедушевые денежные доходы населения':regions[region][year]['Среднедушевые денежные доходы населения'], 'Выпуск квалифицированных рабочих и кадров':regions[region][year]['Выпуск квалифицированных рабочих и кадров'], 'Численность врачей всех специальностей':regions[region][year]['Численность врачей всех специальностей'], 'Соотношение мужчин и женщин':regions[region][year]['Соотношение мужчин и женщин'], 'Население в трудоспособном возрасте':regions[region][year]['Население в трудоспособном возрасте'], 'Общие коэффициенты рождаемости':regions[region][year]['Общие коэффициенты рождаемости'], 'Общие коэффициенты смертности':regions[region][year]['Общие коэффициенты смертности'], 'Удельный вес городского населения':regions[region][year]['Удельный вес городского населения'], 'Удельный вес сельского населения':regions[region][year]['Удельный вес сельского населения'], 'Численность незанятых граждан':regions[region][year]['Численность незанятых граждан'], 'Уровень безработицы':regions[region][year]['Уровень безработицы'],}
    new_row_list = list(new_row.values())
    data_list.append(new_row_list)

counter = 0
data_df = {}
for item in data_list:
  counter = counter + 1
  data_df[counter] = item

df_result = pd.DataFrame.from_dict(data_df, orient='index', columns=['Регион', 'Год', 'Численность населения', 'Численность рабочей силы', 'Среднедушевые денежные доходы населения', 'Выпуск квалифицированных рабочих и кадров', 'Численность врачей всех специальностей', 'Соотношение мужчин и женщин', 'Население в трудоспособном возрасте', 'Общие коэффициенты рождаемости', 'Общие коэффициенты смертности', 'Удельный вес городского населения', 'Удельный вес сельского населения', 'Численность незанятых граждан', 'Уровень безработицы'])
df_result.to_excel("./super_demo_transformed_4.xlsx")  