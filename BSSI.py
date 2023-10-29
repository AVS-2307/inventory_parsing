import os

import numpy as np
import pandas as pd
import xlsxwriter
from pathlib import Path
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Директория, где лежат рабочие файлы
os.chdir(r"C:\Users\AVShestakov\split")

# !!! ПРОВЕРИТЬ ФАЙЛ НА РАЗДЕЛИТЕЛИ, особенно % покрытия. ДОЛЖНО БЫТЬ ; !!!
file = 'BSSI_cov.xlsx'

df_BSSI = pd.read_excel(file, sheet_name='Лист1')

# df по новостройкам
df_newBuilding = df_BSSI[['Стандарт', 'ID объекта Новостройки', '% Прироста покрытия в полигоне новостройки']]

# исключим пустые строки
df_newBuilding = df_newBuilding.loc[df_newBuilding['ID объекта Новостройки'].notna()]

# writer_df_newBuilding = pd.ExcelWriter('new_building.xlsx', engine='xlsxwriter')
# df_newBuilding.to_excel(writer_df_newBuilding, 'newBuilding', index=False)
# writer_df_newBuilding.close()

# df по населенным пунктам
df_cities = df_BSSI[['Стандарт', 'ID объекта НП', '% Прироста покрытия в нп']]

# исключим пустые строки
df_cities = df_cities.loc[df_cities['ID объекта НП'] != '-']

# writer_df_cities = pd.ExcelWriter('cities.xlsx', engine='xlsxwriter')
# df_cities.to_excel(writer_df_cities, 'cities', index=False)
# writer_df_cities.close()

# df по проблемным зонам
# copy() для обхода ошибки Try using .loc[row_indexer,col_indexer] = value instead при replace
df_problemZones = df_BSSI[['Стандарт', 'ID Проблемной зоны', '% Прироста покрытия в пз']].copy()

# исключим пустые строки
df_problemZones = df_problemZones[(df_problemZones['ID Проблемной зоны'] != '-') &
                                  (df_problemZones['ID Проблемной зоны'].notna())]

# заменяем разделение пз с , на ;
df_problemZones['ID Проблемной зоны'] = df_problemZones['ID Проблемной зоны'].str.replace(',', ';')

# заменим прочерки нулями
df_problemZones['% Прироста покрытия в пз'] = df_problemZones['% Прироста покрытия в пз'].replace('-', 0)

writer_df_problemZones = pd.ExcelWriter('probzones.xlsx', engine='xlsxwriter')
df_problemZones.to_excel(writer_df_problemZones, 'Sheet1', index=False)
writer_df_problemZones.close()
# разделим колонку на несколько с разделителем точка с запятой
df_problemZones[['ID пз1', 'ID пз2', 'ID пз3', 'ID пз4', 'ID пз5', 'ID пз6', 'ID пз7', 'ID пз8', 'ID пз9']] \
    = df_problemZones['ID Проблемной зоны'].str.split(';', expand=True)

# заменим пустые ячейки ID пз1 значениями ячейки ID Проблемной зоны
df_problemZones.loc[df_problemZones['ID пз1'].isnull(), 'ID пз1'] = df_problemZones['ID Проблемной зоны']

df_problemZones[['% пз1', '% пз2', '% пз3', '% пз4', '% пз5', '% пз6', '% пз7', '% пз8', '% пз9']] = \
    df_problemZones['% Прироста покрытия в пз'].str.split(';', expand=True)

# заменим пустые ячейки % пз1 значениями ячейки % Прироста покрытия в пз
df_problemZones.loc[df_problemZones['% пз1'].isnull(), '% пз1'] = df_problemZones['% Прироста покрытия в пз']

# объединим все пз в один файл
df_pz1 = df_problemZones[['Стандарт', 'ID пз1', '% пз1']]
df_pz1.columns = ['Стандарт', 'ID пз', '% пз']
df_pz2 = df_problemZones[['Стандарт', 'ID пз2', '% пз2']]
df_pz2.columns = ['Стандарт', 'ID пз', '% пз']
df_pzTot = pd.concat([df_pz1, df_pz2], axis=0, ignore_index=True)

writer_df_pzTot = pd.ExcelWriter('pz_tot.xlsx', engine='xlsxwriter')
df_pzTot.to_excel(writer_df_pzTot, 'Sheet1', index=False)
writer_df_pzTot.close()
