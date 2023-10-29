import os

import numpy as np
import pandas as pd
import xlsxwriter
from pathlib import Path
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Директория, где лежат рабочие файлы
os.chdir(r"C:\Users\AVShestakov\split")

file = 'inventory_Все филиалы2.xlsx'

# parsing Inventory Ericsson
df_inventoryEric = pd.read_excel(file, sheet_name='eri')

# парсим rru_type
df_inventoryEric_rru900 = df_inventoryEric['rru_900'].str.split(',', expand=True)
df_inventoryEric_rru1800 = df_inventoryEric['rru_1800'].str.split(',', expand=True)
df_inventoryEric_rru2100 = df_inventoryEric['rru_2100'].str.split(',', expand=True)
df_inventoryEric_rru900 = df_inventoryEric_rru900.rename(columns={0: 'rru_900_1', 1: 'rru_900_2',
                                                                  2: 'rru_900_3', 3: 'rru_900_4',
                                                                  4: 'rru_900_5', 5: 'rru_900_6',
                                                                  })

df_inventoryEric_rru1800 = df_inventoryEric_rru1800.rename(columns={0: 'RRU_1800_1', 1: 'RRU_1800_2',
                                                                    2: 'RRU_1800_3', 3: 'RRU_1800_4',
                                                                    4: 'RRU_1800_5', 5: 'RRU_1800_6',
                                                                    6: 'RRU_1800_7', 7: 'RRU_1800_8',
                                                                    8: 'RRU_1800_9', 9: 'RRU_1800_10',
                                                                    })

df_inventoryEric_rru2100 = df_inventoryEric_rru2100.rename(columns={0: 'RRU_2100_1', 1: 'RRU_2100_2',
                                                                    2: 'RRU_2100_3', 3: 'RRU_2100_4',
                                                                    4: 'RRU_2100_5', 5: 'RRU_2100_6',
                                                                    6: 'RRU_2100_7', 7: 'RRU_2100_8',
                                                                    8: 'RRU_2100_9', 9: 'RRU_2100_10',
                                                                    })

final_df_eri = pd.concat([df_inventoryEric, df_inventoryEric_rru900, df_inventoryEric_rru1800,
                          df_inventoryEric_rru2100], axis=1)

# parsing Inventory Nokia
df_inventoryNok = pd.read_excel(file, sheet_name='nok')

# разделим колонку LCELL и GCELL с разделителем запятая
df_inventoryNok_Lcell = df_inventoryNok['LCELL'].str.split(',', expand=True)
df_inventoryNok_Lcell = df_inventoryNok_Lcell.rename(columns={0: 'LCELL1', 1: 'LCELL2', 2: 'LCELL3'})

df_inventoryNok_Gcell = df_inventoryNok['GCELL'].str.split(',', expand=True)
df_inventoryNok_Gcell = df_inventoryNok_Gcell.rename(columns={0: 'GCELL1', 1: 'GCELL2'})

final_df_nok = pd.concat([df_inventoryNok, df_inventoryNok_Lcell, df_inventoryNok_Gcell], axis=1)

# parsing Inventory Huawei
df_inventoryHua = pd.read_excel(file, sheet_name='hua')

# парсим board_type
df_inventoryHua_boards = df_inventoryHua['board_type'].str.split(',', expand=True)
df_inventoryHua_boards = df_inventoryHua_boards.rename(columns={0: 'board_type1', 1: 'board_type2',
                                                                2: 'board_type3', 3: 'board_type4',
                                                                4: 'board_type5', 5: 'board_type6',
                                                                6: 'board_type7', 7: 'board_type8',
                                                                8: 'board_type9',
                                                                })

# парсим кол-во RRU (в данном случае 900-й band)
df_inventoryHua_rru900 = df_inventoryHua['RRU_900'].str.split(',', expand=True)
df_inventoryHua_rru1800 = df_inventoryHua['RRU_1800'].str.split(',', expand=True)
df_inventoryHua_rru2100 = df_inventoryHua['RRU_2100'].str.split(',', expand=True)
df_inventoryHua_rru900 = df_inventoryHua_rru900.rename(columns={0: 'RRU_900_1', 1: 'RRU_900_2',
                                                                2: 'RRU_900_3', 3: 'RRU_900_4',
                                                                4: 'RRU_900_5', 5: 'RRU_900_6',
                                                                6: 'RRU_900_7', 7: 'RRU_900_8',
                                                                8: 'RRU_900_9', 9: 'RRU_900_10',
                                                                10: 'RRU_900_11',
                                                                })

df_inventoryHua_rru1800 = df_inventoryHua_rru1800.rename(columns={0: 'RRU_1800_1', 1: 'RRU_1800_2',
                                                                  2: 'RRU_1800_3', 3: 'RRU_1800_4',
                                                                  4: 'RRU_1800_5', 5: 'RRU_1800_6',
                                                                  6: 'RRU_1800_7', 7: 'RRU_1800_8',
                                                                  8: 'RRU_1800_9', 9: 'RRU_1800_10',
                                                                  10: 'RRU_1800_11', 11: 'RRU_1800_12',
                                                                  12: 'RRU_1800_13', 13: 'RRU_1800_14',
                                                                  14: 'RRU_1800_15', 15: 'RRU_1800_16',
                                                                  })

df_inventoryHua_rru2100 = df_inventoryHua_rru2100.rename(columns={0: 'RRU_2100_1', 1: 'RRU_2100_2',
                                                                  2: 'RRU_2100_3', 3: 'RRU_2100_4',
                                                                  4: 'RRU_2100_5', 5: 'RRU_2100_6',
                                                                  6: 'RRU_2100_7', 7: 'RRU_2100_8',
                                                                  8: 'RRU_2100_9', 9: 'RRU_2100_10',
                                                                  10: 'RRU_2100_11', 11: 'RRU_2100_12',
                                                                  12: 'RRU_2100_13', 13: 'RRU_2100_14',
                                                                  14: 'RRU_2100_15', 15: 'RRU_2100_16',
                                                                  })

df_inventoryHua_rru900 = df_inventoryHua_rru900.replace('', None)
df_inventoryHua_rru900 = df_inventoryHua_rru900.fillna(0)
df_inventoryHua_rru900 = df_inventoryHua_rru900.astype(int)
df_inventoryHua_rru1800 = df_inventoryHua_rru1800.replace('', None)
df_inventoryHua_rru1800 = df_inventoryHua_rru1800.fillna(0)
df_inventoryHua_rru1800 = df_inventoryHua_rru1800.astype(int)
df_inventoryHua_rru2100 = df_inventoryHua_rru2100.replace('', None)
df_inventoryHua_rru2100 = df_inventoryHua_rru2100.fillna(0)
df_inventoryHua_rru2100 = df_inventoryHua_rru2100.astype(int)

final_df_hua = pd.concat([df_inventoryHua, df_inventoryHua_boards, df_inventoryHua_rru900,
                          df_inventoryHua_rru1800, df_inventoryHua_rru2100], axis=1)

final_df_hua['rru900_sum'] = final_df_hua[['RRU_900_1', 'RRU_900_2', 'RRU_900_3', 'RRU_900_4',
                                           'RRU_900_5', 'RRU_900_6', 'RRU_900_7', 'RRU_900_8',
                                           'RRU_900_9', 'RRU_900_10', 'RRU_900_11']].sum(axis=1)
final_df_hua['rru1800_sum'] = final_df_hua[['RRU_1800_1', 'RRU_1800_2', 'RRU_1800_3',
                                            'RRU_1800_4', 'RRU_1800_5', 'RRU_1800_6', 'RRU_1800_7', 'RRU_1800_8',
                                            'RRU_1800_9', 'RRU_1800_10',
                                            'RRU_1800_11', 'RRU_1800_12', 'RRU_1800_13', 'RRU_1800_14', 'RRU_1800_15',
                                            'RRU_1800_16']].sum(axis=1)

final_df_hua.drop(['RRU_900_1', 'RRU_900_2', 'RRU_900_3', 'RRU_900_4', 'RRU_900_5', 'RRU_900_6', 'RRU_900_7',
                   'RRU_900_8', 'RRU_900_9', 'RRU_900_10', 'RRU_900_11', 'RRU_1800_1', 'RRU_1800_2', 'RRU_1800_3',
                   'RRU_1800_4', 'RRU_1800_5', 'RRU_1800_6', 'RRU_1800_7', 'RRU_1800_8', 'RRU_1800_9', 'RRU_1800_10',
                   'RRU_1800_11', 'RRU_1800_12', 'RRU_1800_13', 'RRU_1800_14', 'RRU_1800_15', 'RRU_1800_16'],
                  axis=1, inplace=True)

writer_df = pd.ExcelWriter('invent_split.xlsx', engine='xlsxwriter')
# df_inventoryHua_boards.to_excel(writer_df, 'hua_boards', index=False)
# df_inventoryHua_rru900.to_excel(writer_df, 'hua_rru900', index=False)
df_inventoryHua_rru1800.to_excel(writer_df, 'hua_rru1800', index=False)
# df_inventoryEric_rru900.to_excel(writer_df, 'eri_rru900', index=False)
# df_inventoryEric_rru1800.to_excel(writer_df, 'eri_rru1800', index=False)
final_df_nok.to_excel(writer_df, 'nok')
final_df_hua.to_excel(writer_df, 'hua')
final_df_eri.to_excel(writer_df, 'eri')
writer_df.close()

# Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
