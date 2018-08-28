import os
import openpyxl
import pandas as pd

conf_file = pd.ExcelFile("Copy of EMFP CSMB DT Aug Block Config A04.xlsx")
sheets = conf_file.sheet_names

writer = pd.ExcelWriter('pandas_multiple.xlsx', engine='xlsxwriter')

for sheet in sheets:
    if sheet == "Tracker" or sheet == "Revision":
        pass
    else:
        df = conf_file.parse(sheet)
        df.columns.str.strip()
        if sheet == "Redskull(5675)":
            print(df)
            print(df.shape)
            print(df.iat[8,2])
            print(df.iat[8,1])
            base_idx = []
            for col in range(0,df.shape[1]):
                try:
                    row = list(df.iloc[:,col].str.contains("BASE")).index(True)
                except ValueError:
                    row = -1

                if row >= 0:
                    base_idx.append([row,col])

            base_mod_idx = [[base_idx[0][0], base_idx[0][1]-1]]
            print("base desc index: {}".format(base_idx))
            
            print("base mod index: {}".format(base_mod_idx))
            print("base mod: {}".format(df.iat[base_mod_idx[0][0],base_mod_idx[0][1]]))
            print("base desc: {}".format(df.iat[base_idx[0][0],base_idx[0][1]]))

        df.to_excel(writer, sheet_name=sheet)

writer.save
