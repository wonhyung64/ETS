#%%
import pandas as pd
import openpyxl as xl
import re
import os
from functools import reduce

# %%
work_dir = "D:/won/data/ETS"

os.listdir(work_dir)

df = pd.read_excel(f"{work_dir}/ghg_emissions_v7.xlsx")
concat_df = pd.read_excel(f"{work_dir}/corp_value.xlsx", header=2)
concat_df = concat_df.iloc[:,2:]
concat_df_col = ["company"] + [f"{col.split('/Annual')[1].split('/')[0][1:]}_{col.split('/Annual')[0]}" for col in concat_df.columns[1:]]
concat_df.columns = concat_df_col

concat_df['company'] = [re.sub(r'주식회사|\(주\)|\(유\)|유한회사|㈜|\s', '', i) for i in concat_df['company']]

concat_df_tmp = concat_df.drop_duplicates(['company'], keep=False)
len(concat_df_tmp)

# %% Kis MERGE
temp = pd.merge(df, concat_df_tmp, how='left', on='company')

#%% EXPORT
temp.to_excel("D:/won/data/ETS/ghg_emissions_v8.xlsx", index=False)