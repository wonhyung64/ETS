#%% MODULE IMPORT
import pandas as pd
import numpy as np
import seaborn as sns
import statsmodels.api as smf
import os

from matplotlib import pyplot as plt


#%% PREPROCESSING FCN
def extract(df, year, var,  category='업체', logtrans=False):
    """extract variables from data

    Args:
        df (DataFrame): Original data
        year (int): year that u want to see
        var (str): Dependent varialbe
        category (str, optional): category of company. Defaults to '업체'.
        logtrans (bool, optional): T/F of log transform. Defaults to False.

    Returns:
        DataFrame: data frame that you will use
    """
    rdd_y = var+'_'+str(year)
    df = df[['company','category','tCO2_'+str(year-4),'tCO2_'+str(year-3),'tCO2_'+str(year-2),rdd_y]].copy()
    df = df.loc[df['category']=='업체',].copy()

    df['tCO2_'+str(year-4)] = df['tCO2_'+str(year-4)].apply(pd.to_numeric, errors='coerce')
    df['tCO2_'+str(year-3)] = df['tCO2_'+str(year-3)].apply(pd.to_numeric, errors='coerce')
    df['tCO2_'+str(year-2)] = df['tCO2_'+str(year-2)].apply(pd.to_numeric, errors='coerce')

    df['tCO2_mean'] = [np.mean(df.iloc[i,2:5]) for i in range(df.shape[0])]

    df = df.loc[df[rdd_y].notna()]
    df = df.loc[df['tCO2_mean'].notna()]

    if logtrans == True:
        df[rdd_y] = np.log(df[rdd_y])
        df['tCO2_mean'] = np.log(df['tCO2_mean'])

    return df


def get_outlier(df=None, column=None, weight=1.5):
    """IQR - 3 

    Args:
        df (DataFrame, optional): [dataframe that you will delete outlier]. Defaults to None.
        column (str, optional): [columne that you will delete outlier]. Defaults to None.
        weight (float, optional): [i dont know ...]. Defaults to 1.5.

    Returns:
        [DataFrame]: [datafrme that u will use]
    """
    quantile_25 = np.percentile(df[column].values, 25)
    quantile_75 = np.percentile(df[column].values, 75)

    IQR = quantile_75 - quantile_25
    IQR_weight = IQR*weight
    
    lowest = quantile_25 - IQR_weight
    highest = quantile_75 + IQR_weight
    
    outlier_idx = df[column][ (df[column] < lowest) | (df[column] > highest) ].index
    return outlier_idx


def th(df, threshold, order):
    """Generate column of threshold.

    Args:
        df ([DataFrame]): [data frame that u will generate threshold]
        threshold ([int]): [threshold that u set]

    Returns:
        [dataframe]: [data frame that contains threshold]
    """
    df['threshold'] = (df['tCO2_mean'] > threshold).astype(int)
    
    for i in range(order):
        df['tCO2_mean^'+str(i+1)] = df['tCO2_mean'] ** (i+1)
        df['tCO2_mean^'+str(i+1)+'*threshold'] = df['tCO2_mean^'+str(i+1)] * df['threshold']
    return df


#%% RDD FCN
def rdd(df, year, var, category='업체', log=False, bandwidth=0, order=1):
    """RDD

    Args:
        df ([DataFrame]): [Original data]
        year ([str or int]): [year that u will fit RDD]
    var ([str]): [variable that u will fit RDD]
        category (str, optional): [category that u will fit RDD]. Defaults to '업체'.
        log (bool, optional): [T\F for log transforamtion]. Defaults to False.
        bandwidth (int, optional): [x range that u will use]. Defaults to 0.
    """
    if year == 'all':
        df_tmp = pd.DataFrame()
        list_tmp = [2015,2016,2017,2018,2019]
        for i in list_tmp:
            tmp = extract(df, i, var)
            tmp.columns = ['company','category','tCO2_1','tCO2_2','tCO2_3',var,'tCO2_mean']
            df_tmp=pd.concat([df_tmp,tmp])
        df = df_tmp
        rdd_y = var        
        
    else:
        df = extract(df, year, var, logtrans=log)
        rdd_y = var+'_'+str(year)

        

    if bandwidth != 0:
        df = df.loc[df['tCO2_mean'] <= 125000+bandwidth,]
        df = df.loc[df['tCO2_mean'] >= 125000-bandwidth,]

    outlier_idx = get_outlier(df=df, column='tCO2_mean', weight=1.5)
    df.drop(outlier_idx, axis=0, inplace=True)

    outlier_idx = get_outlier(df=df, column=rdd_y, weight=1.5)
    df.drop(outlier_idx, axis=0, inplace=True)

    if log == True : threshold = np.log(125000)
    else : threshold = 125000

    df = th(df, threshold,order)    

    y = df[rdd_y]
    X = df.iloc[:,7:]
    X = smf.add_constant(X)

    rdd = smf.OLS(y, X).fit()
    under_idx = X['threshold'] == 0
    over_idx = X['threshold'] == 1
    under_y = rdd.fittedvalues[under_idx]
    over_y = rdd.fittedvalues[over_idx]

    plt.plot(X['tCO2_mean^1'].loc[under_idx], under_y,'m-')
    plt.plot(X['tCO2_mean^1'].loc[over_idx], over_y, 'y-')
    plt.scatter(df['tCO2_mean^1'],df[rdd_y])
    plt.vlines(threshold,ymin=df[rdd_y].min(),ymax=df[rdd_y].max(), colors = 'black', linestyles = ':')

    plt.xlabel('log_'*log+'avg_tCO2')
    plt.ylabel('log_'*log+rdd_y)
    plt.title('Regression Discontinuity')

    plt.show()

    print('\n================================= bandwidth :',bandwidth,'==================================\n')

    print(rdd.summary().tables[0])
    print(rdd.summary().tables[1])
# %%
os.chdir("E:\Data\greenhouse_gas_emissions")
# print("Current Working Directory" , os.getcwd())
data = pd.read_excel("ghg_emissions_v4.xlsx")

# rdd(data, year='all', var='LC', log=False,order=1)

#%%
list_bandwidth = []
for i in range(13):
    list_bandwidth.append(10000 * i)
list_var = ['employees','revenue','LC','OP']
# list_year = [2015,2016,2017,2018,2019,2020]
list_order = [1,2,3]


for o in list_order:
    for b in list_bandwidth:
        for v in list_var:
            try:
                data = pd.read_excel("ghg_emissions_v4.xlsx")
                rdd(data, 'all', v, bandwidth=b, order=o)
            except:
                print('\n\n\n\n\n', o, '차항,', 'bandwidth :', b, ',', v,"에서 에러남.",'\n\n\n\n\n')
print('끝.')

#%%
'''
    1. company (A) : 회사명
    
    2. category (B) : 지정구분

    3. address_main (C) : 사업장 주소 시도 분류

    4. address_sub (D) : 사업장 주소 시군구 분류

    5. tCO2 (E~M) : 연도별 온실가스 배출량

    6. GRDP (N~U) : 연도별 GRDP. *2019년 GRDP는 없습니다.

    7. tax (V~AD) : 연도별 시군구 분류 지자체 세입 *단위: 백만원

    8. expenditure (AE~AM) : 연도별 시군구 분류 지자체 지출 *단위: 백만원

    9. energy (AN~AV) : 연도별 시군구 분류 전력사용량 *단위: kwh

    10. ER (AW~BE) : 연도별 시도 분류 고용률 *단위: 백분위

    11. employed (BF~BN) : 연도별 시도 분류 취업자수 *단위: 천명

    12. employees (BO~BX) : 연도별 회사의 종업원 수

    13. revenue (BY~CH) : 연도별 회사의 매출액

    14. LC (CI~ CR) : 연도별 회사의 인건비

    15. OP (CS~DB) : 연도별 회사의 영업이익
'''



# %%

