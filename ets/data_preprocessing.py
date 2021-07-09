#%% MODULE IMPORT
import pandas as pd
import openpyxl as xl
import re
import os
from functools import reduce


#%% EXTRACT tCO2 FCN
def ext_tCO2(data, company, tCO2, year, category='지정구분'):
    data_tmp = pd.DataFrame()
    
    data_tmp['company'] = [re.sub(r'주식회사|\(주\)|\(유\)|유한회사|㈜|\s', '', i) for i in data[company]]
    
    data_tmp['category'] = data[category]
    
    tCO2_year = 'tCO2_'+str(year)
    data_tmp[tCO2_year] = data[tCO2]
    data_tmp[tCO2_year] = pd.to_numeric(data_tmp[tCO2_year], errors='coerce')

    data_tmp = data_tmp.drop_duplicates()
    data_tmp = data_tmp.reset_index(drop=True)
    return data_tmp



#%% EXTRACT tCO2 

os.chdir("E:\\Data\\greenhouse_gas_emissions\\업체별 명세서 주요정보")
print("Current Working Directory ", os.getcwd())

data2011 = pd.read_excel("2011년 업체별 명세서 주요정보.xlsx",header = 1)
data2011 = ext_tCO2(data2011, company='법인명', category='지정구분',tCO2='온실가스 배출량(tCO₂eq)', year=2011)
temp = data2011.loc[data2011.duplicated(['company','tCO2_2011']),'company']
for i in temp:
    if len(data2011.loc[data2011['company']==i,]) == 2:
        data2011.loc[data2011['company']==i,['category']] = '사업장' # 포스코피앤에스는 어차피 주소데이터가 없기 때문에 없어질 예정
data2011 = data2011.drop_duplicates()
data2011 = data2011.reset_index(drop=True)

data2012 = pd.read_excel("2012년 업체별 명세서 주요정보.xlsx",header = 1)
data2012 = ext_tCO2(data2012, company='법인명', category='지정구분',tCO2='온실가스 배출량(tCO₂eq)', year=2012)

data2013 = pd.read_excel("2013년 업체별 명세서 주요정보.xlsx",header = 1)
data2013 = ext_tCO2(data2013, company='관리업체', category='지정구분',tCO2='온실가스 배출량(tCO2)', year=2013)

data2014 = pd.read_excel("2014년 업체별 명세서 주요정보.xlsx",header = 1)
data2014 = ext_tCO2(data2014, company='업체명', category='지정구분',tCO2='온실가스 배출량(tCO₂eq)', year=2014)
modify = {'company':'삼성물산','category':'사업장','tCO2_2014':124767} # 같은 년도에 category가 같은 삼성물산이 두개임. 다른 년도와 배출량의 양을 비교했을 때 합치는게 맞다고 판단.
data2014 = data2014.append(modify, ignore_index=True)
data2014 = data2014.drop_duplicates(['company','category'], keep='last')
data2014.reset_index(drop=True)

data2015 = pd.read_excel("2015년 업체별 명세서 주요정보.xlsx",header = 1)
data2015 = ext_tCO2(data2015, company='업체명', category='지정구분',tCO2='온실가스 배출량(tCO₂eq)', year=2015)

data2016 = pd.read_excel("2016년 업체별 명세서 주요정보.xlsx",header = 1)
data2016 = ext_tCO2(data2016, company='관리업체', category='지정구분',tCO2='온실가스 배출량(tCO2)', year=2016)
data2016.loc[data2016['tCO2_2016']==25146.0,'category'] = '사업장' # 다른 년도의 경원여객자동차 행과 비교했을때 사업장으로 바꾸는것이 맞다고 판단.
data2016 = data2016.reset_index(drop=True)

data2017 = pd.read_excel("2017년 업체별 명세서 주요정보.xls",header = 1)
data2017 = ext_tCO2(data2017, company='관리업체', category='지정구분',tCO2='온실가스 배출량(tCO2)', year=2017)

data2018 = pd.read_excel("2018년 업체별 명세서 주요정보.xls",header = 1)
data2018 = ext_tCO2(data2018, company='관리업체', category='지정구분',tCO2='온실가스 배출량(tCO2)', year=2018)

data2019 = pd.read_excel("2019년 업체별 명세서 주요정보.xls",header = 1)
data2019 = ext_tCO2(data2019, company='관리업체', category='지정구분',tCO2='온실가스 배출량(tCO2)', year=2019)

list_tmp = [data2011,data2012,data2013,data2014,data2015,data2016,data2017,data2018,data2019]
data_company = reduce(lambda left, right: pd.merge(left, right, how='outer', on=['company',
            'category']),list_tmp)



#%% EXTRACT ADDRESS FCN

def ext_address(data, company, category, address):
    temp = pd.DataFrame()
    
    temp['company'] = [re.sub(r'주식회사|\(주\)|\(유\)|유한회사|㈜|\s', '', i) for i in data[company]]
    
    temp['category'] = data[category]

    temp['address_main'] = data[address].str.split(' ').str[0]
    
    temp['address_sub'] = data[address].str.split(' ').str[1]
    
    return temp



# %% EXTRACT ADDRESS
os.chdir('E:\\Data\\greenhouse_gas_emissions\\업체별 주소')
print('Current Working Directory ', os.getcwd())

address1 = pd.read_excel("목표관리대상업체.xls")
address1 = ext_address(data=address1, company='관리업체명', category='지정구분', address='주소')

address2 = pd.read_excel("할당대상업체_1차.xls")
address2 = ext_address(data=address2, company='업체명', category='적용기준', address='소재지')

address3 = pd.read_excel("할당대상업체_2차.xls")
address3 = ext_address(data=address3, company='업체명', category='적용기준', address='소재지')

address4 = pd.read_excel("할당대상업체_3차.xls")
address4 = ext_address(data=address4, company='업체명', category='적용기준', address='소재지')

address = pd.concat([address2, address3, address4, address1], axis=0)

address.loc[address['address_main'] == '수원시','address_main'] = '경기도' 
address.loc[address['address_main'] == '대전관역시','address_main'] = '대전광역시'
address.loc[address['address_main'] == '서을특별시','address_main'] = '서울특별시'
address.loc[address['address_main'] == '서울시','address_main'] = '서울특별시'
address.loc[address['address_main'] == '서울','address_main'] = '서울특별시'
address.loc[address['address_main'] == '부산','address_main'] = '부산광역시'
address.loc[address['address_main'] == '인천','address_main'] = '인천광역시'
address.loc[address['address_main'] == '경기','address_main'] = '경기도'
address.loc[address['address_main'] == '경남','address_main'] = '경상남도'
address.loc[address['address_main'] == '전북','address_main'] = '전라북도'
address.loc[address['address_main'] == '충북','address_main'] = '충청북도'
address.loc[address['address_main'] == '강원','address_main'] = '강원도'
address.loc[address['address_main'] == '대구','address_main'] = '대구광역시'
address.loc[address['address_main'] == '충남','address_main'] = '충청남도'
address.loc[address['address_main'] == '전남','address_main'] = '전라남도'
address.loc[address['address_main'] == '울산','address_main'] = '울산광역시'

address.loc[address['address_sub'] == '전동면','address_sub'] = '세종시'
address.loc[address['address_sub'] == '전의면','address_sub'] = '세종시'
address.loc[address['address_sub'] == '장안구','address_sub'] = '수원시'
address.loc[address['address_sub'] == '부강면','address_sub'] = '세종시'
address.loc[address['address_sub'] == '','address_sub'] = '세종시'
address.loc[address['address_sub'] == '조치원읍','address_sub'] = '세종시'
address.loc[address['address_sub'] == '한누리대로','address_sub'] = '세종시'
address.loc[address['address_sub'] == '특별시','address_sub'] = '영등포구'

address = address.drop_duplicates()
address = address.reset_index(drop=True)

address_tmp = address.drop_duplicates(['company','address_sub'], keep='first')

address1 = address_tmp.drop_duplicates(['company'], keep=False) # 동일회사명이 없는 subset => company 기준으로 merge

data_tmp1 = address_tmp.drop(address1.index) # 동일 회사명이 있는 subset

'''
a = data_tmp1.loc[data_tmp1.duplicated(['company']),'company'].to_list()
list_tmp=[]
for i in a:
    if data_tmp1.loc[data_tmp1['company']==i,].shape[0]>=3: list_tmp.append(i)
print(list_tmp)
'''

address2 = data_tmp1
address2 = address2.drop(data_tmp1.loc[data_tmp1['company']=='코오롱글로텍구미공장',].index)
address2 = address2.drop(data_tmp1.loc[data_tmp1['company']=='코카콜라음료여주공장',].index)
address2 = address2.drop(data_tmp1.loc[data_tmp1['company']=='삼성물산',].index)
address2 = address2.drop(data_tmp1.loc[data_tmp1['company']=='만호제강',].index)
address2 = address2.drop(data_tmp1.loc[data_tmp1['company']=='아세아제지',].index)
address2 = address2.drop(data_tmp1.loc[data_tmp1['company']=='한국실리콘',].index)
address2 = address2.drop(data_tmp1.loc[data_tmp1['company']=='삼영화학공업',].index)

#print([address2.shape[0],data_tmp1.shape[0]])


#%% GRDP
os.chdir("E:\\Data\\greenhouse_gas_emissions\\grdp")
print("Current Working Directory ",os.getcwd())

gangwon = pd.read_excel("강원도 지역내총생산 20112018.xlsx")

gyeonggi1 = pd.read_excel('경기도 지역내총생산 2011-2014.xlsx')
gyeonggi2 = pd.read_excel('경기도 지역내총생산 2015-2018.xlsx')
gyeonggi = pd.merge(gyeonggi1,gyeonggi2, on=['address_main','address_sub'])

gyeongnam = pd.read_excel('경상남도 지역내총생산 2011-2017.xlsx')

gyeongbuk = pd.read_excel('경상북도 지역내 총생산 2011-2017.xlsx.')

gwangju = pd.read_excel('광주광역시 지역내총생산 2011-2018.xlsx')

daeku = pd.read_excel('대구광역시 지역내총생산 2011-2018.xlsx')
daeku['address_sub'] = daeku['address_main']
daeku['address_main'] = '대구광역시'

daejeon1 = pd.read_excel('대전광역시 지역내총생산 1.xlsx')
daejeon2 = pd.read_excel('대전광역시 지역내총생산 2.xlsx')
daejeon = pd.merge(daejeon1, daejeon2, on=['address_main','address_sub'])

busan1 = pd.read_excel('부산광역시 지역내총생산 1.xlsx')
busan2 = pd.read_excel('부산광역시 지역내총생산 2.xlsx')
busan = pd.merge(busan1, busan2, on=['address_main','address_sub'])

seoul = pd.read_excel('서울특별시 지역내총생산.xlsx')

ulsan1 = pd.read_excel('울산광역시 지역내총생산 1.xlsx')
ulsan2 = pd.read_excel('울산광역시 지역내총생산 2.xlsx')
ulsan = pd.merge(ulsan1, ulsan2, on=['address_main','address_sub'])

incheon1 = pd.read_excel('인천광역시 지역내총생산 1.xlsx')
incheon2 = pd.read_excel('인천광역시 지역내총생산 2.xlsx')
incheon = pd.merge(incheon1, incheon2, on=['address_main','address_sub'])

jeonnam = pd.read_excel('전라남도 지역내총생산.xlsx')

jeonbuk1 = pd.read_excel('전라북도 지역내총생산 1.xlsx')
jeonbuk2 = pd.read_excel('전라북도 지역내총생산 2.xlsx')
jeonbuk = pd.merge(jeonbuk1, jeonbuk2, on=['address_main','address_sub'])

jaeju = pd.read_excel('제주특별자치도 지역내총생산.xlsx')

chungnam1 = pd.read_excel('충청남도 지역내총생산 1.xlsx')
chungnam2 = pd.read_excel('충청남도 지역내총생산 2.xlsx')
chungnam = pd.merge(chungnam1, chungnam2, on=['address_main','address_sub'])

chungbuk = pd.read_excel('충청북도_경제활동별_지역내총생산_20210526093536.xlsx')

region = [gangwon, gyeonggi, gyeongnam, gyeongbuk, gwangju, daeku, daejeon, busan, seoul,
        ulsan, incheon, jeonnam, jeonbuk, jaeju, chungnam, chungbuk]
grdp = pd.concat(region, axis=0)


# %% EXTRACT TAX/EXPENDITURE FCN
def budget(data, col_name):    
    data = data.drop(data.loc[data['자치단체'].str.contains('본청$'),].index)
    data = data.drop(data.loc[data['자치단체'].str.contains('계$'),].index)
    data = data.reset_index()

    data_tmp = pd.DataFrame()

    data_tmp['address_main'] = [('').join(list(i)[0:2]) for i in data['자치단체']]

    data_tmp.loc[data_tmp['address_main'] == '서울','address_main'] = '서울특별시'
    data_tmp.loc[data_tmp['address_main'] == '부산','address_main'] = '부산광역시'
    data_tmp.loc[data_tmp['address_main'] == '인천','address_main'] = '인천광역시'
    data_tmp.loc[data_tmp['address_main'] == '경기','address_main'] = '경기도'
    data_tmp.loc[data_tmp['address_main'] == '경남','address_main'] = '경상남도'
    data_tmp.loc[data_tmp['address_main'] == '전북','address_main'] = '전라북도'
    data_tmp.loc[data_tmp['address_main'] == '충북','address_main'] = '충청북도'
    data_tmp.loc[data_tmp['address_main'] == '강원','address_main'] = '강원도'
    data_tmp.loc[data_tmp['address_main'] == '대구','address_main'] = '대구광역시'
    data_tmp.loc[data_tmp['address_main'] == '충남','address_main'] = '충청남도'
    data_tmp.loc[data_tmp['address_main'] == '전남','address_main'] = '전라남도'
    data_tmp.loc[data_tmp['address_main'] == '울산','address_main'] = '울산광역시'

    data_tmp['address_sub'] = [('').join(list(i)[2:]) for i in data['자치단체']]

    data_tmp[col_name] = data['합계']

    return data_tmp


# %% EXTRACT TAX   
os.chdir('E:\Data\greenhouse_gas_emissions\세입')
print('Current Working Directory ', os.getcwd())

tax2011 = pd.read_excel('세입2011.xlsx', header = 1)
tax2011 = budget(tax2011, 'tax_2011')

tax2012 = pd.read_excel('세입2012.xlsx', header = 1)
tax2012 = budget(tax2012, 'tax_2012')

tax2013 = pd.read_excel('세입2013.xlsx', header = 1)
tax2013 = budget(tax2013, 'tax_2013')

tax2014 = pd.read_excel('세입2014.xlsx', header = 1)
tax2014 = budget(tax2014, 'tax_2014')

tax2015 = pd.read_excel('세입2015.xlsx', header = 1)
tax2015 = budget(tax2015, 'tax_2015')

tax2016 = pd.read_excel('세입2016.xlsx', header = 1)
tax2016 = budget(tax2016, 'tax_2016')

tax2017 = pd.read_excel('세입2017.xlsx', header = 1)
tax2017 = budget(tax2017, 'tax_2017')

tax2018 = pd.read_excel('세입2018.xlsx', header = 1)
tax2018 = budget(tax2018, 'tax_2018')

tax2019 = pd.read_excel('세입2019.xlsx', header = 1)
tax2019 = budget(tax2019, 'tax_2019')

list_tmp = [tax2011,tax2012,tax2013,tax2014,tax2015,tax2016,tax2017,tax2018,tax2019]
tax = reduce(lambda left, right: pd.merge(left,right,how='inner',on=['address_main','address_sub']),list_tmp)



#%% EXTRACT EXPENDITURE
os.chdir('E:\Data\greenhouse_gas_emissions\세출')
print('Current Working Directory ', os.getcwd())

expenditure2011 = pd.read_excel('세출2011.xlsx', header = 1)
expenditure2011 = budget(expenditure2011, 'expenditure_2011')

expenditure2012 = pd.read_excel('세출2012.xlsx', header = 1)
expenditure2012 = budget(expenditure2012, 'expenditure_2012')

expenditure2013 = pd.read_excel('세출2013.xlsx', header = 1)
expenditure2013 = budget(expenditure2013, 'expenditure_2013')

expenditure2014 = pd.read_excel('세출2014.xlsx', header = 1)
expenditure2014 = budget(expenditure2014, 'expenditure_2014')

expenditure2015 = pd.read_excel('세출2015.xlsx', header = 1)
expenditure2015 = budget(expenditure2015, 'expenditure_2015')

expenditure2016 = pd.read_excel('세출2016.xlsx', header = 1)
expenditure2016 = budget(expenditure2016, 'expenditure_2016')

expenditure2017 = pd.read_excel('세출2017.xlsx', header = 1)
expenditure2017 = budget(expenditure2017, 'expenditure_2017')

expenditure2018 = pd.read_excel('세출2018.xlsx', header = 1)
expenditure2018 = budget(expenditure2018, 'expenditure_2018')

expenditure2019 = pd.read_excel('세출2019.xlsx', header = 1)
expenditure2019 = budget(expenditure2019, 'expenditure_2019')

list_tmp = [expenditure2011,expenditure2012,expenditure2013,expenditure2014,expenditure2015,expenditure2016,expenditure2017,expenditure2018,expenditure2019]
expenditure = reduce(lambda left, right: pd.merge(left,right,how='inner',on=['address_main','address_sub']),list_tmp)


#%% EXTRACT ENERGY FCN
def extract_energy(data, year, colname = '합 계'):
    data = data.loc[data['계약종별']== colname,]
    data = data.reset_index()

    data_tmp = pd.DataFrame()
    temp = 'energy_'+ str(year)
    data_tmp['address_main'] = data['시도']
    data_tmp['address_sub'] = data['시군구']
    data_tmp[temp] = data[['1월','2월','3월','4월','5월','6월','7월','8월',
                                    '9월','10월','11월','12월']].sum(axis=1)

    return data_tmp


# %% EXTRACT ENERGY
os.chdir("E:\Data\greenhouse_gas_emissions\전력사용량")
print("Current Working Directory ", os.getcwd())

energy2011 = pd.read_excel("시군구별 전력사용량(2011년).xlsx", header=2)
energy2011 = extract_energy(energy2011, 2011, colname = '총계')

energy2012 = pd.read_excel("시군구별 전력사용량(2012년).xlsx", header=2)
energy2012 = extract_energy(energy2012, 2012, colname='총계')

energy2013 = pd.read_excel("시군구별 전력사용량(2013년).xlsx", header=2)
energy2013 = extract_energy(energy2013, 2013, colname ='총계')

energy2014 = pd.read_excel("150128 시군구별 전력사용량(2014년 1월_12월).xlsx", header=2)
energy2014 = extract_energy(energy2014, 2014)

energy2015 = pd.read_excel("시군구별 전력사용량(2015년 12월).xlsx", header=2)
energy2015 = extract_energy(energy2015, 2015)

energy2016 = pd.read_excel("시군구별 전력사용량(2016년 12월).xlsx", header=2)
energy2016 = extract_energy(energy2016, 2016)

energy2017 = pd.read_excel("8.시군구별 전력사용량(홈페이지 게시용)_201712.xlsx", header=2)
energy2017 = extract_energy(energy2017, 2017)

energy2018 = pd.read_excel("8.시군구별 전력사용량(홈페이지 게시용)_201812.xlsx", header=2)
energy2018 = extract_energy(energy2018, 2018)

energy2019 = pd.read_excel("(2019)시군구별 전력사용량.xlsx", header=2)
energy2019 = extract_energy(energy2019, 2019)

energys = [energy2011,energy2012,energy2013,energy2014,energy2015,energy2016,
            energy2017,energy2018,energy2019]
energy = reduce(lambda left, right: pd.merge(left, right, on=['address_main',
            'address_sub']),energys)


# %% EXTRACT ER
os.chdir("E:\Data\greenhouse_gas_emissions\고용지표")
print("Current Working Directory ")

data1 = pd.read_excel("고용률1.xlsx", header=1)
data_temp = pd.DataFrame()
data_temp[['address_main','ER_2011','ER_2012']] = data1[['행정구역(시도)','고용률 (%)', '고용률 (%).1']]
data1 = data_temp

data2 = pd.read_excel('고용률2.xlsx', header=1)
data_temp = pd.DataFrame()
data_temp[['address_main','ER_2013','ER_2014','ER_2015','ER_2016','ER_2017',
'ER_2018','ER_2019']] = data2[['행정구역(시도)','고용률 (%)', '고용률 (%).1', 
        '고용률 (%).2', '고용률 (%).3','고용률 (%).4', '고용률 (%).5', '고용률 (%).6']]
data2 = data_temp

ER = pd.merge(data1, data2, on=['address_main'])


# %% EXTRACT employed
os.chdir("E:\Data\greenhouse_gas_emissions\고용지표")
print("Current Working Directory ")

data1 = pd.read_excel("취업자1.xlsx", header=1)
data_temp = pd.DataFrame()
data_temp[['address_main','employed_2011','employed_2012']] = data1[['행정구역(시도)','취업자 (천명)', '취업자 (천명).1']]
data1 = data_temp

data2 = pd.read_excel('취업자2.xlsx', header=1)
data_temp = pd.DataFrame()
data_temp[['address_main','employed_2013','employed_2014','employed_2015','employed_2016','employed_2017',
'employed_2018','employed_2019']] = data2[['행정구역(시도)','취업자 (천명)', '취업자 (천명).1', '취업자 (천명).2', '취업자 (천명).3',
       '취업자 (천명).4', '취업자 (천명).5', '취업자 (천명).6']]
data2 = data_temp

employed = pd.merge(data1, data2, on=['address_main'])


# %% MERGE REGION DATA
list_tmp = [grdp,tax,expenditure,energy]
data_region = reduce(lambda left, right: pd.merge(left, right, on=['address_main',
                        'address_sub']), list_tmp)

list_tmp = [data_region, ER, employed]
data_region = reduce(lambda left, right: pd.merge(left,right,on=['address_main']), list_tmp)

data_region = data_region.reset_index(drop=True)


# %% MERGE COMPANY DATA
# %%
def extract_merge(data,tco2,company='company'):
    data_tmp = pd.DataFrame()

    data_tmp[company] = data[company]

    data_tmp[tco2] = data[tco2]
    
    return data_tmp


# %%
list_tmp = [address1,extract_merge(data2011,tco2='tCO2_2011'),extract_merge(data2012,tco2='tCO2_2012'),extract_merge(data2013,tco2='tCO2_2013'),extract_merge(data2014,tco2='tCO2_2014'),
        extract_merge(data2015,tco2='tCO2_2015'),extract_merge(data2016,tco2='tCO2_2016'),extract_merge(data2017,tco2='tCO2_2017'),extract_merge(data2018,tco2='tCO2_2018'),
        extract_merge(data2019,tco2='tCO2_2019')]

data_company1 = reduce(lambda left, right: pd.merge(left,right,how='left',on=['company']),list_tmp)

list_tmp = [address2, data2011,data2012,data2013,data2014,data2015,data2016,data2017,data2018,data2019]

data_company2 = reduce(lambda left, right: pd.merge(left, right, how='left',on=['company','category']),list_tmp)

data_company = pd.concat([data_company1,data_company2],axis=0)

data_company = data_company.dropna(how='all', subset=['tCO2_2011', 'tCO2_2012', 'tCO2_2013', 'tCO2_2014', 'tCO2_2015', 'tCO2_2016', 'tCO2_2017', 'tCO2_2018', 'tCO2_2019'])

data_company = data_company.drop_duplicates(['company','tCO2_2011','tCO2_2012','tCO2_2013','tCO2_2014','tCO2_2015','tCO2_2016','tCO2_2017','tCO2_2018','tCO2_2019'], keep=False)


# %% EXPORT
data_final = pd.merge(data_company, data_region, how="inner", on=['address_main','address_sub'])


os.chdir("E:\Data\greenhouse_gas_emissions")
print("Current Working Directory ", os.getcwd())
data_final.to_excel("ghg_emissions_v3.xlsx", index=False)


# %% Kis VALUE
os.chdir("E:\Data\greenhouse_gas_emissions")
print("Current Working Directory ", os.getcwd())

kis = pd.read_excel("KisVALUE.xlsx", header=1)

kis.columns = ['company',
    'employees_2011','employees_2012','employees_2013','employees_2014','employees_2015','employees_2016','employees_2017','employees_2018','employees_2019','employees_2020',
    'revenue_2011','revenue_2012','revenue_2013','revenue_2014','revenue_2015','revenue_2016','revenue_2017','revenue_2018','revenue_2019','revenue_2020',
    'LC_2011','LC_2012','LC_2013','LC_2014','LC_2015','LC_2016','LC_2017','LC_2018','LC_2019','LC_2020',
    'OP_2011','OP_2012','OP_2013','OP_2014','OP_2015','OP_2016','OP_2017','OP_2018','OP_2019','OP_2020']

kis['company'] = [re.sub(r'주식회사|\(주\)|\(유\)|유한회사|㈜|\s', '', i) for i in kis['company']]

kis_tmp = kis.drop_duplicates(['company'], keep=False)


# %% Kis MERGE
os.chdir("E:\Data\greenhouse_gas_emissions")
print("Current Working Directory", os.getcwd())
ghg_emissions = pd.read_excel("ghg_emissions.xlsx")

temp = pd.merge(ghg_emissions, kis_tmp, how='left', on='company')


#%% EXPORT
os.chdir("E:\Data\greenhouse_gas_emissions")
print("Current Working Directory ", os.getcwd())
temp.to_excel("ghg_emissions_v4.xlsx", index=False)

# %%
