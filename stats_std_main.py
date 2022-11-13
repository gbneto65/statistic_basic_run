# trial statistical analyser
# repository at Github


import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import itertools
from scipy.stats import iqr
import numpy as np


# get setup parameters from excel file - setup sheet
wb = openpyxl.load_workbook('data.xlsx')
sheet_setup = wb['setup']
trial_id = sheet_setup['B3'].value
trial_title = sheet_setup['B4'].value
group_name = sheet_setup['B5'].value # should be text
variable_name = sheet_setup['B6'].value # should be text
number_treatments = sheet_setup['B7'].value # defined the number of treatments
outlier_iqr_multiplier = sheet_setup['B15'].value



# read excel file
df = pd.read_excel('data.xlsx')
print(df)


# basic descriptive stat
print (df.describe(percentiles=[.5]).applymap('{:,.2f}'.format))
columns_names = list(df.columns.values)

# Skewness of data
print("\nSkewness of data\n")
for i in range (len(df.columns)):
    print (f'{columns_names[i]} ==> {round(df.iloc [:,i].skew(),2)}')


# detect outliers - IQR
outliers = []
def detect_outliers(df, i, iqr_mult):
         percent_25 = df.quantile(.25)
         percent_75 = df.quantile(.75)
         iqr = percent_75 - percent_25
         low_bond = percent_25 - iqr_mult * iqr
         high_bond = percent_75 + iqr_mult * iqr
         outliers_df = df[(df > high_bond) | (df < low_bond)]
         outliers_list = outliers_df.iloc[:1].tolist()
         return outliers_list

# outliers detection - using IQR methodology
print(f'\nOutliers by group (IQR methodology - factor = {outlier_iqr_multiplier} ):')
for i in range (len(df.columns)):
    outlier_list = detect_outliers(df.iloc[:,i], i, outlier_iqr_multiplier)
    print(f'{columns_names[i]} ==> {outlier_list}')

print()

# not properly working
def remove_outliers(df, i, iqr_mult):
    df_i = df.iloc[:, i]
    percent_25 = df.iloc[:, i].quantile(.25)
    percent_75 = df.iloc[:, i].quantile(.75)
    iqr = percent_75 - percent_25
    low_bond = percent_25 - iqr_mult*iqr
    high_bond = percent_75 + iqr_mult * iqr
    outliers_list = df_i[(df_i > high_bond) | (df_i < low_bond)]
    return outliers_list



    '''
    df_wo_out_low = df_i[df_i > (percent_25 - iqr_mult*iqr)]
    print(df_wo_out_low)
    df_wo_out_high = df_i[df_i < (percent_75 + iqr_mult * iqr)]
    print(df_wo_out_high)

    df_wo_out = [*df_wo_out_low, *df_wo_out_high]

    #df.drop=(df[df_i < np.percentile(df_i, 1) - distance].index, inplace=True)
    '''
    return

for i in range (len(df.columns)):
    df_wo_outlier = remove_outliers(df, i, outlier_iqr_multiplier)
    #print(df_wo_outlier)


# create 1 df based on all treatments
df1 = df.melt()
col_name=(list(df1.columns))


# rename columns baed on excel file info
df1=df1.rename(columns = {col_name[0]:group_name})
df1=df1.rename(columns = {col_name[1]:variable_name})
print(df1)





# boxplot
df1.boxplot(by =group_name,
           column =variable_name,
           grid = False,
           )
plt.show()


'''
sheet_data = wb['data']
group_name = sheet_data['A1'].value
var_name = sheet_data['B1'].value

print(group_name)

# read excel file
df = pd.read_excel('data.xlsx')
print(df)
# num of groups

list_groups = df[group_name].to_numpy()
print(list_groups)

df_group = df.groupby([group_name]).agg(list)
df_trans = df_group.transpose()
print(df_trans)



# descriptive stat
print (df.groupby(group_name).describe(percentiles=[.5]).applymap('{:,.2f}'.format))

# boxplot
df.boxplot(by =group_name,
           column =[var_name],
           grid = False,
           )
plt.show()





#df[var_name] = df.groupby([group_name])[var_name].transform(lambda x: x/x.sum())
#print(df)

df['IQR'] = df[group_name].map(df.groupby(group_name)[var_name].agg(iqr))
print(df)

#df['Q1'] = df[group_name].map(df.groupby(group_name)[var_name].agg([percentile(25)]))


q1 = df.groupby(group_name).quantile(.25)


print(q1)

#df['lower'] = df['IQR'] - 1.5 * df['IQR']


# method 1
#df['is_outlier'] = np.where(df['IQR'] >= df['IQR'] + 1.5 | df['IQR'] <= df['IQR'] + 1.5, 'yes', 'no')
#print(df)
#df['is_outlier'] = pd.Series('no', index=df.index).mask(df['salary']>50, 'yes')
#df['Outlier'] =


'''


