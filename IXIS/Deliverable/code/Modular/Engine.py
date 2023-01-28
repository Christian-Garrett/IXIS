import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

file1 = 'Deliverable/code/Modular/Input/DataAnalyst_Ecom_data_sessionCounts.csv'
file2 = 'Deliverable/code/Modular/Input/DataAnalyst_Ecom_data_addsToCart.csv'

# load in the session count excel file
session_counts_df = pd.read_csv(file1)
data_df = session_counts_df.copy()

# load in the add to cart excel file
cart_adds_df = pd.read_csv(file2)
add_data_df = cart_adds_df.copy()

print(data_df.info())
print('\nSample Data: \n', data_df.head())
print('\nData Description Summary: \n', data_df.describe())

# check to see if there are any duplicate records
print('\nNumber of duplicate records: ', len(data_df[data_df.duplicated()]))

# output a plot of the histograms for sessions, transactions and QTY
fig, axes = plt.subplots(1, 3, figsize=(15,5))
data_df.hist('sessions', bins=10, ax=axes[0])
data_df.hist('transactions', bins=10, ax=axes[1])
data_df.hist('QTY', bins=10, ax=axes[2])

plt.savefig("Deliverable/code/Modular/Output/session transaction qty histograms.png")
plt.clf()

# print out a list of all of the unique browser x device combinations
channel_count_df = data_df.copy()
channel_count_df['channels'] = channel_count_df['dim_browser'] + "_" + channel_count_df['dim_deviceCategory']
print('\nUnique Browser x Device Combinations: \n', channel_count_df['channels'].unique())

# print out the number of unique broswer x device combinations
print('\nTotal number of unique broswer x device combinations: \n', len(channel_count_df['channels'].unique()))

# get browser x device value counts
cat_df = data_df.copy()
cat_df = cat_df.drop(['dim_date', 'sessions', 'transactions', 'QTY'], axis=1)
val_counts = cat_df.value_counts()
val_counts_df = val_counts.to_frame()
val_counts_df = val_counts_df.reset_index()
val_counts_df.columns = ['browser', 'device', 'browser x device counts']
print('\nTop Browser x Device Value Counts: \n', val_counts_df.head())

# save a histogram plot of browser x device value counts to an output file
val_counts_df.hist(bins=50)
plt.savefig("Deliverable/code/Modular/Output/browser-device value counts histogram.png")
plt.clf()

# get browser value counts
browser_info = cat_df.filter(items=['dim_browser'])
val_counts = browser_info.value_counts()
val_counts_df = val_counts.to_frame()
val_counts_df = val_counts_df.reset_index()
val_counts_df.columns = ['browser', 'browser counts']
val_counts_df.head()
print('\nTop Browser Value Counts: \n', val_counts_df.head())

# save a histogram plot of browser value counts to an output file
val_counts_df.hist(bins=50)
plt.savefig("Deliverable/code/Modular/Output/browser value counts histogram.png")
plt.clf()

# print a list of browsers
print('\nBrowsers in the dataset: \n', data_df['dim_browser'].unique())

# print the number of browsers in the dataset
print('\nNumber of unique browsers in the dataset: ', len(data_df['dim_browser'].unique()))

# get the device value counts
device_info = cat_df.filter(items=['dim_deviceCategory'])
val_counts = device_info.value_counts()
val_counts_df = val_counts.to_frame()
val_counts_df = val_counts_df.reset_index()
val_counts_df.columns = ['device', 'device counts']
print('\nDevice Value Counts: ', val_counts_df.head())

# save a pie chart of the device value counts to an output file
val_counts_df = val_counts_df.set_index('device')
val_counts_df.plot.pie(y='device counts', figsize=(5, 5))
plt.savefig("Deliverable/code/Modular/Output/device value counts histogram.png")
plt.clf()

# print the ATC data
print('\nAdd To Cart Data: \n', add_data_df)


'''
Create a Monthly Transaction Totals View

'''
dt_copy_df = data_df.copy()
# change to datetime format
dt_copy_df['dim_date'] = pd.to_datetime(dt_copy_df['dim_date'])
monthly_transactions_df = dt_copy_df.copy()
monthly_transactions_df.set_index('dim_date', inplace=True)
# resample the data to monthly figures
monthly_transactions_df = monthly_transactions_df.resample('M').sum().reset_index()
# change the date info to a month-year format
monthly_transactions_df['month_year'] = monthly_transactions_df['dim_date'].dt.to_period('M')
monthly_transactions_df = monthly_transactions_df.drop(['dim_date', 'sessions', 'QTY'], axis=1)
monthly_transactions_df.columns = ['monthly_total_transactions', 'month_year']
# print monthly transaction totals
print('\nAdd To Cart Data: \n', monthly_transactions_df)


'''
Add eCommerce Conversion Rate (ECR) field to the primary 
dataframe and convert the date to datetime 'month-year' format

'''
primary_table_df = dt_copy_df.copy()
primary_table_df['dim_date'] = pd.to_datetime(primary_table_df['dim_date'])
primary_table_df['month_year'] = primary_table_df['dim_date'].dt.to_period('M')
# avoid divide by zero errors
primary_table_df['ECR'] = np.where(primary_table_df['sessions'] == 0, 0, primary_table_df['transactions'] / primary_table_df['sessions'])
primary_table_df = primary_table_df.drop(columns='dim_date')


'''
Convert the date to a datetime 'month-year' format for the Add to Cart data

'''
add_data_df['dim_month'] = add_data_df['dim_month'].apply(lambda x: '{:02}'.format(x))
add_data_df['month_year'] = add_data_df['dim_year'].astype(str) + add_data_df['dim_month'].astype(str)
add_data_df['month_year'] = pd.to_datetime(add_data_df['month_year'], format='%Y%m')
add_data_df['month_year'] = add_data_df['month_year'].dt.to_period('M')
add_data_df = add_data_df.drop(columns=['dim_year', 'dim_month'])

'''
Output a grouped bar graph showing the add to cart counts vs transactions

'''
x = monthly_transactions_df.copy()
x.columns = ['count', 'month']
x['type'] = "transactions"
y = add_data_df.copy()
y.columns = ['count', 'month']
y['type'] = 'cart'
bar_group_df = x.append(y)

sns.barplot(x='month', y='count', hue='type', data=bar_group_df)
plt.xticks(rotation=45)
plt.title("monthly cart adds and transaction totals")

plt.savefig("Deliverable/code/Modular/Output/transaction vs add to car bar graph.png")
plt.clf()


'''
Create a full table that includes the ATC data as well as the total monthly transactions

'''
full_table_df = pd.merge(primary_table_df, add_data_df, on='month_year')
full_table_df = pd.merge(full_table_df, monthly_transactions_df, on='month_year')


'''
Create the first excel output sheet deliverable by aggregating the data on year and device category

'''

sheet1 = full_table_df.groupby(['month_year', 'dim_deviceCategory'])['sessions', 'transactions', 'QTY', 'ECR'].sum()
sheet1 = sheet1.round(2)
sheet1 = sheet1.reset_index()
sheet1['month_year'] = sheet1['month_year'].astype(str)
sheet1


'''
Narrow down to the last 2 months of the full data set, calculate the % transactions for each
record using the monthly totals and use that to calculate the relative number of cart adds

'''
full_table_df.set_index('month_year', inplace=True)
recent_data_df = full_table_df.copy()
recent_data_df = recent_data_df['2013-05':'2013-06']
recent_data_df.reset_index(inplace=True)
recent_data_df['%transactions'] = recent_data_df['transactions'] / recent_data_df['monthly_total_transactions']
recent_data_df['cart_adds'] = recent_data_df['%transactions'] * recent_data_df['addsToCart']
# print the table with the updated calculations
print('\nRelative Inference for Add to Cart Data: \n', recent_data_df.head())


'''
Aggregate the data by date, device and browser, then combine the device_browser 
columns to make processing easier for the month over month calculations 

'''
agg_recent_df = recent_data_df.groupby(['month_year', 'dim_deviceCategory', 'dim_browser'])['sessions', 'transactions', 'QTY', 'ECR', 'cart_adds'].sum()
agg_recent_df.reset_index(inplace=True)
aug_recent_df = agg_recent_df.copy()
aug_recent_df['combo'] = aug_recent_df['dim_deviceCategory'] + "_" + aug_recent_df['dim_browser'] 
aug_recent_df = aug_recent_df.round(2)
aug_recent_df.head(10)


'''
Find the monthly differences for the current and previous month

'''

monthly_pivot_df = pd.pivot_table(aug_recent_df, values=['sessions', 'transactions', 'QTY', 'ECR', 'cart_adds'], index=['month_year'], columns=['combo'], aggfunc=np.sum, fill_value=0)
# use a pivot table in conjunction with the diff function to calulate the monthly differences
monthly_diff_df = monthly_pivot_df.diff()
monthly_diff_pivot_df = monthly_diff_df.copy()
monthly_diff_pivot_df.reset_index(inplace=True)
# use the melt function and drop records containing null, leaving only the current month records
monthly_diff_melt_df = pd.melt(monthly_diff_pivot_df, id_vars=[('month_year', '')], value_vars=monthly_diff_pivot_df.columns.tolist()[1:])
monthly_diff_melt_df=monthly_diff_melt_df.dropna(axis=0)
cols = ['date','type','combo', 'value']
monthly_diff_melt_df.columns = cols
# pivot back to the original data orientation
full_mxm_df = pd.pivot_table(monthly_diff_melt_df, values=['value'], index=['combo'], columns=['type'], aggfunc=np.sum, fill_value=0)
full_mxm_df = full_mxm_df.reset_index()
new_cols = ['combo', 'aECR', 'aQTY', 'aCart_Adds', 'aSessions', 'aTransactions']
full_mxm_df.columns = new_cols
# add the date info back in manually since it was unnecessary to the pivot transformations
full_mxm_df['curr_month'] = '2013-06'
full_mxm_df.head()


'''
Filter down to current month's records and join with the previous month's differences

'''
aug_curr_df = aug_recent_df.loc[aug_recent_df['month_year'] == '2013-06']
curr_aug_df = aug_curr_df.copy()
# remove the date column so it's not duplicated in the join
curr_aug_df.drop(['month_year'], axis=1, inplace=True)
momo_table_df = full_mxm_df.set_index('combo').join(curr_aug_df.set_index('combo'))
momo_table_df.reset_index(inplace=True)
momo_table_df = momo_table_df.fillna(0)


'''
Add in (calculate) the previous month's values using the difference info that was generated,
then calculate the relative change info to create the sheet 2 excel output deliverable.

'''
# previous month data calculations
momo_table_df['pSessions'] = momo_table_df['sessions'] - momo_table_df['aSessions'] 
momo_table_df['pTransactions'] = momo_table_df['transactions'] - momo_table_df['aTransactions'] 
momo_table_df['pQTY'] = momo_table_df['QTY'] - momo_table_df['aQTY'] 
momo_table_df['pECR'] = momo_table_df['ECR'] - momo_table_df['aECR']
momo_table_df['pCart_Adds'] = momo_table_df['cart_adds'] - momo_table_df['aCart_Adds']
# relative change data calculations
momo_table_df['rSessions'] = np.where((momo_table_df['sessions'] != 0) & (momo_table_df['pSessions'] == 0), momo_table_df['sessions'] * 100, ((momo_table_df['sessions'] - momo_table_df['pSessions']) / momo_table_df['pSessions']) * 100)
momo_table_df['rTransactions'] = np.where((momo_table_df['transactions'] != 0) & (momo_table_df['pTransactions'] == 0), momo_table_df['transactions'] * 100, ((momo_table_df['transactions'] - momo_table_df['pTransactions']) / momo_table_df['pTransactions']) * 100)
momo_table_df['rQTY'] = np.where((momo_table_df['QTY'] != 0) & (momo_table_df['pQTY'] == 0), momo_table_df['QTY'] * 100, ((momo_table_df['QTY'] - momo_table_df['pQTY']) /  momo_table_df['pQTY']) * 100)
momo_table_df['rECR'] = np.where((momo_table_df['ECR'] != 0) & (momo_table_df['pECR'] == 0), momo_table_df['ECR'] * 100, ((momo_table_df['ECR'] - momo_table_df['pECR']) / momo_table_df['pECR']) * 100)
momo_table_df['rCart_Adds'] = np.where((momo_table_df['cart_adds'] != 0) & (momo_table_df['pCart_Adds'] == 0), momo_table_df['cart_adds'] * 100, ((momo_table_df['cart_adds'] - momo_table_df['pCart_Adds']) / momo_table_df['pCart_Adds']) * 100)
# remove any formatting artifacts appropriately
momo_table_df = momo_table_df.fillna(0)
momo_table_df = momo_table_df.round(2)
momo_table_df = momo_table_df.replace([np.inf, -np.inf], 0)
# restore the device and browser data to separate columns
momo_table_df[['device', 'browser']] = momo_table_df['combo'].str.split('_', expand=True)
momo_table_df.drop(['combo'], axis=1, inplace=True)
# reorder and update column names
momo_columns = ['curr_month', 'device', 'browser', 'sessions', 'pSessions', 'aSessions', 'rSessions', 'cart_adds', 'pCart_Adds', 'aCart_Adds', 'rCart_Adds', 'transactions', 'pTransactions', 'aTransactions', 'rTransactions', 'QTY', 'pQTY', 'aQTY', 'rQTY', 'ECR', 'pECR', 'aECR', 'rECR'] 
sheet2 = momo_table_df.filter(items=momo_columns)
sheet2_cols = ['Curr_Month', 'Device', 'Browser', 'New_Sessions', 'Old_Sessions', 'Session_Diff', 'Session_%Change', 'New_CartAdds', 'Old_CartAdds', 'CartAdds_Diff', 'CartAdds_%Change', 'New_Transactions', 'Old_Transactions', 'Transactions_Diff', 'Transactions_%Change', 'New_QTY', 'Old_QTY', 'QTY_Diff', 'QTY_%Change', 'New_ECR', 'Old_ECR', 'ECR_Diff', 'ECR_%Change']
sheet2.columns = sheet2_cols


# output the required deliverable files in .xslx format
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
ws2 = wb.create_sheet("Sheet2")
ws2.title = "Sheet2"

for r in dataframe_to_rows(sheet1, index=True, header=True):
    ws.append(r)

for r in dataframe_to_rows(sheet2, index=True, header=True):
    ws2.append(r)

wb.save("Deliverable/code/Modular/Output/Online Retail Performance Analysis.xlsx")



sheet1_viz_df = sheet1.copy()
'''
Output an eCommerce Conversion Rate by device line graph

'''
desktop_ecr_df = sheet1_viz_df.loc[sheet1_viz_df['dim_deviceCategory'] == 'desktop']
desktop_ecr_df = desktop_ecr_df.filter(items=['month_year', 'ECR'])
desktop_ecr_df = desktop_ecr_df.set_index('month_year')
mobile_ecr_df = sheet1_viz_df.loc[sheet1_viz_df['dim_deviceCategory'] == 'mobile']
mobile_ecr_df = mobile_ecr_df.filter(items=['month_year', 'ECR'])
mobile_ecr_df = mobile_ecr_df.set_index('month_year')
tablet_ecr_df = sheet1_viz_df.loc[sheet1_viz_df['dim_deviceCategory'] == 'tablet']
tablet_ecr_df = tablet_ecr_df.filter(items=['month_year', 'ECR'])
tablet_ecr_df = tablet_ecr_df.set_index('month_year')

df2 = pd.merge(desktop_ecr_df, mobile_ecr_df, left_index=True, right_index=True)
ECR_df = pd.merge(df2, tablet_ecr_df, left_index=True, right_index=True)
ECR_df.columns = ['desktop', 'mobile', 'tablet']

plt.figure(figsize=(12, 5), dpi=150)
ECR_df['desktop'].plot(label='Desktop', color='green')
ECR_df['mobile'].plot(label='Mobile')
ECR_df['tablet'].plot(label='Tablet')
plt.title('e-Commerce Conversion Rate by Device')
plt.xlabel('Months')
plt.legend()
plt.savefig("Deliverable/code/Modular/Output/ECR by Device Line Graphs.png")
plt.clf()


'''
Output a Quantity per Transaction line graph

'''
sheet1_viz_df['QTY per Transaction'] = sheet1_viz_df['QTY'] / sheet1_viz_df['transactions']

desktop_qpt_df = sheet1_viz_df.loc[sheet1_viz_df['dim_deviceCategory'] == 'desktop']
desktop_qpt_df = desktop_qpt_df.filter(items=['month_year', 'QTY per Transaction'])
desktop_qpt_df = desktop_qpt_df.set_index('month_year')
mobile_qpt_df = sheet1_viz_df.loc[sheet1_viz_df['dim_deviceCategory'] == 'mobile']
mobile_qpt_df = mobile_qpt_df.filter(items=['month_year', 'QTY per Transaction'])
mobile_qpt_df = mobile_qpt_df.set_index('month_year')
tablet_qpt_df = sheet1_viz_df.loc[sheet1_viz_df['dim_deviceCategory'] == 'tablet']
tablet_qpt_df = tablet_qpt_df.filter(items=['month_year', 'QTY per Transaction'])
tablet_qpt_df = tablet_qpt_df.set_index('month_year')

df2 = pd.merge(desktop_qpt_df, mobile_qpt_df, left_index=True, right_index=True)
QPT_df = pd.merge(df2, tablet_qpt_df, left_index=True, right_index=True)
QPT_df.columns = ['desktop', 'mobile', 'tablet']

plt.figure(figsize=(12, 5), dpi=150)
QPT_df['desktop'].plot(label='Desktop', color='green')
QPT_df['mobile'].plot(label='Mobile')
QPT_df['tablet'].plot(label='Tablet')
plt.title('Quantity per Transaction by Device')
plt.xlabel('Months')
plt.legend()
plt.savefig("Deliverable/code/Modular/Output/QPT by Device Line Graphs.png")
plt.clf()


sheet2_viz_df = sheet2.copy()
'''
Print out records with no session activity in the current month, by device and browser

'''
No_Sessions_df = sheet2_viz_df.filter(items=['Device', 'Browser', 'New_Sessions', 'Old_Sessions'])
No_Sessions_df = No_Sessions_df.loc[No_Sessions_df['New_Sessions'] == 0]
No_Sessions_df = No_Sessions_df.reset_index().drop(['index'], axis=1)
No_Sessions_df = No_Sessions_df.sort_values(by=['Old_Sessions'], ascending=False)
No_Sessions_df.index = np.arange(1, len(No_Sessions_df) + 1)
print('\nRecords with no session activity in the current month: \n', No_Sessions_df)


'''
Print records with no transactions in the current month, by device and browser

'''
No_Transactions_df = sheet2_viz_df.filter(items=['Device', 'Browser', 'New_Transactions', 'Old_Transactions'])
No_Transactions_df = No_Transactions_df.loc[No_Transactions_df['New_Transactions'] == 0]
No_Transactions_df = No_Transactions_df.reset_index().drop(['index', 'New_Transactions'], axis=1)
No_Transactions_df = No_Transactions_df.sort_values(by=['Old_Transactions'], ascending=False)
No_Transactions_df.index = np.arange(1, len(No_Transactions_df) + 1)
No_Transactions_df.columns = ['Device', 'Browser', 'Previous Month Transactions']
print('\nRecords with no transactions in the current month: \n', No_Transactions_df)


'''
Print records with the most transactions, by device and browser, in the current month

'''
Most_Transactions_df = sheet2_viz_df.filter(items=['Device', 'Browser', 'New_Transactions', 'Transactions_%Change'])
Most_Transactions_df = Most_Transactions_df.reset_index().drop(['index'], axis=1)
Most_Transactions_df = Most_Transactions_df.sort_values(by=['New_Transactions'], ascending=False)
Most_Transactions_df.index = np.arange(1, len(Most_Transactions_df) + 1)
print('\nRecords with the most transactions in the current month: \n', Most_Transactions_df.head(20))


'''
Print records with the highest QTY per transaction, by device and browser, in the current month

'''
Trans_Cart_df = sheet2_viz_df.filter(items=['Device', 'Browser', 'New_QTY', 'New_Transactions'])
Trans_Cart_df['QPT'] = np.where(Trans_Cart_df['New_Transactions'] == 0, 0, Trans_Cart_df['New_QTY'] / Trans_Cart_df['New_Transactions'])
Trans_Cart_df = Trans_Cart_df.round(2)
QPT_df = Trans_Cart_df.sort_values(by=['QPT'], ascending=False)
QPT_df.index = np.arange(1, len(QPT_df) + 1)
print('\nRecords with the highest qty per transaction in the current month: \n', QPT_df.head(25))


'''
Best ECR percentage improvement, by device and browser, in the current month

'''
ECR_change_df = sheet2_viz_df.filter(items=['Device', 'Browser', 'ECR_%Change', 'New_Transactions', 'New_QTY'])
ECR_change_df = ECR_change_df.sort_values(by=['ECR_%Change'], ascending=False)
ECR_Top5_df = ECR_change_df.head()
ECR_Top5_df = ECR_Top5_df.reset_index().drop(['index'], axis=1)
ECR_Top5_df.index = np.arange(1, len(ECR_Top5_df) + 1)
print('\nRecords with the best ECT % improvement in the current month: \n', ECR_Top5_df)


'''
Highest ECR percentage reduction, by device and browser, in the current month

'''
ECR_Worst5_df = ECR_change_df.tail()
ECR_Worst5_df = ECR_Worst5_df.reset_index().drop(['index'], axis=1)
ECR_Worst5_df.sort_values(by=['ECR_%Change'], inplace=True)
ECR_Worst5_df.index = np.arange(1, len(ECR_Worst5_df) + 1)
print('\nRecords with the highest ECR % reduction in the current month: \n', ECR_Worst5_df)


'''
Output ECR percentage change by device boxplots

'''
fig, ax = plt.subplots(figsize=(10, 7))
sns.set_style("whitegrid")
sns.boxplot(x = 'Device', y = 'ECR_%Change', data = sheet2_viz_df)

plt.savefig("Deliverable/code/Modular/Output/ECR % Change by Device Boxplots.png")
plt.clf()


'''
Print the highest ECR, by device and browser, for the current month

'''
ECR_raw_df = ECR_raw_df.sort_values(by=['New_ECR'], ascending=False)
ECR_raw_df.columns = ['Device', 'Browser', 'Curr ECR']

ECR_Raw_Top10_df = ECR_raw_df.head(10)
ECR_Raw_Top10_df = ECR_Raw_Top10_df.reset_index().drop(['index'], axis=1)
ECR_Raw_Top10_df.index = np.arange(1, len(ECR_Raw_Top10_df) + 1)
print('\nRecords with the highest ECR, by device and browser, in the current month: \n', ECR_Raw_Top10_df)



'''
Output ECR by device boxplots

'''
fig, ax = plt.subplots(figsize=(10, 7))
sns.set_style("whitegrid")
sns.boxplot(x = 'Device', y = 'New_ECR', data = sheet2_viz_df)

plt.savefig("Deliverable/code/Modular/Output/ECR by Device Boxplots.png")
plt.clf


