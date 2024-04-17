import pandas as pd
import openpyxl
import numpy as np
from fuzzywuzzy import process
from tqdm import tqdm


cibcinfo = pd.read_excel(r'Y:\_Personal Folders\Jason\Escrow\Input\CIBC.xlsx', sheet_name='cibc')
lonewolf = pd.read_excel(r'Y:\_Personal Folders\Jason\Escrow\Input\Lonewolf.xlsx', sheet_name='lonewolf')

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\grouped_transactions.xlsx') as writer:
    cibcinfo.to_excel(writer, sheet_name='cibcinfo full', index=False)
    lonewolf.to_excel(writer, sheet_name='Lonewolf full', index=False)
with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\cibc_detailed_transaction_list.xlsx') as writer:
    cibcinfo.to_excel(writer, sheet_name='cibcinfo full', index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\cibc_achcredit_transaction_list.xlsx') as writer:
    cibcinfo.to_excel(writer, sheet_name='cibcinfo full', index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\cibc_wire_transaction_list.xlsx') as writer:
    cibcinfo.to_excel(writer, sheet_name='cibcinfo full', index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\cibc_remote_deposit_transaction_list.xlsx') as writer:
    cibcinfo.to_excel(writer, sheet_name='cibcinfo full', index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\lonewolf_achcredit_transaction_list.xlsx') as writer:
    lonewolf.to_excel(writer, sheet_name='lonewolfinfo full', index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\lonewolf_wire_transaction_list.xlsx') as writer:
    lonewolf.to_excel(writer, sheet_name='lonewolfinfo full', index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\lonewolf_remote_deposit_transaction_list.xlsx') as writer:
    lonewolf.to_excel(writer, sheet_name='lonewolfinfo full', index=False)



with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\lonewolf_detailed_transaction_list.xlsx') as writer:
    lonewolf.to_excel(writer, sheet_name='Lonewolf full', index=False)

mask = (cibcinfo['BAI Description'] == 'ACH CREDIT') & (~cibcinfo['Detail'].astype(str).str.contains('Earnnest'))

# Update 'BAI Description' based on the mask
cibcinfo.loc[mask, 'BAI Description'] = 'INCOMING WIRE TRANSFER'

cibcinfo = cibcinfo[~cibcinfo['Transaction Amount'].isna()]

values_to_match = [ 'BOOK TRANSFER DEBIT', 'OUTGOING WIRE TRANSFER']  # Replace with your specific values

# Create a boolean mask based on the conditions
mask = cibcinfo['BAI Description'].isin(values_to_match)

# Update the column value based on the mask
cibcinfo.loc[mask, 'BAI Description'] = 'OUTGOING WIRE TRANSFER'

cibcinfo['BAI Description'] = np.where(cibcinfo['BAI Description'] == 'DEPOSIT ITEM RETURNED', "REMOTE DEPOSIT",cibcinfo['BAI Description'])

cibcinfo['BAI Description'] = np.where(cibcinfo['BAI Description'] == 'BOOK TRANSFER CREDIT', "INCOMING WIRE TRANSFER",cibcinfo['BAI Description'])

categories = [ 'ACH CREDIT',  'REMOTE DEPOSIT', 'INCOMING WIRE TRANSFER']

cibcchecks = cibcinfo[cibcinfo['BAI Description'] == 'CHECK PAID']

cibcinfo['Post Date'] = cibcinfo['Post'].dt.strftime('%m-%d-%Y')

cibcdates = cibcinfo['Post Date'].unique()

combined_df = pd.DataFrame()

for cat in categories:
    new_df = cibcinfo[cibcinfo['BAI Description'] == cat]
    new_df = new_df.sort_values(by='Transaction Amount',ascending=False)


    with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\cibc_detailed_transaction_list.xlsx', engine='openpyxl', mode='a') as writer:
        for date in cibcdates:
            dates_df = new_df[new_df['Post Date'] == date]
            dates_df.to_excel(writer, sheet_name=f"{date}_{cat}_cibc", index=False)

    new_df = new_df.groupby(['Post Date', 'BAI Description'])['Transaction Amount'].sum()
    new_df = new_df.reset_index()

    new_df = new_df.rename(columns={'Transaction Amount':'CIBC Amount'})

    combined_df = pd.concat([new_df,combined_df])

achcreditcibc = cibcinfo[cibcinfo['BAI Description'] == 'ACH CREDIT']
remotecibc = cibcinfo[cibcinfo['BAI Description'] == 'REMOTE DEPOSIT']
wirecibc  = cibcinfo[cibcinfo['BAI Description'] == 'INCOMING WIRE TRANSFER']

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\cibc_achcredit_transaction_list.xlsx', engine='openpyxl', mode='a') as writer:
    for date in cibcdates:
        dates_df = achcreditcibc[achcreditcibc['Post Date'] == date]
        dates_df.to_excel(writer, sheet_name=f"{date}_achcredit_cibc", index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\cibc_wire_transaction_list.xlsx', engine='openpyxl', mode='a') as writer:
    for date in cibcdates:
        dates_df = wirecibc[wirecibc['Post Date'] == date]
        dates_df.to_excel(writer, sheet_name=f"{date}_wire_cibc", index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\cibc_remote_deposit_transaction_list.xlsx', engine='openpyxl', mode='a') as writer:
    for date in cibcdates:
        dates_df = remotecibc[remotecibc['Post Date'] == date]
        dates_df.to_excel(writer, sheet_name=f"{date}_remotedeposit_cibc", index=False)




lonewolf['BAI Description'] = np.where(lonewolf['refer'].str.contains('WIRE|ACH', case=False), "INCOMING WIRE TRANSFER", "" )
lonewolf['BAI Description'] = np.where(lonewolf['refer'].str.contains('EARN|ARN|NEST', case=False), "ACH CREDIT", lonewolf['BAI Description'] )

lonewolf['BAI Description'] = np.where(lonewolf['refer'].str[:3] == "EFT", 'OUTGOING WIRE TRANSFER', lonewolf['BAI Description'] )

condition1 = lonewolf['BAI Description'].isna() | (lonewolf['BAI Description'] == '')

# Condition 2: If 'Column2' cannot be converted to numeric
condition2 = ~pd.to_numeric(lonewolf['refer'], errors='coerce').notna()

lonewolf.loc[condition1 & condition2,'BAI Description'] = 'REMOTE DEPOSIT'

lonewolf['BAI Description'] = np.where(lonewolf['BAI Description'] == '',"CHECK PAID", lonewolf['BAI Description'] )

lonewolfchecks = lonewolf[lonewolf['BAI Description'] == 'CHECK PAID']

lonewolf['Post Date'] = lonewolf['date'].dt.strftime('%m-%d-%Y')
lonewolf_combined_df = pd.DataFrame()
for cat in categories:
    new_df = lonewolf[lonewolf['BAI Description'] == cat]
    new_df = new_df.sort_values(by='amount',ascending=False)
    with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\lonewolf_detailed_transaction_list.xlsx', engine='openpyxl', mode='a') as writer:
        for date in cibcdates:
            dates_df = new_df[new_df['Post Date'] == date]
            dates_df.to_excel(writer, sheet_name=f"{date}_{cat}_lonewolf", index=False)
    
    new_df = new_df.groupby(['Post Date', 'BAI Description'])['amount'].sum()
    new_df = new_df.reset_index()

    new_df = new_df.rename(columns={'amount':'Lonewolf Amount'})

    lonewolf_combined_df = pd.concat([new_df,lonewolf_combined_df])


achcreditlonewolf = lonewolf[lonewolf['BAI Description'] == 'ACH CREDIT']
remotelonewolf = lonewolf[lonewolf['BAI Description'] == 'REMOTE DEPOSIT']
wirelonewolf  = lonewolf[lonewolf['BAI Description'] == 'INCOMING WIRE TRANSFER']

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\lonewolf_achcredit_transaction_list.xlsx', engine='openpyxl', mode='a') as writer:
    for date in  cibcdates:
        dates_df = achcreditlonewolf[achcreditlonewolf['Post Date'] == date]
        dates_df.to_excel(writer, sheet_name=f"{date}_achcredit_lonewolf", index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\lonewolf_wire_transaction_list.xlsx', engine='openpyxl', mode='a') as writer:
    for date in  cibcdates:
        dates_df = wirelonewolf[wirelonewolf['Post Date'] == date]
        dates_df.to_excel(writer, sheet_name=f"{date}_wire_lonewolf", index=False)

with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\lonewolf_remote_deposit_transaction_list.xlsx', engine='openpyxl', mode='a') as writer:
    for date in  cibcdates:
        dates_df = remotelonewolf[remotelonewolf['Post Date'] == date]
        dates_df.to_excel(writer, sheet_name=f"{date}_remotedeposit_lonewolf", index=False)


combined_df['key'] = combined_df['Post Date'] + combined_df['BAI Description']
lonewolf_combined_df['key'] =  lonewolf_combined_df['Post Date'] +  lonewolf_combined_df['BAI Description']

combined_df = pd.merge(combined_df, lonewolf_combined_df[['key', 'Lonewolf Amount']], on='key', how = 'left')

combined_df['difference'] = combined_df['CIBC Amount'] - combined_df['Lonewolf Amount']


with pd.ExcelWriter(r'Y:\_Personal Folders\Jason\Escrow\Output\grouped_transactions.xlsx', engine='openpyxl', mode='a') as writer:
    combined_df.to_excel(writer, sheet_name="Transactions by Date", index=False)
    cibcchecks.to_excel(writer, sheet_name="cibc_checks", index=False)
    lonewolfchecks.to_excel(writer, sheet_name="lonewolf_checks", index=False)