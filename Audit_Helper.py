# -*- coding: utf-8 -*-
"""
Created on Tue Jul  2 15:24:01 2019

@author: ryoum
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 15:43:50 2019

@author: Richie Youm

#~~~~~~~~~~~~~~~~~~~~~~~~Welcome~~~~~~~~~~~~~~~~~~~~~~~~

Welcome to Club Audit Helper!

This is an internally developed program designed to assist auditing process of student-run clubs.

Name: Club Audit Helper (Temporary)
Version: pre-alpha
Contributor: Richie Youm
Support: seongmin.youm@mail.mcgill.ca

This program works with the transaction list files

FUNCTIONALITY:
    
1. Detect incomplete entries

2. Detect keywords signifying incompliance of internal regulations from event name and descriptions

3. Detect any cash spendings

4. Summarizes amount of undeposited revenues

5. Summarizes profit / losses by event

6. Summarizes deposits recorded

7. Summarizes executive reimbursements by frequency

8. Summarizes club's expense types by frequency

Note:
You need to have all of the transaction lists already downloaded in the same folder for the program to work.
Please keep updating the keywords file, "compliance_keywords.xlsx"

"""

import numpy as np
import pandas as pd

file_names=pd.read_excel('AuditClubList.xlsx') #excel list of CLU to audit
keywords=pd.read_excel('Compliance_Keywords.xlsx') #excel list of keywords
audit_report_file_path='Audit Report/' #where the report would be generated
clu_file_path='FallAudit19-' #file path of the transaction lists and their naming convention before CLU number (ie. WinterAudit20-)

def preprocessing(file_name):

    df=pd.read_excel(file_name)
    df=df.iloc[3:-1,1:12]
    df.columns=df.iloc[0,:].values
    df=df.iloc[4:]
    df=df.reset_index(drop=True)
    
    for i in df.columns:
        df[i]=np.where(df[i]=="Please select",None,df[i])
            
    df['Date']=pd.to_datetime(df['Date'],format='%m/%d/%Y')
    
    for i in ['Event #', 'Amount']:
        df[i]=df[i].astype(float)
        
    df=df[pd.notnull(df['Amount'])]
    df.columns=['Revenue / Expense', 'From / To', 'Event #', 'Event Name', 'Description',
                'Payment Type','Reimbursed Exec', 'Account #', 'Amount', 'Date', 'Ref']
    
    for i in df.columns.drop(['Event #', 'Account #', 'Amount', 'Date', 'Ref']):
        df[i]=df[i].str.lower()
    
    return df

def profitLoss(df):
    
    x=df.groupby(['Revenue / Expense','Event #'],as_index=False).sum()
    
    x_rev=x[x['Revenue / Expense']=='revenue'].reset_index(drop=True)
    x_exp=x[x['Revenue / Expense']=='expense'].reset_index(drop=True)
    x_dep=x[x['Revenue / Expense']=='deposit'].reset_index(drop=True).drop('Revenue / Expense',axis=1)
    x_dep=x_dep.set_index('Event #')
    
    x_tot=x_rev.merge(x_exp,how='outer',on='Event #')
    cols=['Event #','Revenue / Expense_x','Amount_x','Revenue / Expense_y','Amount_y']
    x_tot=x_tot[cols]
    
    total_rev=x_rev['Amount'].sum()
    total_exp=x_exp['Amount'].sum()
    total_profit=total_rev-total_exp
    
    x_tot=x_tot.drop(['Revenue / Expense_x','Revenue / Expense_y'],axis=1)
    x_tot.columns=['Event #', 'Revenues', 'Expenses']
    x_tot['Expenses']=x_tot['Expenses']*-1
    x_tot=x_tot.fillna(0)
    x_tot['Profit / Loss']=[x_tot['Revenues'][i]+x_tot['Expenses'][i] for i in range(x_tot.shape[0])]
    x_tot=x_tot.set_index('Event #').sort_values(by='Event #')
    return x_dep,x_tot,total_profit

def cashSpending(df):
    
    cashSpent=[]
    for i in range(df.shape[0]):
        if df['Payment Type'][i]=='cash' and df['Revenue / Expense'][i]=='expense':
            cashSpent.append(i)
    cashSpent=df.iloc[cashSpent,:]
    return cashSpent

def revExcess(df):
    
    x_dep,x_tot,_=profitLoss(df)
    rev_tot=x_tot['Revenues'].sum()
    dep_tot=x_dep['Amount'].sum()
    excessCash=rev_tot-dep_tot
    excessCash=np.where(excessCash>=0,excessCash,0)
    return excessCash

def reimbursementFreq(df):
    
    top_freq=[]
    
    freq=df['Reimbursed Exec'].value_counts()
    freq=freq.drop('n / a').reset_index()
    freq.columns=['Reimbursed Exec','Frequency']

    for i in range(freq.shape[0]):
        if freq['Frequency'][i] >= freq['Frequency'].quantile(.75):
            top_freq.append(freq['Reimbursed Exec'][i])
    freq=freq.set_index('Reimbursed Exec')
    return freq, top_freq
    
def mainExpFreq(df):
    
    expFreq=df['Account #'].value_counts()
    expFreq=expFreq.reset_index()
    expFreq.columns=['Expense Type','Frequency']
    expFreq=expFreq.set_index('Expense Type')
    return expFreq

def keywordDetector(df):
    
    cleaned_df=df.fillna('No Input')
    flagged=cleaned_df[cleaned_df['Description'].str.contains('|'.join(keywords['Keywords']))]
    flagged=flagged.append(cleaned_df[cleaned_df['Event Name'].str.contains('|'.join(keywords['Keywords']))])
    flagged=flagged.drop_duplicates(subset=flagged.columns)
    flagged=flagged.sort_values(by="Date")
    return flagged

def incompleteAlert(df):
    
    return df[df.isnull().any(axis=1)]

def get_col_widths(df):
    
    index_max=max([len(str(i)) for i in df.index.values] + [len(str(df.index.name))])
    return [index_max]+[max([len(str(i)) for i in df[col].values] +[len(col)]) for col in df.columns]

def spacer():
    for i in range(3):
        print("")

for file_name in file_names['CLU Number'].replace(' ',''):
    file_name=clu_file_path+file_name+'.xlsx'
    
    #calling necessary functions
    df=preprocessing(file_name)
    x_dep,x_tot,total_profit=profitLoss(df)
    freq,top_freq=reimbursementFreq(df)
    expFreq=mainExpFreq(df)
    incomplete_df=incompleteAlert(df)
    flagged=keywordDetector(df)
    cashSpent=cashSpending(df)
    excessCash=revExcess(df)
    
    #for console outputs
    spacer()
    
    print("NA VALUES?:")
    print("Following entries contain NA values: (if empty, ignore)")
    print("Look at excel rows of: ",np.array(incomplete_df.index+9))
    print(incomplete_df)
    
    spacer()
    
    print("PROFIT / LOSS REPORT: \n")
    print(x_tot)
    print("This club's total profit is ${}".format(total_profit))
    
    spacer()
    
    print("REIMBURSEMENT FREQUENCY: \n")
    print(freq,"\n")
    print("Most reimbursed exec: ",top_freq)
    
    spacer()
    
    print('KEYWORD DETECTOR: \n')
    print(flagged)
    
    spacer()
    
    print('CASH SPENDING DETECTOR: \n')
    print(cashSpent)
    
    spacer()
    
    print('AMOUNT DEPOSITED: \n')
    print(x_dep, "\n")
    print('REVENUE UNDEPOSITED: \n')
    print('$',excessCash)
    
    spacer()
    
    #names of files to be saved
    excel_name=file_name.split('/')[-1].split('.')[0]
    writer=pd.ExcelWriter(audit_report_file_path+str(excel_name)+'_report.xlsx', engine='xlsxwriter')
           
    #converting dataframes to excel
    df.to_excel(writer,sheet_name='Transaction List')
    incomplete_df.to_excel(writer,sheet_name='Compliance Warning', startrow=3)
    flagged.to_excel(writer,sheet_name='Compliance Warning',startrow=(incomplete_df.shape[0]+7))
    cashSpent.to_excel(writer,sheet_name='Compliance Warning',startrow=(incomplete_df.shape[0]+flagged.shape[0]+11))
    x_tot.to_excel(writer,sheet_name='General Report',startrow=3)
    x_dep.to_excel(writer,sheet_name='General Report',startrow=(x_tot.shape[0]+8))
    freq.to_excel(writer,sheet_name='General Report',startrow=3,startcol=5)
    expFreq.to_excel(writer,sheet_name='General Report',startrow=3,startcol=9)
    
    #declaring workbook
    workbook=writer.book
    
    #declaring worksheets
    worksheet1=writer.sheets['Transaction List']
    worksheet2=writer.sheets['Compliance Warning']
    worksheet3=writer.sheets['General Report']
    
    #formats
    bold_format=workbook.add_format()
    bold_colored_format=workbook.add_format()
    bold_big_format=workbook.add_format()
    currency_format=workbook.add_format({'num_format':'	#,##0.00'})
                                             
    bold_format.set_bold()
    bold_colored_format.set_bold()
    bold_colored_format.set_bg_color('orange')
    bold_big_format.set_bold()
    bold_big_format.set_font_size(15)
    
    #worksheet 2
    worksheet2.write('B2','Incomplete Entries Detected',bold_big_format) #If there are empty values, fill it up in the ORIGINAL document (if it's clear), and run the file again. Program may not work with empty values in it. You may also deduct points for non-compliance, if needed.
    worksheet2.write('B{}'.format(incomplete_df.shape[0]+6),'Keywords Detected',bold_big_format) #Please review the following entries. Duplicates have been removed (check orignial if needed)

    worksheet2.write('B{}'.format(incomplete_df.shape[0]+flagged.shape[0]+10),'Cash Spending Detected',bold_big_format)
    worksheet2.write('B{}'.format(incomplete_df.shape[0]+flagged.shape[0]+15),'Revenue Undeposited: $ {}'.format(excessCash) ,bold_big_format)
    
    #worksheet 3
    worksheet3.write('B2','Profits and Losss', bold_big_format)
    worksheet3.write('F2','Reimbursement Freq', bold_big_format)
    worksheet3.write('J2','Expense Type Freq', bold_big_format)
    worksheet3.write(x_tot.shape[0]+6,1,"Deposits", bold_big_format)
    
    worksheet3.write(x_tot.shape[0]+4,0,"Total: ")
    worksheet3.write(x_tot.shape[0]+4,1,np.sum(x_tot['Revenues']))
    worksheet3.write(x_tot.shape[0]+4,2,np.sum(x_tot['Expenses']))
    worksheet3.write(x_tot.shape[0]+4,3,np.sum(x_tot['Profit / Loss']))
    
    worksheet3.write(x_tot.shape[0]+x_dep.shape[0]+9,0,"Total: ")
    worksheet3.write(x_tot.shape[0]+x_dep.shape[0]+9,1,np.sum(x_dep['Amount']))
    
    worksheet3.conditional_format('D5:D'+str(x_tot.shape[0]+4),{'type':'3_color_scale'})
    try:
        worksheet3.conditional_format('F5:F'+str(freq.shape[0]+4),{'type':'formula','criteria':"G5:G{} >= {}".format(int(freq.shape[0]+1),int(freq.quantile(.75))),'format':bold_colored_format})
    except ValueError:
        pass
    worksheet3.conditional_format('J5:J'+str(expFreq.shape[0]+4),{'type':'formula','criteria':"K5:K{} >= {}".format(int(expFreq.shape[0]+1),int(expFreq.quantile(.75))),'format':bold_colored_format})
    worksheet3.set_column('B:D',None,currency_format)
    
    #spaces the columns out
    for i,width in enumerate(get_col_widths(df)):
        worksheet1.set_column(i,i,width)
        worksheet2.set_column(i,i,width)
    
    writer.save()