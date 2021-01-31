import pandas as pd
import openpyxl
from openpyxl import load_workbook
import datetime as dt
import os
import xlsxwriter

comboinput= input('Enter current COMBO Query: ')
covidfileinput = input('Enter COVID file to be analyzed: ')
covidyestfileinput = input('Enter the most recent COVID file to be analyzed against: ')

dfcombo = pd.read_excel(filepath + comboinput, index_col='Rule', skiprows=1)
dfcombo.rename(mapper= {'HOS': 'Hospital',
                       'FIRE': 'Fire House',
                       'SCHL': 'School',
                       'POL': 'Police',
                       'Air': 'Airport',
                       'Off': 'Office',
                       'HOME': 'Home'}, inplace=True)

dfcombo.reset_index(inplace=True)
dfcombo['Req Fund']= dfcombo['Req Fund'].astype('str')
dfcombo['Dept']= dfcombo['Dept'].astype('str')
dfcombo['Program'] = dfcombo['Program'].astype('str')
dfcombo['COMBO CONCAT'] = dfcombo['Rule'] + dfcombo['Req Fund'] + dfcombo['Oper Unit'] + dfcombo['Dept'] + dfcombo['Project']
combolist = list(dfcombo['COMBO CONCAT'])

dfcovid = pd.read_excel(filepath + covidfileinput,skiprows=1)
dfyestcovid = pd.read_excel(filepath + covidyestfileinput, skiprows=1)

dfyestcovid['Req ID'] = dfyestcovid['Req ID'].apply(lambda x: '{0:0>10}'.format(x)) #adding leading zeros
dfcovid['Req ID'] = dfcovid['Req ID'].apply(lambda x: '{0:0>10}'.format(x)) #adding leading zeros

dfcovid['PO Fund'] = dfcovid['PO Fund'].astype('str')

dfcovid['PO Dept'] = dfcovid['PO Dept'].astype('str')
dfcovid['COVID CONCAT'] = dfcovid['PO GL Unit'] + dfcovid['PO Fund'] + dfcovid['PO Oper Unit'] + dfcovid['PO Dept'] + dfcovid['PO Project']

dfcovid['COMBO CONCAT'] = dfcovid['COVID CONCAT'].isin(combolist)

comdict = dfyestcovid.set_index('Req ID')['Comments'].to_dict()

dfcovid['Comments']= dfcovid['Req ID'].map(comdict)

dfcovid['Supplier ID'] = dfcovid['Supplier ID'].apply(lambda x: '{0:0>10}'.format(x))

dfcovid['PO Date'] = pd.to_datetime(dfcovid['PO Date'])
dfcovid['Req Date'] = pd.to_datetime(dfcovid['Req Date'])
dfcovid['Entered'] = pd.to_datetime(dfcovid['Entered'])

dfcovid['PO Date'] = dfcovid['PO Date'].dt.strftime('%m/%d/%Y')
dfcovid['Req Date'] = dfcovid['Req Date'].dt.strftime('%m/%d/%Y')
dfcovid['Entered'] = dfcovid['Entered'].dt.strftime('%m/%d/%Y')

dfcovid['Req Merch Amt'] = dfcovid['Req Merch Amt'].apply(lambda x: "${:,.2f}".format((x)))
dfcovid['Merchandise Amt'] = dfcovid['Merchandise Amt'].apply(lambda x: "${:,.2f}".format((x)))

writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
dfcovid.to_excel(writer, sheet_name='COVID REQ NAMES', index=False, startrow=0, startcol=0)
dfcombo.to_excel(writer, sheet_name='COMBOs', index=False, startrow=0, startcol=0)
writer.save()
