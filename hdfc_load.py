import xlrd
import psycopg2
book=xlrd.open_workbook("Monthly Portfolios for July 2020 final_0.xls")
database=psycopg2.connect(host="localhost",user="postgres",password="root",database="SIP")
cursor=database.cursor()
truncate="""truncate table tbl_hdfc_pf"""
cursor.execute(truncate)
insert="""INSERT INTO tbl_hdfc_pf (hdfc_fund_name,isin,coupon,instrument_name,industry_rating,quantity,market_fair_value,nav)
values (%s,%s,%s,%s,%s,%s,%s,%s)"""


for s in range(1,100):
 xl_sheet = book.sheet_by_index(s)
 sheet = book.sheet_by_index(s)
 r=8
 hdfc_fund_name=xl_sheet.name
 isin=sheet.cell(r,1).value
 instrument_name=sheet.cell(r,3).value
 #print ('---------------------')
 #print ('Sheet name: %s' % xl_sheet.name)
 #print ('---------------------')
 while isin != 'Grand Total':
  try:
   if len(instrument_name)!=0 and len(isin)!=0:
    #print(isin,instrument_name)
    isin=sheet.cell(r,1).value
    coupon=sheet.cell(r,2).value
    instrument_name=sheet.cell(r,3).value
    industry_rating=sheet.cell(r,4).value
    quantity=sheet.cell(r,5).value
    market_fair_value=sheet.cell(r,6).value
    nav=sheet.cell(r,7).value
    quantity=round(quantity,4)
    market_fair_value=round(market_fair_value,4)
    nav=round(nav,4)
    values=(hdfc_fund_name,isin,coupon,instrument_name,industry_rating,quantity,market_fair_value,nav)
    cursor.execute(insert,values)
    r=r+1
    isin=sheet.cell(r,1).value
    instrument_name=sheet.cell(r,3).value
   else: 
    r=r+1
    isin=sheet.cell(r,1).value
    instrument_name=sheet.cell(r,3).value
  except Exception as e:
   database.rollback()
   print ('Failed Sheet name: %s' % xl_sheet.name)
   print(e)
   print(isin)
   print('-----')
   print('ERROR')
   print('-----')
   break
 database.commit() 
 #print ("I just imported above HDFC data into SIP DB")
 #print ('---------------------')
cursor.close()
database.commit()
database.close()
#columns=str(sheet.ncols)
#rows=str(sheet.nrows)
#print ("I just imported HDFC " + columns + " columns " + "and rows  " + rows + " to SIP DB")
