
# Import Module 
import tabula
import pandas as pd
import datetime as dt

#configs

date_max = dt.date.today()
date_min = dt.date(2022,1,1)

#functions

def clean_trailing_zeroes(df, i, columns):
    for col in columns:
        try:
            df.at[i,col] = str(df.loc[i,col]).replace(".0","")
        except(ValueError):
            print('non numeric ' + col + ' at index ' + i )
            
#TODO: Validation to functions to reuse on master record

def poison_data(df):
    df.at[10,'Shipped Date'] = '01/01/2000'
    df.at[20,'Shipped Date'] = '01/01/2030'
    df.at[55, 'Shipped Date'] = 'AAA'
   
    df.at[30,'Zip'] = 333
    df.at[40,'Zip'] = 444455556666
    df.at[50,'Zip'] = 'ABC123'
    
    df.at[60,'Order Value'] = '$2000.00'
    df.at[65,'Order Value'] = '-$20.00'
    df.at[15,'Order Value'] = 'Fifty Dollars'
   
    df.at[25,'Shipping Cost'] = '$20000.00'
    df.at[35,'Shipping Cost'] = '-$20.00'
    df.at[45,'Shipping Cost'] = 'Fifty Dollars'
    
    df.loc[0:5,'Shipment ID'] = '94612361042629129568'
    
    return df

    

# Read PDF File

mr_path = "Results/master_record.csv"

df = tabula.read_pdf("Reports/2023-09-05 3 MONTHS.pdf", pages='all', stream=True)

df_processed = pd.DataFrame ()

for page in df: 
    
    #Spilts Customer ID and ZIP columns when tabula recognizes as single column due to low whitespace, split on final whitespace char, whitespace in ZIP will break this
    try:
        page[['Customer ID', 'Zip']] = page["Customer ID Zip"].apply(lambda x: pd.Series(str(x).rsplit(" ", 1)))
        page = page.drop('Customer ID Zip', axis=1)
    except:
        print("CIDZIP Pass")
    #Extractor returns empty column from some input files
    try:
        page = page.drop('Unnamed: 0', axis=1)
    except:
        print('Blank Col Pass')

    #Scans for entries without shipped dates - These are rows where the Customer ID field has overflowed, concats with previous row
    try:
        i = 0
        while len(page.index) > i:          
            print(page.loc[i,'Shipped Date'])
            x = page.loc[i,'Shipped Date']
            if not page.loc[i,'Shipped Date'] or pd.isnull(page.loc[i,'Shipped Date']) or (page.loc[i,'Shipped Date']==''):       
                page.loc[i-1,'Customer ID'] += ' ' + page.loc[i,'Customer ID']
            i+=1  
    except:
        print('Shipped Date Unproccessed')
    #clears all rows used in previous step
    page = page[page['Shipment ID'].notnull()]  
    
    df_processed = pd.concat([df_processed, page], axis=0)
    
    
df_processed.reset_index(drop=True,inplace=True)   

df_processed = poison_data(df_processed)


#Removes duplicate shipment IDs and reindexes --- Needs Fix
try: df_processed = df_processed.drop_duplicates(subset=['Shipment ID'])
                                                 
except:
    print('Duplicate test failure')

df_processed.reset_index(drop=True,inplace=True)

 #Validation 

drop_indexes = []
    
for i in df_processed.index:             
    
    #Standardizes values that are sometimes extracted as numpy.float, sets as str
    clean_trailing_zeroes(df_processed, i, ['Zip','Order Number', 'PO Number', 'Shipment ID'])     
          
          
    #Allows current row until flipped      
    failed = False
    
    #Date-range Check
    try:
        date_df = dt.datetime.strptime((df_processed.loc[i,'Shipped Date']),'%m/%d/%Y').date()
        if not (date_min <= date_df <= date_max):
            failed = True
            print("Invalid value in shipped date at index " + str(i))
    except ValueError:
        failed = True
        print("Invalid value in shipped date at index " + str(i))
        
        
    #Zip Check, verifies numeric and length     
    if not (df_processed.loc[i,'Zip'].isnumeric() and 5 <= len(df_processed.loc[i,'Zip']) <= 9):
        failed = True
        print("Invalid value in ZIP at index " + str(i))
        
    #Order Value Check
    order_value = df_processed.loc[i,'Order Value']
    if order_value.startswith("$"):
        order_value = order_value.replace("$","").strip()
    if not ((order_value.isnumeric() or order_value.replace(".", "").isnumeric()) and 0 <= float(order_value) <=10000):
        failed = True
        print("Invalid value in order value at index " + str(i))
        
    #Shipping cost check
    shipping_cost = df_processed.loc[i,'Shipping Cost']
    if shipping_cost.startswith("$"):
        shipping_cost = shipping_cost.replace("$","").strip()
    if not ((shipping_cost.isnumeric() or shipping_cost.replace(".", "").isnumeric()) and 0 <= float(shipping_cost) <=1000):
        failed = True
        print("Invalid value in shipping cost at index " + str(i))
    
    if failed == True:
        drop_indexes.append(i)
        
print(drop_indexes)

df_processed.drop(index=drop_indexes,inplace=True, axis=0)        

df_processed = df_processed.reset_index(drop=True)

df_processed.to_csv(mr_path, mode='w')


#look into these methods https://xlsxwriter.readthedocs.io/example_pandas_column_formats.html
#   with pd.ExcelWriter('Results/output.xlsx')
#     df_processed.to_excel('Results/output.xlsx')

