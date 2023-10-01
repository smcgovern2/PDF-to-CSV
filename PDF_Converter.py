
# Import Module 
import tabula
import pandas as pd #TODO enforce version number for deprecation
import datetime as dt

#Configs

date_max = dt.date.today()
date_min = dt.date(2022,1,1)

df_columns=['Shipped Date','Customer ID','Zip','Shipment ID','Order Number','PO Number', 'Order Value', 'Shipping Cost']

poison = False

skip_prompt = False
#Functions
#Also converts values to string
def clean_trailing_zeroes(df, i, columns):
     for col in columns:
        try:
            df.at[i,col] = str(df.loc[i,col]).replace(".0","")
        except(ValueError):
            print('non numeric ' + col + ' at index ' + i )

#Validates Currency, returns true when invalid
def validate_currency_range(val, max, min):
    
    if val.startswith("$"):
        val = val.replace("$","").strip()
    if not ((val.isnumeric() or val.replace(".", "").isnumeric()) and min <= float(val) <= max):
        return True
    else:
        return False

#validates and returns df        
def validate_dataframe (df):
    #Removes duplicate shipment IDs 
    try: df = df.drop_duplicates(subset=['Shipment ID'])
                                                    
    except:
        print('Duplicate test failure')

    df.reset_index(drop=True,inplace=True)
    
    drop_indexes = []
    
    for i in df.index:             
        
        #Standardizes values that are sometimes extracted as numpy.float, sets as str
        clean_trailing_zeroes(df, i, ['Zip','Order Number', 'PO Number', 'Shipment ID'])     
            
            
        #Allows current row until flipped      
        failed = False
        
        #Date-range Check
        try:
            date_df = dt.datetime.strptime((df.loc[i,'Shipped Date']),'%m/%d/%Y').date()
            if not (date_min <= date_df <= date_max):
                failed = True
                print("Invalid value in shipped date at index " + str(i))
        except ValueError:
            failed = True
            print("Invalid value in shipped date at index " + str(i))
            
            
        #Zip Check, verifies numeric and length     
        if not (df.loc[i,'Zip'].isnumeric() and 5 <= len(df.loc[i,'Zip']) <= 9):
            failed = True
            print("Invalid value in ZIP at index " + str(i))
            
        #Order Value Check
        order_value = df.loc[i,'Order Value']
        if not failed:
            failed = validate_currency_range(order_value,10000,0)
        
        #Shipping cost check
        shipping_cost = df.loc[i,'Shipping Cost']
        if not failed:
            failed = validate_currency_range(shipping_cost,1000,0)
            
        if failed == True:
            drop_indexes.append(i)
            
    print(drop_indexes)

    df = df.drop(index=drop_indexes, axis=0)        
    
    df.reset_index(drop=True, inplace=True)
    
    return df

#Alter dataframe for validation testing        
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

if not skip_prompt:    
    input_path = input("Input path:/n")
    output_path = input("output path:/n")

if skip_prompt:
    input_path = "reports/test data.pdf"
    output_path = "results/test data.csv"

mr_path = "results/master_record.csv"

# Read PDF File

df_input = tabula.read_pdf(input_path, pages='all', stream=True)

df_processed = pd.DataFrame (columns=df_columns)

try:
    df_master = pd.read_csv(mr_path, dtype=str)
except FileNotFoundError:    
    df_master = pd.DataFrame(df_columns)

for page in df_input: 
    
    #Spilts Customer ID and ZIP columns when tabula recognizes as single column due to low whitespace, split on final whitespace char
    try:
        page[['Customer ID', 'Zip']] = page["Customer ID Zip"].apply(lambda x: pd.Series(str(x).rsplit(" ", 1)))
        page = page.drop('Customer ID Zip', axis=1)
    except:
        print("CIDZIP Pass")
    #Adresses extractor returning empty column from some input files
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

if poison:
    df_processed = poison_data(df_processed)

df_processed = validate_dataframe(df_processed)

df_processed.to_csv(output_path, mode='w')

df_master = pd.concat([df_master,df_processed], axis=0)

df_master = validate_dataframe(df_master)

df_processed.to_csv(mr_path, mode='w')
    
    
#look into these methods https://xlsxwriter.readthedocs.io/example_pandas_column_formats.html
#   with pd.ExcelWriter('Results/output.xlsx')
#     df_processed.to_excel('Results/output.xlsx')


