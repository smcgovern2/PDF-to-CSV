
# Import Module 
import tabula
import pandas as pd #TODO enforce version number for deprecation
import datetime as dt
import warnings
import numpy as np
import PySimpleGUI as sg

#Configs/Globals

warnings.filterwarnings('ignore', category=FutureWarning)

date_max = dt.date.today()
date_min = dt.date(2022,1,1)



df_columns=['Shipped Date','Customer ID','Zip','Shipment ID','Order Number','PO Number', 'Order Value', 'Shipping Cost']

df_actual=pd.DataFrame(columns=df_columns)

df_filtered=pd.DataFrame(columns=df_columns)

df_raw=pd.DataFrame(columns=df_columns)



poison = False

skip_prompt = True

rmins = [
    0.00,
    12.00,
    15.00,
    25.00,
    35.00,
    45.00,
    55.00,
    75.00,
    100.00,
    125.00,
    150.00,
    175.00,
    200.00,
    225.00,
    250.00,
    275.00,
    300.00,
    325.00,
    350.00,
    375.00,
    400.00,
    425.00,
    450.00,
    475.00,
    500.00,
    525.00,
    550.00,
    575.00,
    600.00,
    625.00,
    650.00,
    675.00,
    700.00,
    725.00,
    750.00,
    775.00,
    800.00,
    825.00,
    850.00,
    875.00,
    900.00,
    925.00,
    950.00,
    975.00,
    1000.00,
    1025.00,
    1050.00,
    1075.00,
    1100.00,
    1125.00,
    1150.00,
    1175.00,
    1200.00,
    1225.00,
    1250.00,
    1275.00,
    1300.00,
    1350.00,
    1400.00,
    1450.00,
    1500.00,
    1550.00,
    1600.00,
    1650.00,
    1700.00,
    1750.00,
    1800.00,
    1850.00,
    1900.00,
    2000.00,
    2250.00,
    2500.00,
    2750.00,
    3000.00,
    3500.00,
    4000.00,
    4500.00,
    5000.00,
    5500.00,
    6000.00,
    6500.00,
    7000.00,
    8000.00,
    np.inf
]

rlabels = [
    "All Orders",
    "$0.00 - $11.99",
    "$12.00 - $14.99",
    "$15.00 - $24.99",
    "$25.00 - $34.99",
    "$35.00 - $44.99",
    "$45.00 - $54.99",
    "$55.00 - $74.99",
    "$75.00 - $99.99",
    "$100.00 - $124.99",
    "$125.00 - $149.99",
    "$150.00 - $174.99",
    "$175.00 - $199.99",
    "$200.00 - $224.99",
    "$225.00 - $249.99",
    "$250.00 - $274.99",
    "$275.00 - $299.99",
    "$300.00 - $324.99",
    "$325.00 - $349.99",
    "$350.00 - $374.99",
    "$375.00 - $399.99",
    "$400.00 - $424.99",
    "$425.00 - $449.99",
    "$450.00 - $474.99",
    "$475.00 - $499.99",
    "$500.00 - $524.99",
    "$525.00 - $549.99",
    "$550.00 - $574.99",
    "$575.00 - $599.99",
    "$600.00 - $624.99",
    "$625.00 - $649.99",
    "$650.00 - $674.99",
    "$675.00 - $699.99",
    "$700.00 - $724.99",
    "$725.00 - $749.99",
    "$750.00 - $774.99",
    "$775.00 - $799.99",
    "$800.00 - $824.99",
    "$825.00 - $849.99",
    "$850.00 - $874.99",
    "$875.00 - $899.99",
    "$900.00 - $924.99",
    "$925.00 - $949.99",
    "$950.00 - $974.99",
    "$975.00 - $999.99",
    "$1,000.00 - $1,024.99",
    "$1,025.00 - $1,024.99",
    "$1,050.00 - $1,074.99",
    "$1,075.00 - $1,099.99",
    "$1,100.00 - $1,124.99",
    "$1,125.00 - $1,149.99",
    "$1,150.00 - $1,174.99",
    "$1,175.00 - $1,199.99",
    "$1,200.00 - $1,224.99",
    "$1,225.00 - $1,249.99",
    "$1,250.00 - $1,274.99",
    "$1,275.00 - $1,299.99",
    "$1,300.00 - $1,349.99",
    "$1,350.00 - $1,399.99",
    "$1,400.00 - $1,449.99",
    "$1,450.00 - $1,499.99",
    "$1,500.00 - $1,549.99",
    "$1,550.00 - $1,599.99",
    "$1,600.00 - $1,649.99",
    "$1,650.00 - $1,699.99",
    "$1,700.00 - $1,749.99",
    "$1,750.00 - $1,799.99",
    "$1,800.00 - $1,849.99",
    "$1,850.00 - $1,899.99",
    "$1,900.00 - $1,999.99",
    "$2,000.00 - $2,249.99",
    "$2,250.00 - $2,499.99",
    "$2,500.00 - $2,749.99",
    "$2,750.00 - $2,999.99",
    "$3,000.00 - $3,499.99",
    "$3,500.00 - $3,999.99",
    "$4,000.00 - $4,499.99",
    "$4,500.00 - $4,999.99",
    "$5,000.00 - $5,499.99",
    "$5,500.00 - $5,999.99",
    "$6,000.00 - $6,499.99",
    "$6,500.00 - $6,999.99",
    "$7,000.00 - $7,999.99",
    "$8,000.00 and up"
]


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
            tester = df.iloc[i].tolist()
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

def group_by_order_value(df):
    df_local = df.copy()
    for col in df_local.columns:
        if df_local[col].dtype == 'object' and '$' in df_local[col].iloc[0]:
            df_local[col] = df_local[col].str.replace('$', '')
            
    df_local['Order Value'] = pd.to_numeric(df_local['Order Value'])
    df_local['Shipping Cost'] = pd.to_numeric(df_local['Shipping Cost'])
    
    grouped_df = pd.DataFrame(columns=['Range', 'Average Shipping Cost'])
    grouped_df['Range'] = pd.Series(rlabels)
    
    i = 1
    avgs=[]
    while i <= len(rmins) - 1:
        avg = 0
        inrange = pd.DataFrame()
        
        try:
            if i < len(rmins):
                inrange = df_local[df_local['Order Value'].between(rmins[i-1],rmins[i]-0.01)]
                avg = inrange['Shipping Cost'].mean()
            if  i == len(rmins):
                inrange = df_local[df_local['Order Value'].ge(rmins[i-1])]
                avg = inrange['Shipping Cost'].mean()
        except KeyError:
            avg = 0
            
        if(pd.isna(avg)):
            avg = 0
        avgs.append(avg)
        i += 1
    
    grouped_df['Average Shipping Cost'] = avgs
        
    print(grouped_df.to_string())
        
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

#Returns Dataframe
def import_pdf(input_path):
    
    df_raw = tabula.read_pdf(input_path, pages='all', stream=True)
    df_processed=pd.DataFrame(columns=df_columns)
    i=1
    for page in df_raw: 
        i +=1

        #Spilts Customer ID and ZIP columns when tabula recognizes as single column due to low whitespace, split on final whitespace char
        try:
            page[['Customer ID', 'Zip']] = page["Customer ID Zip"].apply(lambda x: pd.Series(str(x).rsplit(" ", 1)))
            page = page.drop('Customer ID Zip', axis=1)
        except:
            print("CIDZIP Pass")
            
        #As above for Shipment ID and Order Number
        try:
            page[['Shipment ID', 'Order Number']] = page["Shipment ID Order Number"].apply(lambda x: pd.Series(str(x).rsplit(" ", 1)))
            page = page.drop('Shipment ID Order Number', axis=1)
        except:
            print("SIDON Pass")
            
        #Adresses extractor returning empty column from some input files
        try:
            page = page[page.columns.intersection(df_columns)]
        except:
            print('Blank Col Pass')

        #Scans for entries without shipment ids - These are most often rows where the Customer ID field has overflowed, concats with previous row, then drops any rows without a shipment id
        if list(page.columns.values).sort() == df_columns.sort() and len(list(page.columns.values)) > 0:
            drop_indexes =[]
            j = 0
            while len(page.index) > j:          
                x = page.loc[j,'Shipment ID']
                if not x or pd.isnull(x) or x=='' or x=='nan':
                    try:
                        page.loc[j-1,'Customer ID'] += ' ' + page.loc[j,'Customer ID']
                    except:
                        pass
                    drop_indexes.append(j)                           
                j+=1 
            page = page.drop(index=drop_indexes, axis=0)
            page.reset_index(drop=True, inplace=True)
            df_processed = pd.concat([df_processed, page], axis=0)
        else:
            print('Page ' + str(j) + ' failed to process with columns: ' + str(page.columns.values.tolist()))
    
    df_processed.reset_index(drop=True,inplace=True)
    return df_processed

def build_gui():
    
    NAME_SIZE = 23


    def name(name):
        dots = NAME_SIZE-len(name)-2
        return sg.Text(name + ' ' + '•'*dots, size=(NAME_SIZE,1), justification='r',pad=(0,0), font='Courier 10')

    
    layout_l=[ 
        [name('Order Value Range: ')],
        [sg.Listbox(rlabels, s=(23,25), key='-range_select-')],
        [sg.Button('Filter')]]
    

    layout_r=[
        [name('Matching orders: ')],
        [sg.Table(headings=df_columns, values=[], num_rows=28, key='-table_display-')]]
    
    layout = [  [name('Input File: '),sg.InputText(key='-input_path-'), sg.FileBrowse(), sg.Button('Load')],
                [name('Save CSV as:'), sg.InputText(key='-output_path-'), sg.FileSaveAs(file_types=("CSV","*.csv"), initial_folder="/results")],
                [sg.HSep()],
                [sg.Col(layout_l, p=0),sg.Col(layout_r, p=0)]]

    # Create the Window
    window = sg.Window('PDF Converter', layout, resizable=True)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
            break

        if event == 'Load':          
            try:
                input_path = values['-input_path-']
                df_actual = import_pdf(input_path)
            except Exception as e:
                print('Failed at import')
            
            try:
                df_actual = validate_dataframe(df_actual)
            except:
                print('Failed at Validation')
                
            try:
                table_values = df_actual.values.tolist()
                window['-table_display-'].update(values=table_values)
            except:
                print('failed updating gui')
        
        
        if event == 'Filter':
            try:
                selection_range = window.Element('-range_select-').Widget.curselection()[0]
                print(selection_range)
                if selection_range == 0:
                    df_filtered = df_actual
                else:
                    rmin = rmins[selection_range - 1]
                    rmax = rmins[selection_range]
                    df_filtered = df_actual[df_actual['Order Value'].replace('\$|,', '', regex=True).astype(float).between(rmin,rmax-0.01)]
                print(df_filtered)
                try:
                    table_values = df_filtered.values.tolist()
                    window['-table_display-'].update(values=table_values)
                except:
                    print('failed updating gui')
            except Exception as e:
                pass
            
        if event == '-output_file-':
            outfile_name = values['-output_file-']
            if outfile_name:
                df_actual.to_csv(outfile_name, mode='w')
            
    window.close()



build_gui()

if not skip_prompt:    
    input_path = input("Input path:\n")
    output_path = input("output path:\n")

if skip_prompt:
    input_path = "reports/Test Data.pdf"
    output_path = "results/Test Data.csv"

mr_path = "results/master_record.csv"

# Read PDF File

df_input = tabula.read_pdf(input_path, pages='all', stream=True)

df_processed = pd.DataFrame (columns=df_columns)

try:
    df_master = pd.read_csv(mr_path, dtype=str)
except FileNotFoundError:    
    df_master = pd.DataFrame(df_columns)
    


if poison:
    df_processed = poison_data(df_processed)

df_processed = validate_dataframe(df_processed)

group_by_order_value(df_processed)

df_processed.to_csv(output_path, mode='w')

df_master = pd.concat([df_master,df_processed], axis=0)

df_master = validate_dataframe(df_master)

group_by_order_value (df_master)

df_processed.to_csv(mr_path, mode='w')
    
    
#look into these methods https://xlsxwriter.readthedocs.io/example_pandas_column_formats.html
#   with pd.ExcelWriter('Results/output.xlsx')
#     df_processed.to_excel('Results/output.xlsx')11

