import pandas as pd

# Get execution percentage of construction works from selected columns
# Columns selected like excel labels and converted later to worksheet titles
# Save new worksheet in excel format .xlsx to be applied in Business Inteligence platform

sheet_index_number = 0
startLine = 3 # Firt excel file line with titles = 5 -> StartLine = 3
xi = 'A' # start of general information interval
xf = 'L' # end of general information interval
yi = 'AO' # start of sum interval
yf = 'AU' # end of sum interval
zi = 'BD' # start of status informations
zf = 'BG' # end of status informations


file_path = r'C:\Users\matheus.reis\Desktop\acompanhamento_fisico_mensal_ecosul.xlsx'
output_file_path = r'C:\Users\matheus.reis\Desktop\ecosul.xlsx'
Concession = 'ECOSUL'

# file_path = r'C:\Users\matheus.reis\Desktop\acompanhamento_fisico_mensal_riosp.xlsx'
# output_file_path = r'C:\Users\matheus.reis\Desktop\RIOSP.xlsx'
# Concession = 'RioSP'


# # # # # # # # # # # # # # # # # # # 

# Get the sheet name by index
sheet_name = pd.ExcelFile(file_path, engine='openpyxl').sheet_names[sheet_index_number]
# Cretae dataframe from specified worksheet
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')


#------Label dataframe columns like excel (A,..., Z, AA, AB,...)----------
excel_labels = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ', 'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ', 'EA', 'EB', 'EC', 'ED', 'EE', 'EF', 'EG', 'EH', 'EI', 'EJ', 'EK', 'EL', 'EM', 'EN', 'EO', 'EP', 'EQ', 'ER', 'ES', 'ET', 'EU', 'EV', 'EW', 'EX', 'EY', 'EZ', 'FA', 'FB', 'FC', 'FD', 'FE', 'FF', 'FG', 'FH', 'FI', 'FJ', 'FK', 'FL', 'FM', 'FN', 'FO', 'FP', 'FQ', 'FR', 'FS', 'FT', 'FU', 'FV', 'FW', 'FX', 'FY', 'FZ', 'GA', 'GB', 'GC', 'GD', 'GE', 'GF', 'GG', 'GH', 'GI', 'GJ', 'GK', 'GL', 'GM', 'GN', 'GO', 'GP', 'GQ', 'GR']

# Generate Excel-like column labels
num_columns = df.shape[1]
excel_labels = excel_labels[:num_columns]

df.columns = excel_labels
#-------------------------------------------------------------------------

# copy and duplicate data
def copy_lines_duplicating(df):
    xf_index_less1 = df.columns.get_loc(xf) - 1
    columns_to_duplicate = list(df.loc[:,xi:df.columns[xf_index_less1]].columns) + list(df.loc[:,zi:zf].columns)
    for i in range(startLine, len(df), 2):
        if i + 1 < len(df):
            df.loc[i + 1, columns_to_duplicate] = df.loc[i, columns_to_duplicate]
    return df


modified_df = copy_lines_duplicating(df) # Duplicate data in pairs (Fill blanks in excel merged cells)

# Columns yi to yf considered in the sum
columns_to_sum = list(modified_df.loc[:, yi:yf].columns)


# New Dataframe 'numeric_df' to retrieve sum of the specified columns yi to yf
numeric_df = modified_df.loc[:,columns_to_sum].apply(pd.to_numeric, errors='coerce') # If data is no number convert in NaN, but by data type
                                                                                        # It shpould not be converted to NaN properly
numeric_df = numeric_df.applymap(lambda x: np.nan if x < 0 or x > 1 else x) # In case conversion to NaN fails, convert to NaN if converted data
                                                                            # is negative or bigger than 1
numeric_df = numeric_df.fillna(0)
numeric_df['Total'] = numeric_df.sum(axis=1)


# Adds column of total sums
modified_df['Total Acumulado'] = numeric_df['Total']

# Adjust concluded woks to 100% in column 'Total Acumulado'
modified_df.loc[(modified_df[zi] == 'CONCLUÍDA') | (modified_df[zi] == 'CONCLUÍDO') , 'Total Acumulado'] = 1   # 100% executada

modified_df.loc[modified_df[zi] == 'NÃO INICIADA', 'Total Acumulado'] = 0   # 0% executada

#Apply percentage format in column 'Total Acumulado'
#modified_df['Total Acumulado'] = modified_df['Total Acumulado'].apply(lambda x: f"{x:.2%}")

# Copy titles from xf and yi columns to the next line (next line is used to create to define final columns' labels)
modified_df.loc[startLine + 1, [xf,yi]] = modified_df.loc[startLine, [xf,yi]]

columns_to_show = ( list(modified_df.loc[:, xi:xf].columns) + list(modified_df.loc[:, yi:yf].columns) + 
                    list(modified_df.loc[:, zi:zf].columns) ) + ['Total Acumulado']

modified_df = modified_df.loc[:,columns_to_show]

# Rename all columns' labels
labels = list(modified_df.iloc[startLine+1, :len(modified_df.columns)-1]) + ['Total acumulado']
modified_df.columns = labels

modified_df = modified_df.loc[startLine+2:, :]

modified_df['Concessionária'] = Concession

# Get the last column name
last_column = modified_df.columns[-1]

# Reorder columns to move the last column to the first position
modified_df = modified_df[[last_column] + list(modified_df.columns[:-1])]


#print(modified_df)

modified_df.to_excel(output_file_path, index=False) # index=False means not to print line indexes

