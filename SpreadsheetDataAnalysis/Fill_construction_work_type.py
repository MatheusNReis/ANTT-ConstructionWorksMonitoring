# Fill the construction work type column according to detailed description and observation fields in spreadsheet

#import pdb # Import the pdb module
#pdb.set_trace()

import win32com.client # from pywin32 library

# Imputs
WorksheetName = "Concession Name"
ColumnDescription = 3 # Column number with description of the construction work
ColumnConstructionType = 4 # column number with Kind of construction work
Initial_row = 7 # row that start construction description

# Connect to the Excel application
excel = win32com.client.Dispatch("Excel.Application")

# Get the active workbook
workbook = excel.ActiveWorkbook
worksheet = workbook.worksheets(WorksheetName)



# Function to get the last filled row in a specific column
def get_last_filled_row(worksheet, column):
    xlUp = -4162  # Numeric value for xlUp
    last_filled_row = worksheet.Cells(worksheet.Rows.Count, column).End(xlUp).Row
    return last_filled_row

# Get the last filled row in column
last_row = get_last_filled_row(worksheet, ColumnDescription) + 1 # +1 to adjust the merge cell reference 



# Function to get the top-left cell value if the cell is merged
def get_top_left_cell_value(worksheet, row, col):
    cell = worksheet.Cells(row, col)
    if cell.MergeCells:
        return worksheet.Cells(cell.MergeArea.Row, cell.MergeArea.Column).Value
    else:
        return cell.Value


# Verify the kind of construction work described
for i in range(Initial_row, last_row):
    cell_value = get_top_left_cell_value(worksheet, i, ColumnDescription).lower()


    if cell_value and "reforma" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Reforma"
    
    if cell_value and "recuperação" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Recuperação"
        
    if cell_value and "adequação" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Adequação"

    if cell_value and "restauração" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Restauração"

    if cell_value and "retificação" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Retificação"
    
    if cell_value and ("travessia" in cell_value) and ("urbana" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Travessia Urbana"
    
    if cell_value and "rua" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Rua"

    if cell_value and "alça" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Alça"
    
    if cell_value and "viaduto" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Viaduto"

    if cell_value and "ponte" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Ponte"

    if cell_value and "agulha" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Agulha"
    
    if (cell_value and "faixa" in cell_value) and ("adicion" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Faixa Adicional"

    if cell_value and "oae" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "OAE"
    
    if cell_value and ("obra" in cell_value) and ("arte" in cell_value) and ("especia" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "OAE"
    
    if cell_value and "margin" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Via Marginal"

    if (cell_value and "terceira" in cell_value) and ("faixa" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Terceira Faixa"
    
    if cell_value and "duplicaç" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Duplicação"

    if cell_value and "iluminação" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Iluminação"
    
    if cell_value and "entroncamento" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Entroncamento"
    
    if cell_value and "melhor" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Melhorias"

    if cell_value and "acesso" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Acesso"

    if cell_value and ("melhoria" in cell_value) and ("acesso" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Melhoria"

    if cell_value and "passarela" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Passarela"
    
    if cell_value and "contorno" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Contorno"

    if cell_value and "retorno" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Retorno"
    
    # Keep "Intersec" before "Diamante" and "Trombeta" and "Parclo"
    if cell_value and "interse" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Intersecção"

    if cell_value and ("melhoria" in cell_value) and ("intersec" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Melhoria"

    if (cell_value and "interse" in cell_value) and (("nível" in cell_value) or("nivel" in cell_value)):
        worksheet.Cells(i, ColumnConstructionType).Value = "Intersecção em nível"

    if (cell_value and "interse" in cell_value) and (("desnível" in cell_value) or("desnivel" in cell_value)):
        worksheet.Cells(i, ColumnConstructionType).Value = "Intersecção em desnível"
    
    # Keep "trevo" before "Diamante" and "Trombeta"
    if cell_value and "trevo" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Trevo"

    if (cell_value and "trevo" in cell_value) and (("nível" in cell_value) or("nivel" in cell_value)):
        worksheet.Cells(i, ColumnConstructionType).Value = "Trevo em nível"

    if (cell_value and "trevo" in cell_value) and (("desnível" in cell_value) or("desnivel" in cell_value)):
        worksheet.Cells(i, ColumnConstructionType).Value = "Trevo em desnível"

    if cell_value and "diamante" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Diamante"

    if cell_value and "trombeta" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Trombeta"

    if cell_value and "parclo" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Parclo"
    
    if cell_value and "prf" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "PRF"
    
    if cell_value and "ppd" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "PPD"

    if cell_value and ("ponto" in cell_value) and ("parada" in cell_value) and ("descanso" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "PPD"

    if cell_value and "posto de fiscalização" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Posto de Fiscalização"

    if cell_value and "uop" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "UOP"

    if cell_value and "uop" in cell_value and "Delegacia" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "UOP + Delegacia"

    if (cell_value and "rotatória" in cell_value) or (cell_value and"Rotatoria" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Rotatória"

    if cell_value and "contorno" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Contorno"

    if (cell_value and "reversível" in cell_value) or (cell_value and "reversivel" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Faixa Reversível"

    if cell_value and "ppv" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "PPV"

    if cell_value and "ppv fixo" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "PPV fixo"

    if cell_value and ("pesage" in cell_value) and ("veicular" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "PPV"
    
    if cell_value and "posto de pesagem veicular fixo" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "PPV fixo"
    
    if cell_value and ("posto" in cell_value) and ("pesage" in cell_value) and (("veícul" in cell_value) or ("veicul" in cell_value)) and ("fixo" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "PPV"

    if cell_value and "margin" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Via Marginal"
    
    if cell_value and "defensa" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Defensa"

    if cell_value and ("barreira" in cell_value) and ("concreto" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Barreira de concreto"

    if cell_value and ("barreira" in cell_value) and ("ruído" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Barreira de ruído"

    if cell_value and ("ponto" in cell_value) and ("ônibus" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Ponto de ônibus"
    
    if cell_value and "balança" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Balança"

    if cell_value and "balança fixa" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Balança fixa"

    if cell_value and "balanças fixas" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Balança fixa"

    if cell_value and "cftv" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "CFTV"

    if cell_value and ("circuito" in cell_value) and ("fechado") and ("tv" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "CFTV"

    if cell_value and "edificação" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Edificação"
    
    if cell_value and "bso" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "BSO"

    if cell_value and ("base" in cell_value) and ("operaciona" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "BSO"

    if cell_value and "cco" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "CCO"

    if cell_value and "passagem superior" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Passagem em desnível"

    if cell_value and "passagem inferior" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Passagem em desnível"

    if cell_value and ("passage" in cell_value) and (("desnível" in cell_value) or ("desnivel" in cell_value)):
        worksheet.Cells(i, ColumnConstructionType).Value = "Passagem em desnível"
    praça
    if cell_value and ("sist" in cell_value) and ("controle" in cell_value) and ("velocidade" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Sistema de controle de velocidade"
        
    if cell_value and ("control" in cell_value) and ("velocidade" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Controlador de velocidade"

    if cell_value and ("sist" in cell_value) and ("detec" in cell_value) and ("altura" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Sistema de detecção de altura"

    if cell_value and ("sist" in cell_value) and ("detec" in cell_value) and ("incidente" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "DAI"

    if cell_value and ("sist" in cell_value) and ("meteorol" in cell_value) and ("incidente" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Sistema monitoração meteorológica"

    if cell_value and "terrapleno" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Terrapleno"

    if cell_value and ("sist" in cell_value) and ("elétrico" in cell_value) and ("iluminação" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Sistema elétrico e de iluminação"

    if cell_value and ("elemento" in cell_value) and ("protec" in cell_value) and ("segurança" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "EPS"

    if cell_value and "acostamento" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Acostamento"
    
    if cell_value and "drenage" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "Drenagem"
    
    if cell_value and ("obra" in cell_value) and ("arte" in cell_value) and ("corrente" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "OAC"

    if cell_value and ("drenage" in cell_value) and ("obra" in cell_value) and ("arte" in cell_value) and ("corrente" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Drenagem e OAC"

    if cell_value and ("praça" in cell_value) and ("pedágio" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Praça de pedágio"
    
    if cell_value and "sau" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "SAU"

    if cell_value and ("sau" in cell_value) and ("bso" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "BSO + SAU"

    if cell_value and ("base" in cell_value) and ("bpr" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Base BPRv"

    if cell_value and "sat" in cell_value:
        worksheet.Cells(i, ColumnConstructionType).Value = "SAT"

    if cell_value and ("detecção" in cell_value) and ("sensoriamento" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "SAT"

    if cell_value and ("sist" in cell_value) and ("comunicação" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Sistema de comunicação"
    
    # if cell_value and ("sist" in cell_value) and ("comunicação" in cell_value) and ("wireless" in cell_value):
    #     worksheet.Cells(i, ColumnConstructionType).Value = "Sistema de comunicação wireless"

    if cell_value and ("fibra" in cell_value) and (("óptica" in cell_value) or ("ótica" in cell_value)):
        worksheet.Cells(i, ColumnConstructionType).Value = "Fibra óptica"

    if cell_value and ("canteiro" in cell_value) and ("central" in cell_value) and ("faixa" in cell_value) and ("domínio" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Canteiro Central e faixa de domínio"

    if cell_value and ("pain" in cell_value) and ("mensage" in cell_value) and ("variáve" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "PMV"

    if cell_value and ("faixa" in cell_value) and ("aceleração" in cell_value) and ("desaceleração" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Faixa de aceleração e desaceleração"

    if cell_value and ("escritório" in cell_value) and ("antt" in cell_value):
        worksheet.Cells(i, ColumnConstructionType).Value = "Escritório ANTT"


print("Process completed successfully.")
