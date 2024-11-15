Sub CopiarCodigo_de_TabDimensao_para_PlanilhaConcessionaria()


'Obras_BI.xlsm

'Transfers codes relating to construction works from the dimension table to the concessionaire's data spreadsheet

'Transfere códigos referentes às obras da tabela dimensão para planilha de dados da concessionária

'Created by Matheus Nunes Reis on 08/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-ConstructionWorksMonitoring/64b7559b9dc94f31ea709efdb32b6a577a650594/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = Workbooks("Obras_BI.xlsm").Sheets("TabDimensao")
    Set wp = Workbooks("acompanhamento_fisico_mensal_concessionaria.xlsx").Sheets("CONCESSIONARIA")

    ' Loop para atribuir os valores
    Dim i As Integer
    Dim j As Integer
    For i = 2 To 129
        ' Obter o valor da célula Ci na planilha "TabDimensao"
        valor = ws.Cells(i, "C").Value

        ' Atribuir o valor para as células B(5+2i) e B(6+2i) na planilha "CONCESSIONARIA"
        For j = 5 + 2 * (i - 1) To 6 + 2 * (i - 1)
            wp.Cells(j, "B").Value = valor
        Next j
    Next i

End Sub
