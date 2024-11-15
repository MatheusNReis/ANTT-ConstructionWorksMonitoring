Sub extrairdataNoFormatoBRA()

'Obras_BI.xlsm

'Converts date to Brazilian format in order to process the reorganization of construction works data

'Converte data para formato brasileiro a fim de processar a reorganização de dados das obras

'Created by Matheus Nunes Reis on 09/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-ConstructionWorksMonitoring/64b7559b9dc94f31ea709efdb32b6a577a650594/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


' Extrair a data da célula A1
Dim data As Date
data = Workbooks("acompanhamento_fisico_mensal_concessionaria.xlsx").Sheets("CONCESSIONARIA").Range("AB6").Value

' Extrair o dia, mês e ano
Dim Dia As Integer, Mes As Integer, Ano As Integer

Dia = Day(data)
Mes = Month(data)
Ano = Year(data)
DataFormato = DateSerial(Ano, Mes, Dia)

' Escreve data na célula "G2" da planilha
Workbooks("Obras_BI.xlsm").Sheets("Planejado").Range("G2").Value = DataFormatoBR

' Imprimir o dia, mês e ano
MsgBox "Dia: " & Dia & vbNewLine & _
"Mês: " & Mes & vbNewLine & _
"Ano: " & Ano & vbNewLine & _
"DataFormato: " & DataFormato

End Sub
