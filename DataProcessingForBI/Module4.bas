Sub AdicionarCodigos()

'Obras_BI.xlsm

'Creation of codes and counting data for construction improvement work, capacity expansion and operation.

'Criação de códigos e contagem de dados para trabalhos de melhoria, ampliação de capacidade e operação.

'Created by Matheus Nunes Reis on 12/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-ConstructionWorksMonitoring/64b7559b9dc94f31ea709efdb32b6a577a650594/LICENSE
'MIT License. Copyright © 2024 MatheusNReis

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim AMPCount As Long
    Dim MELCount As Long
    Dim OPCount As Long
    Dim AMPKeywords As Variant
    Dim MELKeywords As Variant
    Dim OPKeywords As Variant
    Dim cellValue As String

    ' Configura planilha e determina a última linha na coluna D
    Set ws = Workbooks("Obras_BI.xlsm").Sheets("TabDimensao")
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    ' Inicialização de contagem
    AMPCount = 1
    MELCount = 1
    OPCount = 1
    
    ' Define keywords para cada categoria
    AMPKeywords = Array("Adic", "Duplic", "Futura Pista", "Nova Pista")
    MELKeywords = Array("OAE", "Passarela", "Trevo", "Acesso", "Alça", "bus", "Barreira", _
                        "Marginais", "Faixa Revers", "Retorno")
    OPKeywords = Array("PPD", "Pesagem", "UOP", "SAT", "DAI", "TV", "Iluminação", "Tráfego", _
                       "Mensage", "Fibra", "Velocidade", "Meteoro", "Wireless")

    
    For i = 2 To lastRow
        cellValue = ws.Cells(i, 4).Value

        If ContainsKeyword(cellValue, AMPKeywords) Then
            'AMP keywords encontrada
            ws.Cells(i, 3).Value = "AMP" & Format(AMPCount, "00")
            AMPCount = AMPCount + 1
        
        
        ElseIf ContainsKeyword(cellValue, MELKeywords) Then
            'MEL keyword encontrada
            ws.Cells(i, 3).Value = "MEL" & Format(MELCount, "00")
            MELCount = MELCount + 1
        
        
        ElseIf ContainsKeyword(cellValue, OPKeywords) Then
            'OP keywords encontrada
            ws.Cells(i, 3).Value = "OP" & Format(OPCount, "00")
            OPCount = OPCount + 1
        End If
    Next i
    
    'Notifica que o processamento foi concluído
    MsgBox "Códigos gerados na planilha TabDimensao!"
End Sub
