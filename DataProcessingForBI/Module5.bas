'Obras_BI.xlsm

'Module4 support function that checks if a cell contains any keyword in the list

'Função de suporte ao módulo4 que verifica se uma célula contém qualquer keyword da lista

'Created by Matheus Nunes Reis on 13/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-ConstructionWorksMonitoring/64b7559b9dc94f31ea709efdb32b6a577a650594/LICENSE
'MIT License. Copyright © 2024 MatheusNReis

 
Function ContainsKeyword(cellValue As String, keywords As Variant) As Boolean
    Dim keyword As Variant
    For Each keyword In keywords
        If InStr(1, cellValue, keyword, vbTextCompare) > 0 Then
            ContainsKeyword = True
            Exit Function
        End If
    Next keyword
    ContainsKeyword = False
End Function
