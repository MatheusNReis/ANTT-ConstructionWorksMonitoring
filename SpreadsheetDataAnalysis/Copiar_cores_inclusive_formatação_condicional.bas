Sub CopiarCores()


'Copy colors from previous cell of pre-defined specific column including condicional formatting
    
'copiar cores da célula anterior de coluna específica pré-determinada incluindo formatação condicional

'Created by Matheus Nunes Reis on 06/08/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-ConstructionWorksMonitoring/64b7559b9dc94f31ea709efdb32b6a577a650594/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


    Dim cor As Long
    Dim ws As Worksheet
    Dim cel As Range
    Dim i As Long
    Dim Coluna As String
    
    'Dados de entrada: Inicializar variável Coluna,
                       'Valor inicial de i no loop for
    Coluna = "CR"
    
    ' Defina a planilha onde as células estão localizadas
    Set ws = ThisWorkbook.Sheets("Obras Concessionárias")
    ' Exibir mensagem para garantir que o código esteja sendo executado
    MsgBox "Iniciando a verificação de cores."
    ' Loop através das células de Coluna
    For i = 2 To ws.Cells(ws.Rows.Count, Coluna).End(xlUp).Row Step 2
        ' Verifique a cor de preenchimento da célula (incluindo formatação condicional)
        cor = ws.Cells(i, Coluna).DisplayFormat.Interior.Color
        ' Se a célula tiver cor de preenchimento
        If cor <> -4142 Then ' -4142 é a cor padrão quando não há cor de preenchimento
            ' Copie a cor para a linha seguinte
            ws.Cells(i + 1, Coluna).Interior.Color = cor
            ' Adicione uma mensagem para verificar se o código está sendo executado
            'MsgBox "Cor copiada para linha " & i + 1
            ' Imprima a cor detectada no Immediate Window
            'Debug.Print "Cor detectada: " & cor
        End If
    Next i
    ' Exibir mensagem para garantir que o loop tenha sido concluído
    MsgBox "Fim da cópia das cores."
End Sub

