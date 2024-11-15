Sub ReorganizarDados()

'Obras_BI.xlsm

'Data reorganization and processing for construction works monitoring

'Reorganização e processamento de dados para acompanhamento de obras

'Created by Matheus Nunes Reis on 04/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-ConstructionWorksMonitoring/64b7559b9dc94f31ea709efdb32b6a577a650594/LICENSE
'MIT License. Copyright © 2024 MatheusNReis

    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long, j As Long
    Dim linhaDestinoPrev As Long, linhaDestinoExec As Long
    Dim percentPlanejado As Variant
    Dim itemPER As String
    Dim codigo As String
    Dim descricao As String
    Dim concessionaria As String
    Dim dataColuna As String
    Dim Dia As Integer, Mes As Integer, Ano As Integer
    
    ' Definir as planilhas de origem e destino
    Set wsOrigem = Workbooks("acompanhamento_fisico_mensal_Cocessionaria.xlsx").Sheets("CONCESSIONARIA")
    Set wsDestino = Workbooks("Obras_BI.xlsm").Sheets("Previsto_Executado")
    
    ' Definir a concessionária com base no nome da aba da planilha de origem
    concessionaria = wsOrigem.Name
    
    ' Limpar somente a aba específica da planilha de destino
    wsDestino.Cells.Clear
    
    ' Cabeçalho na planilha de destino
    wsDestino.Cells(1, 1).Value = "Concessionária"
    wsDestino.Cells(1, 2).Value = "Código"
    wsDestino.Cells(1, 3).Value = "Descrição"
    wsDestino.Cells(1, 4).Value = "Data"
    wsDestino.Cells(1, 5).Value = "% Previsto"
    wsDestino.Cells(1, 6).Value = "% Executado"
    wsDestino.Cells(1, 7).Value = "Observações"
    
    linhaDestinoPrev = 2    'É a linha na planilha Planejado_Executado do arquivo Obras_BI.xlsm, coluna %Previsto
    linhaDestinoExec = 2    'É a linha na planilha Planejado_Executado do arquivo Obras_BI.xlsm, coluna %Executado
    
    ' Encontrar a última linha da planilha de origem
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row
    
    ' Loop pelas linhas da planilha de origem (variável i)
    For i = 7 To ultimaLinha
    
    
        If wsOrigem.Cells(i, 1).MergeCells Then     'Verifica se a célula é mesclada
            itemPER = wsOrigem.Cells(i, 1).MergeArea.Cells(1, 1).Value      'ItemPER recebe o primeiro valor da área mesclada
        Else
            itemPER = wsOrigem.Cells(i, 1).Value    'Se a célula não for mescalada, captura o valor da linha atual
        End If
        
        codigo = wsOrigem.Cells(i, 2).Value
        descricao = wsOrigem.Cells(i, 3).Value
        
        ''' Verificar se a coluna "L" contém "previsto" (Preenche colunas "Concessionária, Código, Descrição, Data e %Previsto")
        If wsOrigem.Cells(i, 12).Value = "previsto" Then
            ' Loop pelas colunas de datas e percentuais planejados (colunas AA a AM)
            For j = 27 To 39  ' AA é a coluna 27 e AM é a coluna 39
                
                If wsOrigem.Cells(6, j).MergeCells Then     'Verifica se a célula é mesclada
                    dataColuna = DateSerial(Year(wsOrigem.Cells(6, j + 1).Value), Day(wsOrigem.Cells(6, j + 1).Value), Month(wsOrigem.Cells(6, j + 1).Value) - 1)
                    wsDestino.Cells(linhaDestinoPrev, 7).Value = wsOrigem.Cells(6, j).MergeArea.Cells(1, 1).Value   'Captura o primeiro valor da área mesclada
                Else
                    dataColuna = DateSerial(Year(wsOrigem.Cells(6, j).Value), Day(wsOrigem.Cells(6, j).Value), Month(wsOrigem.Cells(6, j).Value))
                End If
                
                percentPlanejado = wsOrigem.Cells(i, j).Value
                
                If Not IsEmpty(dataColuna) Then
                    ' Adicionar os valores à planilha de destino se valores de dataColuna para a linha avaliada não forem vazios
                    'OBS: em teoria, a variável dataColuna nunca estará vazia e este bloco if será sempre percorrido
                    wsDestino.Cells(linhaDestinoPrev, 1).Value = concessionaria
                    wsDestino.Cells(linhaDestinoPrev, 2).Value = codigo
                    wsDestino.Cells(linhaDestinoPrev, 3).Value = descricao
                    wsDestino.Cells(linhaDestinoPrev, 4).Value = dataColuna
                    
                    If percentPlanejado = "" Or Trim(percentPlanejado) = "" Or Trim(Replace(percentPlanejado, Chr(160), " ")) = "" Then    'a célula contém um caractere de espaço não quebrável (às vezes inserido ao pressionar Alt + 0160)
                        wsDestino.Cells(linhaDestinoPrev, 5).Value = Format(0, "Percent") 'Preenche na coluna 5 da planilha destino por ser %previsto
                    Else
                        wsDestino.Cells(linhaDestinoPrev, 5).Value = Format(wsOrigem.Cells(i, j).Value, "Percent") 'Preenche na coluna 5 da planilha destino por ser %previsto
                    End If
                    
                    linhaDestinoPrev = linhaDestinoPrev + 1
                End If
            Next j
        End If
        
        ''' Verificar se a coluna "L" contém "executado" (Preenche coluna "%Executado")
        If wsOrigem.Cells(i, 12).Value = "executado" Then
            ' Loop pelas colunas de datas e percentuais planejados (colunas AA a AM)
            For j = 27 To 39  ' AA é a coluna 27 e AM é a coluna 39
                
                    percentPlanejado = wsOrigem.Cells(i, j).Value
                    
                    If percentPlanejado = "" Or Trim(percentPlanejado) = "" Or Trim(Replace(percentPlanejado, Chr(160), " ")) = "" Then    'a célula contém um caractere de espaço não quebrável (às vezes inserido ao pressionar Alt + 0160)
                        wsDestino.Cells(linhaDestinoExec, 6).Value = Format(0, "Percent") 'Preenche na coluna 6 da planilha destino por ser %executado
                    Else
                        wsDestino.Cells(linhaDestinoExec, 6).Value = Format(wsOrigem.Cells(i, j).Value, "Percent") 'Preenche na coluna 6 da planilha destino por ser %executado
                    End If
            
                    linhaDestinoExec = linhaDestinoExec + 1
            Next j
        End If
        
        
    Next i
    
    MsgBox "Dados reorganizados com sucesso!"

End Sub

