Sub AcompanhamentoFisicoObras()


'Vertically unmerges and duplicates information from the construction service monitoring spreadsheet of the concessionaires
'and create column naming the current concessionaire analyzed.
'Applied in spreadsheet parts determined by user.
    
'Desmescla verticalmente e duplica informações da planilha acompanhamento das obras das concessionárias e cria coluna com
'nome da concessionária estudada.
'Aplicado em partes da planilha determinadas pelo usuário.

'Created by Matheus Nunes Reis on 06/08/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-ConstructionWorksMonitoring/64b7559b9dc94f31ea709efdb32b6a577a650594/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


    Dim lastRow As Long
    Dim i As Long
    Dim fileName As String
    
    ' Encontra a última linha preenchida com dados na coluna "A" a partir da sétima linha
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    If lastRow < 7 Then  ' Se a última linha encontrada for menor que 7, não há dados suficientes
        MsgBox "Não há dados suficientes na coluna A a partir da sétima linha.", vbExclamation
        Exit Sub
    End If
    
    
     ' Loop através das linhas para desmesclar e copiar
    For i = 7 To lastRow Step 2 ' Inicia da linha 7 e pula de 2 em 2, até a última linha (LastRow)
        ' Desmescla células e copia conteúdo para a próxima linha
        Range("A" & i & ":J" & i).MergeCells = False
        Range("A" & i & ":J" & i).Copy
        Range("A" & i + 1 & ":J" & i + 1).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
    Next i
    
    
    ' Insere uma nova coluna A
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
     'Obtém o nome do arquivo atual
    fileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
    
    ' Loop para preencher as células da coluna A com números de linha
    For i = 7 To lastRow + 1
        Cells(i, 1).Value = fileName
    Next i
    
    ' Copia a formatação da coluna B para a coluna A
    Columns("B:B").Copy
    Columns("A:A").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
     ' Escreve o título "Concessionária" na célula A6
    Range("A5").Value = "Concessionária"
    
End Sub
