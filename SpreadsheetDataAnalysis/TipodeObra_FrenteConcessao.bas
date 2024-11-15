Sub Aplica_TipoObra_FrenteConcessao()


'Classify each construction service and its concession front as described in monitoring spreadsheet

'Classifica cada obra e sua frente de concessão conforme descrição na planilha de acompanhamento

'Created by Matheus Nunes Reis on 15/08/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-ConstructionWorksMonitoring/64b7559b9dc94f31ea709efdb32b6a577a650594/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


Dim ws As Worksheet
Dim LastRow As Long
Dim ColunaTextoProcurado As String, ColunaTipoObra As String, ColunaFrenteConcessao As String

Set ws = ThisWorkbook.Sheets(1)

LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

ColunaTextoProcurado = "F" 'Coluna com descrição da obra contendo texto procurado
ColunaTipoObra = "D" 'Passarela, Retorno...
ColunaFrenteConcessao = "E" 'Obra de Melhoria, Ampliação de Capacidade...


For i = 2 To LastRow

    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Duplic", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Duplicação"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Ampliação de Capacidade"
    End If
        
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Adequação", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Adequação"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
    End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Margin", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Marginal"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
    End If
        
        'Manter "Intersec" antes de "Diamante" e "Trombeta"
        If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Intersec", vbTextCompare) > 0 Then
            ws.Cells(i, ColunaTipoObra).Value = "Intersecção"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
    
        'Manter "trevo" antes de "Diamante" e "Trombeta"
        If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Trev", vbTextCompare) > 0 Then
            ws.Cells(i, ColunaTipoObra).Value = "Trevo"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
            'Diamante e variações
        If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Diamante", vbTextCompare) > 0 Then
            ws.Cells(i, ColunaTipoObra).Value = "Diamante"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
            
        If (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Diamante", vbTextCompare) > 0) And (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Intersec", vbTextCompare) > 0) Then
            ws.Cells(i, ColunaTipoObra).Value = "Diamante"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
            
        If (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Diamante", vbTextCompare) > 0) And (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Trev", vbTextCompare) > 0) Then
            ws.Cells(i, ColunaTipoObra).Value = "Diamante"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
        'Trombeta e variações
        If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Trombeta", vbTextCompare) > 0 Then
            ws.Cells(i, ColunaTipoObra).Value = "Trombeta"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
            
        If (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Trombeta", vbTextCompare) > 0) And (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Intersec", vbTextCompare) > 0) Then
            ws.Cells(i, ColunaTipoObra).Value = "Trombeta"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
            
        If (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Trombeta", vbTextCompare) > 0) And (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Trev", vbTextCompare) > 0) Then
            ws.Cells(i, ColunaTipoObra).Value = "Trombeta"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Passarela", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Passarela"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
    End If
    
    If (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Acesso", vbTextCompare) > 0) And (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Melhor", vbTextCompare) > 0) Then
        ws.Cells(i, ColunaTipoObra).Value = "Melhoria Acesso"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
    End If
    
    If (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Acesso", vbTextCompare) > 0) And (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Exec", vbTextCompare) > 0) Then
        ws.Cells(i, ColunaTipoObra).Value = "Execução Acesso"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Ampliação de Capacidade"
    End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "PRF", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "PRF"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    
     If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "PPD", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "PPD"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Posto de Fiscalização", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Posto de Fiscalização"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "UOP", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "UOP"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    
    If (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "UOP", vbTextCompare) > 0) And (InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Delegacia", vbTextCompare) > 0) Then
        ws.Cells(i, ColunaTipoObra).Value = "UOP+Delegacia"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    
     If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Passagem Inferior", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Passagem Inferior"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
    End If
    
     If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Retorno", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Retorno"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
    End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Recuperação", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Recuperação"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Recuperação e Manutenção"
    End If
    
        If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Rotatória", vbTextCompare) > 0 Then
            ws.Cells(i, ColunaTipoObra).Value = "Rotatória"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
        'Rotatória repete mesmo (acento)
        If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Rotatoria", vbTextCompare) > 0 Then
            ws.Cells(i, ColunaTipoObra).Value = "Rotatória"
            ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
        End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Contorno", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Contorno"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Obra de Melhoria"
    End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Adicion", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Faixa Adicional"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Ampliação de Capacidade"
    End If
    
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Reversivel", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Faixa Reversível"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Ampliação de Capacidade"
    End If
    'Reversível repete mesmo (acento)
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Reversível", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "Faixa Reversível"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Ampliação de Capacidade"
    End If
    
     If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Posto Pesagem Veicular Fixo", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "PPV"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    'PPV repete mesmo (acento)
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Posto Pesagem Veícular Fixo", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "PPV"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    'PPV repete mesmo ("de" e sem acento)
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Posto de Pesagem Veicular Fixo", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "PPV"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    'PPV repete mesmo ("de" e com acento)
    If InStr(1, ws.Cells(i, ColunaTextoProcurado).Value, "Posto de Pesagem Veícular Fixo", vbTextCompare) > 0 Then
        ws.Cells(i, ColunaTipoObra).Value = "PPV"
        ws.Cells(i, ColunaFrenteConcessao).Value = "Sistemas de Operação"
    End If
    
Next i

MsgBox "Fim do Processo"

End Sub
