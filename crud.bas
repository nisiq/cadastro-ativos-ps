Option Explicit
 
'Variaveis para uso geral nesse modulo
Private lin     As Long
Private i       As Long
Private resposta    As VbMsgBoxResult
Private rng         As Range
Dim responsavel     As String
 
'Salvar os dados na planilha Base Dados
Public Sub Salva_Dados_Base()
 
 
'1º Restrição - Se o campo Responsável estiver vazio
If PlanForm.Range("Campo3") = Empty Then
    MsgBox "Obrigatório Preencher o Responsável.", vbExclamation, "Cadastro de Ativo"
    Exit Sub
End If
'2º Restrição - Se o campo Local estiver vazio
If PlanForm.Range("Campo2") = Empty Then
    MsgBox "Obrigatório Preencher o Local.", vbExclamation, "Cadastro de Ativo"
    Exit Sub
 
End If
'3º Restrição - Se o campo Denominação do Imobilizado estiver vazio
If PlanForm.Range("Campo1") = Empty Then
    MsgBox "Obrigatório Preencher Denominacao do Imobilizado.", vbExclamation, "Cadastro de Ativo"
    Exit Sub
Else
 
lin = PlanBase.Cells(PlanBase.Cells.Rows.Count, "B").End(xlUp).Row + 1
lin = 2
 
'4º Restrição - Se o ativo ja estiver cadastrado na base de dados
Do Until PlanBase.Cells(lin, 1).Value = Empty
    If PlanForm.Range("Campo0").Value = PlanBase.Cells(lin, 1).Value Then
        MsgBox "Ativo já está Cadastrado na Base de Dados", vbCritical, "Cadastro de Ativo"
        Exit Sub
    PlanForm.Select
    End If
lin = lin + 1
Loop
 
'Inicia o salvamento das informações na base de dados
For i = 0 To 13
    PlanBase.Cells(lin, 1 + i) = PlanForm.Range("Campo" & i).Value
    'PlanBase.Cells(lin, 6) = VBA.Environ("COMPUTERNAME")
Next i
    MsgBox "Ativo Cadastrado com Sucesso na Base de Dados", vbOKOnly, "Cadastro de Ativo"
End If
End Sub
 
 
'Consultar dados na base de dados
Public Sub Consulta_Produtos_Base()
 
'1º Restrição - Se o campo estiver vazio
If PlanForm.Range("Campo0") = Empty Then
    MsgBox "Preencha o Imobilizado para Realizar a Busca", vbCritical, "Cadastro de Ativo"
    Exit Sub
Else
 
Set rng = PlanBase.Range("A:A").Find(PlanForm.Range("Campo0").Value, After:=PlanBase.Range("A1"))
 
    If rng Is Nothing Then
        PlanForm.Range("Campo0").Value = Empty
        MsgBox "Codigo Inexistente.", vbCritical, "Cadastro de Ativo"
        Exit Sub
    Else
        For i = 0 To 13
            PlanForm.Range("Campo1") = rng.Offset(0, 1).Value 'Descricao do Ativo
            PlanForm.Range("Campo2") = rng.Offset(0, 2).Value 'Local
            PlanForm.Range("Campo3") = rng.Offset(0, 3).Value 'Responsavel
            PlanForm.Range("Campo4") = rng.Offset(0, 4).Value 'Val. Aquis.
            PlanForm.Range("Campo5") = rng.Offset(0, 5).Value 'Depreciacao ac.
            PlanForm.Range("Campo6") = rng.Offset(0, 6).Value 'Valor contabil
            PlanForm.Range("Campo7") = rng.Offset(0, 7).Value 'Moeda
            PlanForm.Range("Campo8") = rng.Offset(0, 8).Value 'N Inventario
            PlanForm.Range("Campo9") = rng.Offset(0, 9).Value 'Centro
            PlanForm.Range("Campo10") = rng.Offset(0, 10).Value 'Classe Imobilizado
            PlanForm.Range("Campo11") = rng.Offset(0, 11).Value 'Incroporacao
            PlanForm.Range("Campo12") = rng.Offset(0, 12).Value 'Centro de Custo
            PlanForm.Range("Campo13") = rng.Offset(0, 13).Value 'Centro de Custo
        Next
    End If
End If
 
End Sub
 
'Atualiza os dados na base
Public Sub Atualizar_Dados_Base()
 
'1º Restrição - Se o campo estiver vazio
If PlanForm.Range("Campo0") = Empty Then
    MsgBox "Preencha o codigo do Imobilizado para atualizar", vbCritical, "Cadastro de Ativo"
    Exit Sub
Else
 
Set rng = PlanBase.Range("A:A").Find(PlanForm.Range("Campo0").Value, After:=PlanBase.Range("A1"))
 
If rng Is Nothing Then
        PlanForm.Range("Campo0").Value = Empty
        MsgBox "Codigo Inexistente.", vbExclamation, "Cadastro de Ativo"
        Exit Sub
Else
    For i = 1 To 13
        rng.Offset(0, i).Value = PlanForm.Range("Campo" & i).Value
    Next
    MsgBox "Alteração realizada com sucesso.", vbOKOnly, "Cadastro de Ativo"
End If
End If
 
End Sub
Public Sub Deletar_Dados_Base()
    ' 1º Restrição - Se o campo estiver vazio
    If PlanForm.Range("Campo0") = Empty Then
        MsgBox "Preencha o codigo do Imobilizado para remover", vbCritical, "Cadastro de Ativo"
        Exit Sub
    Else
        ' Exibe o prompt para selecionar o responsável
        responsavel = InputBox("Digite o nome do responsável pela deleção:", "Selecionar Responsável")
 
        ' Verifica se o usuário cancelou o input
        If responsavel = "" Then
            MsgBox "Operação cancelada.", vbInformation, "Cadastro de Ativo"
            Exit Sub
        End If
 
        Set rng = PlanBase.Range("A:A").Find(PlanForm.Range("Campo0").Value, After:=PlanBase.Range("A1"))
 
        If rng Is Nothing Then
            PlanForm.Range("Campo0").Value = Empty
            MsgBox "Codigo Inexistente.", vbExclamation, "Cadastro de Ativo"
            Exit Sub
        Else
            ' Armazena o nome do responsável pela remoção na coluna "O" (coluna 15)
            rng.Offset(0, 14).Value = responsavel
            ' Pintar a linha de vermelho
            rng.Resize(, 16).Interior.Color = RGB(255, 0, 0) ' Vermelho
            rng.Offset(0, 15).Value = Now() ' Data da remoção
 
            ' Copiar informações do ativo removido para a planilha "Ativos Removidos"
            Dim wsAtivosRemovidos As Worksheet
            Set wsAtivosRemovidos = ThisWorkbook.Sheets("Ativos Removidos")
 
            Dim ultimaLinha As Long
            ultimaLinha = wsAtivosRemovidos.Cells(wsAtivosRemovidos.Rows.Count, "A").End(xlUp).Row + 1
 
            ' Copiar informações para a planilha de ativos removidos
            rng.EntireRow.Copy wsAtivosRemovidos.Cells(ultimaLinha, 1)
 
            ' Remover a linha da tabela atual
            rng.EntireRow.Delete
 
            MsgBox "Produto marcado para deleção pelo responsável: " & responsavel, vbOKOnly, "Cadastro de Ativo"
        End If
    End If
End Sub
 
 
'Limpar os campos do formulario
Public Sub limpar_Campos_Formulario()
    For i = 0 To 13
        PlanForm.Range("Campo" & i).Select
        Selection.ClearContents
    Next i
    PlanForm.Range("Campo0").Select
End Sub
