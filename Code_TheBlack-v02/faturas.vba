Option Compare Database
Option Explicit

Private Sub BaixarEstoque()
Dim strSQL As String

strSQL = "INSERT INTO Estoque ( DataDeEmissao, PontoDeVenda, TipoDeMovimento, Motivo, " & _
         "DescricaoDoProduto, Quantidade, codFatura ) SELECT Format(Now(),'dd/mm/yy') AS Emissao, " & _
         "NomeDaLoja() AS Loja, 'Saida' AS TipoMovimento, 'Venda' AS Motivo, FaturasItens.DescricaoDoProduto, " & _
         "FaturasItens.Quantidade, FaturasItens.codFatura FROM FaturasItens WHERE (((FaturasItens.codFatura) Not In " & _
         "(SELECT DISTINCT Estoque.codFatura FROM Estoque WHERE ((Not (Estoque.codFatura) Is Null)))) AND ((FaturasItens.Referencia)='peças')) " & _
         "ORDER BY FaturasItens.DescricaoDoProduto"


ExecutarSQL strSQL

End Sub


Private Sub cboCadastro_Click()
    Me.txtNome = Me.cboCadastro.Column(1)
    Me.txtEndereco = Me.cboCadastro.Column(2)
    Me.txtBairro = Me.cboCadastro.Column(3)
    Me.txtCep = Me.cboCadastro.Column(4)
    Me.txtMunicipio = Me.cboCadastro.Column(5)
    Me.txtEstado = Me.cboCadastro.Column(6)
    Me.cboVeiculo.Requery
End Sub

Private Sub cboCadastro_NotInList(NewData As String, Response As Integer)
Dim DB As DAO.Database
Dim rst As DAO.Recordset
Dim Pergunta As Variant

On Error GoTo ErrHandler

Pergunta = MsgBox("O Cliente: " & NewData & "  não faz parte da lista." & vbCrLf & "Deseja acrescentá-lo?", vbQuestion + vbYesNo)


'Pergunta se deseja acrescentar o novo item
If Pergunta = vbYes Then
    Set DB = CurrentDb()
    'Abre a tabela, adiciona o novo item e atualiza a combo
    Set rst = DB.OpenRecordset("Cadastros")
    With rst
        .AddNew
        !codCadastro = NovoCodigo("Select * from Cadastros", "codCadastro")
        !TipoCadastro = "Clientes"
        !Nome = NewData
        .Update
        Response = acDataErrAdded
        .Close
    End With
Else
    Response = acDataErrDisplay
End If

ExitHere:
Set rst = Nothing
Set DB = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description & vbCrLf & Err.Number & _
vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
Resume ExitHere

End Sub



Private Sub cmdOrdem_Click()
On Error GoTo Err_cmdOrdem_Click

    Dim stDocName As String

'    Call CalcularPedido
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

    stDocName = "OrdensDeServicos"
    DoCmd.OpenReport stDocName, acPreview, , "Faturas.codFatura = " & Me.Codigo

Exit_cmdOrdem_Click:
    Exit Sub

Err_cmdOrdem_Click:
    MsgBox Err.Description
    Resume Exit_cmdOrdem_Click
End Sub

Private Sub cmdReceber_Click()
Dim strNotas As String

Dim ValorPago As Currency
Dim ValorRecebido As Currency
Dim Resto As Currency

Dim rstEspecies As DAO.Recordset

Set rstEspecies = CurrentDb.OpenRecordset("Select * from FaturasEspecies where Parcelado = false and codFatura = " & Me.Codigo & " Order by codRecebimento")

'Dim ctlParcelamento As ComboBox
'Dim ctlEspecie As ComboBox
Dim ctlCliente As ComboBox

'Set ctlParcelamento = Me.cboParcelamento
'Set ctlEspecie = Me.cboEspecie
Set ctlCliente = Me.cboCadastro

While Not rstEspecies.EOF

    Me.DataDeEmissao = Format(Now(), "dd/mm/yy")
    Me.StatusDoPedido = "Fatura"
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

    ' Carrega dados para envio a notas do movimento financeiro.
    strNotas = "Cliente: " & ctlCliente.Column(1) & Chr(13) & Chr(10) & "eMail: " & ctlCliente.Column(9) & Chr(13) & Chr(10) & Chr(13) & _
                "Telefone: " & ctlCliente.Column(10) & " Comercial: " & ctlCliente.Column(11) & Chr(13) & Chr(10) & vbNewLine & _
                "CNPJ-CPF: " & ctlCliente.Column(7) & Chr(13) & Chr(10) & "IE-RG: " & ctlCliente.Column(8) & Chr(13) & Chr(10) & _
                "Endereço: " & ctlCliente.Column(2) & " - " & ctlCliente.Column(3) & Chr(13) & Chr(10) & _
                ctlCliente.Column(4) & " - " & ctlCliente.Column(5) & " - " & ctlCliente.Column(6)


'    If Not IsNull(rstEspecies.Fields("Parcelas_Valor")) And Not IsNull(rstEspecies.Fields("Especie_Valor")) Then

        ValorPago = rstEspecies.Fields("ValorRecebido")

        ValorRecebido = ValorPago - (ValorPago / 100 * rstEspecies.Fields("Especie_Valor"))

        Resto = ValorPago - ValorRecebido

        If ValorRecebido > 0 Then

            LancarMovimento Me.Codigo, _
                                Format(Me.DataDeEmissao, "dd/mm/yy"), _
                                ValorRecebido + Resto, _
                                IIf(IsNull(rstEspecies.Fields("Parcelas_Valor")), 1, rstEspecies.Fields("Parcelas_Valor")), _
                                rstEspecies.Fields("Especie"), _
                                 Me.cboCadastro.Column(1), _
                                "Receita", _
                                NomeDaLoja, _
                                "Vendas", strNotas

        End If

        If Resto > 0 Then

            LancarMovimento Me.Codigo, _
                                Format(Me.DataDeEmissao, "dd/mm/yy"), _
                                Resto, _
                                IIf(IsNull(rstEspecies.Fields("Parcelas_Valor")), 1, rstEspecies.Fields("Parcelas_Valor")), _
                                rstEspecies.Fields("Especie"), _
                                Me.cboCadastro.Column(1), _
                                "Despesa", _
                                NomeDaLoja, _
                                "Vendas", strNotas
        End If

'    End If
    rstEspecies.Edit
    rstEspecies.Fields("Parcelado") = True
    rstEspecies.Update

    rstEspecies.MoveNext


Wend

BaixarEstoque

MsgBox "Recebimento Concluído!", vbInformation + vbOKOnly, "Recebimento"

rstEspecies.Close

Me.FaturasRecebimentos.Requery
Me.Recalc

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
Dim strSQL As String
strSQL = "Select * from Faturas"

    If Me.NewRecord Then
        Me.Codigo = NovoCodigo(strSQL, Me.Codigo.ControlSource)
        Me.DataDeEmissao = Format(Now(), "dd/mm/yy")
    End If

End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

'    Call Val_Pedido
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_admPesquisar.lstCadastro.Requery
    DoCmd.Close

Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String

'    Call CalcularPedido
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

    stDocName = "Faturas"
    DoCmd.OpenReport stDocName, acPreview, , "Faturas.codFatura = " & Me.Codigo

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click

End Sub



Private Sub cmdCliente_Click()
On Error GoTo Err_cmdCliente_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Clientes"

    stLinkCriteria = "[codCadastro]=" & Me![cboCadastro]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdCliente_Click:
    Exit Sub

Err_cmdCliente_Click:
    MsgBox Err.Description
    Resume Exit_cmdCliente_Click

End Sub



