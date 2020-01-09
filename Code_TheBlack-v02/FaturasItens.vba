Option Compare Database
Option Explicit

'Private Sub cboCor_NotInList(NewData As String, Response As Integer)
''Permite adicionar a editora à lista
'Dim db As DAO.Database
'Dim rst As DAO.Recordset
'
'On Error GoTo ErrHandler
'
''Pergunta se deseja acrescentar o novo item
'If Confirmar("A Cor: " & NewData & "  não faz parte da " & _
'"lista." & vbCrLf & "Deseja acrescentá-lo?") = True Then
'    Set db = CurrentDb()
'    'Abre a tabela, adiciona o novo item e atualiza a combo
'    Set rst = db.OpenRecordset("admCategorias")
'    With rst
'        .AddNew
'        !codCor = NovoCodigo("FatCores", "codCor")
'        !Cor = NewData
'        .Update
'        Response = acDataErrAdded
'        .Close
'    End With
'Else
'    Response = acDataErrDisplay
'End If
'
'ExitHere:
'Set rst = Nothing
'Set db = Nothing
'Exit Sub
'
'ErrHandler:
'MsgBox Err.Description & vbCrLf & Err.Number & _
'vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
'Resume ExitHere
'End Sub

'Private Sub cboTamanho_NotInList(NewData As String, Response As Integer)
''Permite adicionar a editora à lista
'Dim db As DAO.Database
'Dim rst As DAO.Recordset
'
'On Error GoTo ErrHandler
'
''Pergunta se deseja acrescentar o novo item
'If Confirmar("O Tamanho: " & NewData & "  não faz parte da " & _
'"lista." & vbCrLf & "Deseja acrescentá-lo?") = True Then
'    Set db = CurrentDb()
'    'Abre a tabela, adiciona o novo item e atualiza a combo
'    Set rst = db.OpenRecordset("FatTamanhos")
'    With rst
'        .AddNew
'        !codTamanho = NovoCodigo("FatTamanhos", "codTamanho")
'        !Tamanho = NewData
'        .Update
'        Response = acDataErrAdded
'        .Close
'    End With
'Else
'    Response = acDataErrDisplay
'End If
'
'ExitHere:
'Set rst = Nothing
'Set db = Nothing
'Exit Sub
'
'ErrHandler:
'MsgBox Err.Description & vbCrLf & Err.Number & _
'vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
'Resume ExitHere
'
'End Sub

'Private Sub DescricaoDoProduto_NotInList(NewData As String, Response As Integer)
''Permite adicionar a editora à lista
'Dim db As DAO.Database
'Dim rst As DAO.Recordset
'
'On Error GoTo ErrHandler
'
''Pergunta se deseja acrescentar o novo item
'If Confirmar("O Produtos: " & NewData & "  não faz parte da lista." & vbCrLf & "Deseja acrescentá-lo?") = True Then
'    Set db = CurrentDb()
'    'Abre a tabela, adiciona o novo item e atualiza a combo
'    Set rst = db.OpenRecordset("FatProdutos")
'    With rst
'        .AddNew
'        !codProduto = NovoCodigo("FatProdutos", "codProduto")
'        !DescricaoDoProduto = NewData
'        .Update
'        Response = acDataErrAdded
'        .Close
'    End With
'Else
'    Response = acDataErrDisplay
'End If
'
'ExitHere:
'Set rst = Nothing
'Set db = Nothing
'Exit Sub
'
'ErrHandler:
'MsgBox Err.Description & vbCrLf & Err.Number & _
'vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
'Resume ExitHere
'
'End Sub

Private Sub cboDescricaoDoProduto_Click()
    Me.ValorUnitario = Me.cboDescricaoDoProduto.Column(1)
End Sub

Private Sub cboReferencia_Click()
    Me.cboDescricaoDoProduto.Requery
End Sub

Private Sub Quantidade_Exit(Cancel As Integer)
    If Not IsNull(Me.DescricaoDoProduto) Then Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
End Sub

'Private Sub StatusProduto_Click()
'    Me.StatusProduto = Me.StatusProduto.Column(1)
'End Sub

Private Sub ValorUnitario_Exit(Cancel As Integer)
    If Not IsNull(Me.DescricaoDoProduto) Then Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
End Sub

