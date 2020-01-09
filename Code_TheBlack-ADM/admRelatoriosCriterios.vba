Option Compare Database


Private Sub Descricao_Click()
    Me.Valor = Me.Descricao.Column(0)
End Sub

Private Sub Descricao_Enter()
    Me.Descricao.RowSource = Me.OrigemDaLinha.Value
    Me.Descricao.ColumnWidths = Me.LarguraDasColunas
End Sub


