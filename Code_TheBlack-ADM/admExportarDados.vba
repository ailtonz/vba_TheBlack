Option Compare Database

Private Sub cmdEnviar_Click()
    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim strSQL As String
    
    Set ctlOrigem = Me.lstVendasDisponiveis
    Set ctlDestino = Me.lstVendasSelecionadas
          
    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If ctlOrigem.Selected(intCurrentRow) Then
            strSQL = "UPDATE Movimentos SET Movimentos.Selecao = Yes WHERE (((Movimentos.DataDeEmissao)=Format(#" & ctlOrigem.Column(0, intCurrentRow) & "#,'mm/dd/yyyy')))"
            ExecutarSQL strSQL
            Me.lstVendasSelecionadas.Requery
            Me.lstVendasDisponiveis.Requery
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing
End Sub

Private Sub cmdEnviarTodos_Click()
    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim strSQL As String
    
    Set ctlOrigem = Me.lstVendasDisponiveis
    Set ctlDestino = Me.lstVendasSelecionadas
          
    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If Not IsNull(ctlOrigem.Column(0, intCurrentRow)) Then
            strSQL = "UPDATE Movimentos SET Movimentos.Selecao = Yes WHERE (((Movimentos.DataDeEmissao)=Format(#" & ctlOrigem.Column(0, intCurrentRow) & "#,'mm/dd/yyyy')))"
            ExecutarSQL strSQL
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Me.lstVendasSelecionadas.Requery
    Me.lstVendasDisponiveis.Requery

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing

End Sub

Private Sub cmdExportarEstoque_Click()
    ExportarXLS "expEstoque"
End Sub

Private Sub cmdExportarRecebimentos_Click()
    ExportarXLS "expMovimentos"
End Sub

Private Sub cmdRemover_Click()
    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim strSQL As String
    
    Set ctlOrigem = Me.lstVendasSelecionadas
    Set ctlDestino = Me.lstVendasDisponiveis
    
    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If ctlOrigem.Selected(intCurrentRow) Then
            strSQL = "UPDATE Movimentos SET Movimentos.Selecao = No WHERE (((Movimentos.DataDeEmissao)=Format(#" & ctlOrigem.Column(0, intCurrentRow) & "#,'mm/dd/yyyy')))"
            ExecutarSQL strSQL
            Me.lstVendasSelecionadas.Requery
            Me.lstVendasDisponiveis.Requery
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing
End Sub

Private Sub cmdRemoverTodos_Click()
    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim strSQL As String
    
    Set ctlOrigem = Me.lstVendasSelecionadas
    Set ctlDestino = Me.lstVendasDisponiveis
          
    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If Not IsNull(ctlOrigem.Column(0, intCurrentRow)) Then
            strSQL = "UPDATE Movimentos SET Movimentos.Selecao = No WHERE (((Movimentos.DataDeEmissao)=Format(#" & ctlOrigem.Column(0, intCurrentRow) & "#,'mm/dd/yyyy')))"
            ExecutarSQL strSQL
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Me.lstVendasSelecionadas.Requery
    Me.lstVendasDisponiveis.Requery

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing


End Sub

Private Sub Form_Close()
    Call cmdRemoverTodos_Click
End Sub

Private Sub lstVendasDisponiveis_DblClick(Cancel As Integer)
   Call cmdEnviar_Click
End Sub

Private Sub lstVendasSelecionadas_DblClick(Cancel As Integer)
    Call cmdRemover_Click
End Sub



