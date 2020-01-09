Option Compare Database

Private Sub cboCategoriasPrincipal_Click()
    Me.lstSubCategoriasPrincipal.Requery
End Sub

Private Sub cboCategoriasSecundario_Click()
    Me.lstSubCategoriasSecundario.Requery
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Private Sub cmdRelacao_Click()
    Call Relacao
End Sub

Private Sub Relacao()

    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intOrigem As Integer
    Dim intDestino As Integer
    Dim strSQL As String

    Set ctlOrigem = Me.lstSubCategoriasPrincipal
    Set ctlDestino = Me.lstSubCategoriasSecundario

    For intOrigem = 0 To ctlOrigem.ListCount - 1
        If ctlOrigem.Selected(intOrigem) Then
            For intDestino = 0 To ctlDestino.ListCount - 1
                If ctlDestino.Selected(intDestino) Then
                    strSQL = "INSERT INTO admCategorias ( codCategoria, codRelacao,Categoria,Descricao01,Descricao02 ) SELECT " & NovoCodigo("Select * from admCategorias", "codCategoria") & " as Categoria, " & _
                              ctlDestino.Column(0, intDestino) & " as Relacao,  codCategoria, '" & Me.cboCategoriasSecundario.Column(1) & "' as Descricao1 ,'" & ctlDestino.Column(1, intDestino) & "' as Descricao2 FROM admCategorias where codCategoria = " & ctlOrigem.Column(0, intOrigem) & ""
                    ExecutarSQL strSQL
                    ctlDestino.Selected(intDestino) = False
                End If
            Next intDestino
            ctlOrigem.Selected(intOrigem) = False
            Me.lstRelacionamentos.Requery
        End If
    Next intOrigem
    
    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing

End Sub

Private Sub lstRelacionamentos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String
Dim ctlRelacao As ListBox

Set ctlRelacao = Me.lstRelacionamentos

    Select Case KeyCode
    
        Case vbKeyDelete
           
            strSQL = "Delete * from admCategorias where codCategoria = " & ctlRelacao.Column(0)
            ExecutarSQL strSQL
            ctlRelacao.Requery
                    
    End Select
End Sub

Private Sub lstSubCategoriasPrincipal_Click()
Dim strSQL As String
Dim ctlRelacao As ListBox
Dim ctlPrincipal As ListBox

    Set ctlRelacao = Me.lstRelacionamentos
    Set ctlPrincipal = Me.lstSubCategoriasPrincipal

    strSQL = "SELECT codCategoria, Descricao01 as Categoria,Descricao02 as SubCategoria FROM lstCategoriasRelacionadas WHERE (((lstCategoriasRelacionadas.Categoria)='" & ctlPrincipal.Column(0) & "')) order by codCategoria Desc; "
    
    ctlRelacao.RowSource = strSQL
    ctlRelacao.Requery
    
End Sub
