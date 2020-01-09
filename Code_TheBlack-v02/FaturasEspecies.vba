Option Compare Database

Private Sub cboEspecie_Click()
Dim strSQL As String
Dim ctlParcelamento As ComboBox
Dim ctlEspecie As ComboBox
    
Set ctlParcelamento = Me.cboParcelamento
Set ctlEspecie = Me.cboEspecie
   
ctlParcelamento.Value = ""
   
strSQL = "SELECT admCategorias.Categoria, admCategorias.Descricao01 " & _
         "FROM admCategorias WHERE (((admCategorias.codCategoria) In " & _
         "(Select codRelacao from admCategorias where Categoria = '" & ctlEspecie.Column(2) & "')));"

ctlParcelamento.RowSource = strSQL
ctlParcelamento.Requery
ctlParcelamento.Value = ctlEspecie.Column(3)
Me.Especie_Valor = ctlEspecie.Column(1)

Set ctlParcelamento = Nothing
Set ctlEspecie = Nothing
End Sub

Private Sub cboParcelamento_Click()
    Me.Parcelas_Valor = Me.cboParcelamento.Column(1)
End Sub
