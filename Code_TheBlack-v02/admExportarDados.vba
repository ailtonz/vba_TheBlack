Option Compare Database

Private Sub cmdExportar_Click()

Dim ctlOrigem As ListBox
Dim intCurrentRow As Integer
Dim strDatas As String
Dim strSQL01 As String
Dim strSQL02 As String

Set ctlOrigem = Me.lstDados

strDatas = ""

If Me.cboConteudo.Column(0) <> "" Then

    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If ctlOrigem.Selected(intCurrentRow) Then
            strDatas = strDatas & "#" & Format(ctlOrigem.Column(0, intCurrentRow), "mm/dd/yyyy") & "#, "
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow
    
    strDatas = Left(strDatas, Len(strDatas) - 2) & " "

    ' ESTOQUE
    If Me.cboConteudo.Column(0) = "ESTOQUE" Then
        strSQL01 = "INSERT INTO tmpEstoque ( DataDeEmissao, PontoDeVenda, codFatura, DescricaoDoProduto, Quantidade, TipoDeMovimento, Motivo ) "
        strSQL02 = strSQL01 & " Select * from (SELECT * FROM " & Me.cboConteudo.Column(1) & " WHERE (((" & Me.cboConteudo.Column(2) & ") In (" & strDatas & ")))) as tmp"
           
        ExecutarSQL "Delete * from tmpEstoque"
        
        ExecutarSQL strSQL02
        
        ExportarXLS Me.cboConteudo.Column(0), "tmpEstoque"
        
    ' MOVIMENTO
    ElseIf Me.cboConteudo.Column(0) = "MOVIMENTO" Then
    
        strSQL01 = "INSERT INTO tmpMovimento ( codFatura, codTipoMovimento, Controle, DescricaoDoMovimento, DataDeEmissao, DataDeVencimento, ValorDoMovimento,Categoria,Definicao,Especie,Status,Notas ) "
        strSQL02 = strSQL01 & " Select * from (SELECT * FROM " & Me.cboConteudo.Column(1) & " WHERE (((" & Me.cboConteudo.Column(2) & ") In (" & strDatas & ")))) as tmp"
           
        ExecutarSQL "Delete * from tmpMovimento"
        
        ExecutarSQL strSQL02
        
        ExportarXLS Me.cboConteudo.Column(0), "tmpMovimento"
    End If
    

End If

Set ctlOrigem = Nothing

End Sub

Private Sub cboConteudo_Click()
Dim strSQL As String

strSQL = "SELECT DISTINCT " & Me.cboConteudo.Column(2) & " FROM " & Me.cboConteudo.Column(1) & " ORDER BY " & Me.cboConteudo.Column(2) & " DESC"


Me.lstDados.RowSource = strSQL
Me.lstDados.Requery

End Sub

Private Sub cmdLimparSelecao_Click()
Dim I As Integer

    For I = 0 To Me.lstDados.ListCount
        Me.lstDados.Selected(I) = False
    Next I
    
End Sub

Private Sub cmdSelecionarTudo_Click()
Dim I As Integer

    For I = 0 To Me.lstDados.ListCount
        Me.lstDados.Selected(I) = True
    Next I
    
End Sub

Private Function ExportarXLS(strTitulo As String, strConsulta As String)
    Dim sTemp As String
    
    sTemp = pathDesktopAddress & Format(Now, "yymmdd_hhnn") & "-" & strTitulo & "_" & NomeDaLoja & ".xls"
    
    DoCmd.Hourglass True
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel97, strConsulta, sTemp, True
    
    DoCmd.Hourglass False
    MsgBox "A planilha foi gerada com êxito." & vbCrLf & vbCrLf & "Está em " & sTemp, vbInformation, "Exportação de dados"

End Function
