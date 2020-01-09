Option Compare Database

Public Function NomeDaLoja() As String
Dim strSQL As String
Dim rstLoja As DAO.Recordset
strSQL = "SELECT admCategorias.Descricao04 FROM admCategorias WHERE (((admCategorias.codRelacao)=(SELECT admCategorias.codCategoria FROM admCategorias Where Categoria = 'Empresa')))"

Set rstLoja = CurrentDb.OpenRecordset(strSQL)

NomeDaLoja = rstLoja.Fields("Descricao04")

rstLoja.Close

Set rstLoja = Nothing

End Function

Sub CadastrarProdutos()
Dim rstEstoque As DAO.Recordset
Dim codCategorias As Integer
Dim codRelacao As Integer

Dim strSQL As String
strSQL = "Select * from admCategorias"


Set rstEstoque = CurrentDb.OpenRecordset("Select * from Estoque order by codLancamento")

While Not rstEstoque.EOF
On Error Resume Next

    codCategorias = NovoCodigo(strSQL, "codCategoria")
    codRelacao = 891

    ExecutarSQL "Insert into admCategorias (codCategoria,codRelacao,Categoria) Values (" & codCategorias & "," & codRelacao & ",'" & rstEstoque.Fields("DescricaoDoProduto") & "')", True

    rstEstoque.MoveNext

Wend


End Sub



Function LancarMovimento(codFatura As Long, _
                            dtEmissao As Date, _
                            ValorRecebido As Currency, _
                            Parcelamento As String, _
                            strEspecie As String, _
                            strCliente As String, _
                            strTipoMovimento As String, _
                            strCategoria As String, _
                            strDefinicao As String, _
                            strNotas As String)

Dim matriz As Variant
Dim x As Integer
Dim y As Integer
Dim Parcelas As DAO.Recordset

Set Parcelas = CurrentDb.OpenRecordset("Select * from Movimentos")

matriz = Array()
matriz = Split(Parcelamento, ";")

BeginTrans

For x = 0 To UBound(matriz)
    y = x + 1
    Parcelas.AddNew
    Parcelas.Fields("codFatura") = codFatura
    Parcelas.Fields("codTipoMovimento") = strTipoMovimento
    Parcelas.Fields("DataDeEmissao") = dtEmissao
    Parcelas.Fields("DescricaoDoMovimento") = codFatura & " - " & strCliente
    Parcelas.Fields("DataDeVencimento") = CalcularVencimento2(dtEmissao, CInt(matriz(x)))
    Parcelas.Fields("ValorDoMovimento") = ValorRecebido / (UBound(matriz) + 1)
    Parcelas.Fields("Controle") = "" & y & "/" & (UBound(matriz) + 1) & ""

    Parcelas.Fields("Categoria") = strCategoria
    Parcelas.Fields("Definicao") = strDefinicao
    Parcelas.Fields("Especie") = strEspecie
    Parcelas.Fields("Status") = "Aberto"
    Parcelas.Fields("Notas") = strNotas

    Parcelas.Update
Next

CommitTrans

Parcelas.Close

End Function

Public Function CadastrosDeTestes()
Dim sqlFaturas As String
Dim sqlItens As String


sqlFaturas = "INSERT INTO Faturas ( codFatura, Status, DataDeEmissao, codCadastro, Nome, codVeiculo )" & _
            "SELECT NovoCodigo('Select * from Faturas','codFatura') AS codFatura, 'Fatura' AS Status, Format(Now(),'dd/mm/yy') AS Emissao, Cadastros.codCadastro, Cadastros.Nome, First(Veiculos.codVeiculo) AS PrimeiroDecodVeiculo " & _
            "FROM Cadastros INNER JOIN Veiculos ON Cadastros.codCadastro = Veiculos.codCadastro " & _
            "GROUP BY NovoCodigo('Select * from Faturas','codFatura'), 'Fatura', Format(Now(),'dd/mm/yy'), Cadastros.codCadastro, Cadastros.Nome " & _
            "HAVING (((Cadastros.codCadastro) Not In (Select codCadastro from Faturas))) " & _
            "ORDER BY Cadastros.codCadastro "


ExecutarSQL sqlFaturas



sqlItens = "INSERT INTO FaturasItens ( codFatura, Referencia, Quantidade )" & _
            "SELECT Faturas.codFatura, 'PEÇAS' AS Referencia, 3 AS qtd " & _
            "FROM Faturas WHERE (((Faturas.codFatura) Not In (Select codFatura from FaturasItens)))"

ExecutarSQL sqlItens


End Function


Public Function CadastroMovimentos()
Dim rstEspecies As DAO.Recordset

Set rstEspecies = CurrentDb.OpenRecordset("Select * from tmpParcelamento")

While Not rstEspecies.EOF

LancarMovimento rstEspecies.Fields("codFatura"), _
                rstEspecies.Fields("dtEmissao"), _
                rstEspecies.Fields("ValorRecebido"), _
                rstEspecies.Fields("Parcelamento"), _
                rstEspecies.Fields("Especie"), _
                rstEspecies.Fields("Cliente"), _
                rstEspecies.Fields("TipoMovimento"), _
                rstEspecies.Fields("Categoria"), _
                rstEspecies.Fields("Definicao"), _
                rstEspecies.Fields("Notas")

rstEspecies.MoveNext

Wend

End Function
