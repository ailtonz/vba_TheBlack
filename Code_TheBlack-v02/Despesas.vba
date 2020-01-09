Option Compare Database
Option Explicit
Private Function CalcularVencimento(Dia As Integer, Optional MES As Integer, Optional Ano As Integer) As Date

If Month(Now) = 2 Then
    If Dia = 29 Or Dia = 30 Or Dia = 31 Then
        Dia = 1
        MES = MES + 1
    End If
End If

If MES > 0 And Ano > 0 Then
    CalcularVencimento = Format((DateSerial(Ano, MES, Dia)), "dd/mm/yyyy")
ElseIf MES = 0 And Ano > 0 Then
    CalcularVencimento = Format((DateSerial(Ano, Month(Now), Dia)), "dd/mm/yyyy")
ElseIf MES = 0 And Ano = 0 Then
    CalcularVencimento = Format((DateSerial(Year(Now), Month(Now), Dia)), "dd/mm/yyyy")
End If

End Function
Private Function EstaAberto(strName As String) As Boolean
On Error GoTo EstaAberto_Err
' Testa se o formulário está aberto

   Dim obj As AccessObject, dbs As Object
   Set dbs = Application.CurrentProject
   ' Procurar objetos AccessObject abertos na coleção AllForms.
   
   EstaAberto = False
   For Each obj In dbs.AllForms
        If obj.IsLoaded = True And obj.Name = strName Then
            ' Imprimir nome do obj.
            EstaAberto = True
            Exit For
        End If
   Next obj
    
EstaAberto_Fim:
  Exit Function
EstaAberto_Err:
  Resume EstaAberto_Fim
End Function

Private Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")
If Not rstTabela.EOF Then
   NovoCodigo = rstTabela.Fields("CodigoNovo")
   If IsNull(NovoCodigo) Then
      NovoCodigo = 1
   End If
Else
   NovoCodigo = 1
End If
rstTabela.Close

End Function

Private Sub codCategoria_Exit(Cancel As Integer)
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    
    If Me.NewRecord Then
       Me.DataDeEmissao = Format(Now(), "dd/mm/yy")
       Me.codTipoMovimento = "Despesa"
       Me.Status = "Aberto"
    End If
    
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    If EstaAberto("admPesquisar") Then Form_admPesquisar.lstCadastro.Requery
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

Private Sub codCategoria_Click()
Dim SQL_Definicoes As String
Dim strCategoria As String

strCategoria = IIf(IsNull(Me.codCategoria.Column(0)), "", Me.codCategoria.Column(0))

SQL_Definicoes = "Select Distinct * from qryDefinicaoDespesas where Categoria = '" & Forms!Despesas.codCategoria.Column(0) & "'"

Me.codDefinicao.RowSource = SQL_Definicoes
        
End Sub

Private Sub codDefinicao_GotFocus()
    codCategoria_Click
End Sub

Private Sub cmdRepetir_Click()
Dim DB As DAO.Database
Dim rst As DAO.Recordset
Dim x As Integer
Dim VCTO As Integer
Dim MES As Integer


VCTO = Format(DataDeVencimento, "dd")
MES = Format(DataDeVencimento, "mm") + 1

'Salvar Registro
DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Set DB = CurrentDb()
'Abre a tabela, adiciona o novo item e atualiza a combo
Set rst = DB.OpenRecordset("Movimentos")
    
Dim Message, Title, Default, MyValue

Message = "Quantas vezes este cadastro deve repetir? "    ' Define o aviso.
Title = "Repetir cadastro"       ' Define o título.
Default = "1"    ' Define o padrão.

msgRepetirCadastro:
' Exibe a mensagem, o título e o valor padrão.
MyValue = InputBox(Message, Title, Default)

' Cancelar Processo
If MyValue = "" Then GoTo sair

' Verificar Integridade de informação (É numero?)
If Not IsNumeric(MyValue) Then GoTo msgRepetirCadastro
    
Me.Controle = "(" & 1 & "/" & MyValue & ")"

'Salvar Registro
DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

With rst
    
    For x = 2 To MyValue
        .AddNew
        !codMovimento = NovoCodigo("Movimentos", "codMovimento")
        !DataDeEmissao = Format(Now(), "dd/mm/yy")
        !DataDeVencimento = CalcularVencimento(VCTO, MES, Year(Now))
        !codTipoMovimento = Me.codTipoMovimento
        !DescricaoDoMovimento = Me.DescricaoDoMovimento
        !Categoria = Me.codCategoria
        !Definicao = Me.codDefinicao
        !Controle = "(" & x & "/" & MyValue & ")"
        !ValorDoMovimento = Me.ValorDoMovimento
        !Especie = Me.Especie
        !Status = Me.Status
        !codRelacao = Me.Codigo
        .Update
        MES = MES + 1
    Next x
    
End With

MsgBox "Operação realizada com sucesso!", vbOKOnly + vbInformation, Title

sair:

rst.Close
DB.Close

Set rst = Nothing
Set DB = Nothing

End Sub
