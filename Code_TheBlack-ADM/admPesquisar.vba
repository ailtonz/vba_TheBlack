Option Compare Database
Option Explicit
Public StringF2 As String


Private Sub Form_Load()

    StringF2 = ""
    Filtro strTabela
    Me.KeyPreview = True
    Me.lstCadastro.SetFocus
    Me.lstCadastro.Selected(1) = True

End Sub

Private Sub cmdFiltrar_Click()

    Dim txtFiltro As String
    txtFiltro = InputBox("Digite uma palavra para fazer o filtro:", "Filtro", StringF2, 0, 0)
    StringF2 = txtFiltro
    Filtro strTabela, txtFiltro
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyInsert
        
           cmdNovo_Click
           
        Case vbKeyReturn
        
            cmdAlterar_Click
            
        Case vbKeyDelete
           
            cmdExcluir_Click
        
        Case vbKeyF2
        
            cmdFiltrar_Click
            
    End Select
End Sub

Private Sub cmdNovo_Click()

    Manipulacao strTabela, "Novo"
    
End Sub

Private Sub cmdAlterar_Click()

    Manipulacao strTabela, "Alterar"
    
End Sub

Private Sub cmdExcluir_Click()

    Manipulacao strTabela, "Excluir"
    
End Sub

Private Sub lstCadastro_DblClick(Cancel As Integer)

    cmdAlterar_Click
    
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

Public Sub Manipulacao(Tabela As String, Operacao As String)

If IsNull(Form_admPesquisar.lstCadastro.Value) And Operacao <> "Novo" Then
   Exit Sub
End If

Dim rstFormularios As DAO.Recordset

Set rstFormularios = _
    CurrentDb.OpenRecordset("Select * from Formularios  where codFormulario = " & Tabela)

Select Case Operacao

 Case "Novo"
        
    DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , , acFormAdd
    
 Case "Alterar"

    DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , rstFormularios.Fields("Identificacao") & " = " & Form_admPesquisar.lstCadastro.Value

 Case "Excluir"

    If MsgBox("Deseja excluir este registro?", vbInformation + vbOKCancel) = vbOK Then
       DoCmd.SetWarnings False
       DoCmd.RunSQL ("Delete from (" & rstFormularios.Fields("TabelaPrincipal") & ") as tmp  where " & rstFormularios.Fields("Identificacao") & " = " & Form_admPesquisar.lstCadastro.Value)
       DoCmd.SetWarnings True
    End If

End Select

Form_admPesquisar.lstCadastro.Requery

Saida:

End Sub

Private Function Filtro(strTabela As String, Optional Procurar As String)

Dim rstFormularios As DAO.Recordset
Dim rstForm_Campos As DAO.Recordset
Dim rstResultado As DAO.Recordset

Dim Sql As String
Dim SqlAux As String
Dim Contagem As Integer
Dim a, b, C As Integer
Dim Colunas As Integer

Dim Procuras(30) As String
Dim ProcurasAux As Integer
ProcurasAux = 1

For b = 1 To Len(Procurar)
   If Mid(Procurar, b, 1) = "+" Then
      ProcurasAux = ProcurasAux + 1
   Else
      Procuras(ProcurasAux) = Procuras(ProcurasAux) + Mid(Procurar, b, 1)
   End If
Next b

Set rstFormularios = _
    CurrentDb.OpenRecordset("Select * from Formularios where codFormulario = " & strTabela)
                                             
Set rstForm_Campos = _
    CurrentDb.OpenRecordset("Select * from Formularios_Campos where codFormulario = " & rstFormularios.Fields("codFormulario") & " Order by codOrdem")


Sql = "Select "

While Not rstForm_Campos.EOF
    If rstForm_Campos.Fields("Pesquisa") = True Then
        Sql = Sql & IIf(IsNull(rstForm_Campos.Fields("Nome")), _
                      rstForm_Campos.Fields("Campo"), _
                      rstForm_Campos.Fields("Campo") & _
                      " AS " & rstForm_Campos.Fields("Nome")) & ", "
    End If
    
    rstForm_Campos.MoveNext
Wend

Sql = Left(Sql, Len(Sql) - 2) & " "

Sql = Sql & " from "


If SqlAux = "" Then
   Sql = Sql & "(" & rstFormularios.Fields("TabelaPrincipal") & ") as tmp Where ( "
Else
   Sql = Sql & " Where ("
End If

rstForm_Campos.MoveFirst

For C = 1 To ProcurasAux
   
   rstForm_Campos.MoveFirst
   Sql = Sql & " ( "
   While Not rstForm_Campos.EOF
     If rstForm_Campos.Fields("Filtro") = True Then
        Sql = Sql & rstForm_Campos.Fields("Campo") & " Like '*" _
                  & LCase(Trim(Procuras(C))) & "*' OR "
     End If
     rstForm_Campos.MoveNext
   Wend
   Sql = Left(Sql, Len(Sql) - 3) & ") "
   If C <> ProcurasAux Then
      Sql = Sql + " And "
   End If
   
Next C

Sql = Sql + " ) "

Sql = Sql & "Order By "

rstForm_Campos.MoveFirst

While Not rstForm_Campos.EOF

  If rstForm_Campos.Fields("Ordem") <> "" Then
     Sql = Sql & rstForm_Campos.Fields("Campo") _
               & " " & rstForm_Campos.Fields("Ordem") & ", "
  End If
  
  rstForm_Campos.MoveNext
  
Wend

Sql = Left(Sql, Len(Sql) - 2) & " "

Sql = Sql & ";"

Me.lstCadastro.RowSource = Sql
Me.lstCadastro.ColumnHeads = True
Me.lstCadastro.ColumnCount = rstForm_Campos.RecordCount
Me.Caption = rstFormularios.Fields("TituloDoFormulario")

Dim strTamanho As String

rstForm_Campos.MoveFirst
While Not rstForm_Campos.EOF
  If Not IsNull(rstForm_Campos.Fields("Tamanho")) Then
     strTamanho = strTamanho & str(rstForm_Campos.Fields("Tamanho")) & "cm;"
  End If
  rstForm_Campos.MoveNext
Wend

Me.lstCadastro.ColumnWidths = strTamanho

If IsNull(rstFormularios.Fields("campodesoma")) = True Then
   
   Me.lblSoma.Caption = "Qtd: " & Me.lstCadastro.ListCount - 1
   
Else

   Dim Soma
   Set rstResultado = CurrentDb.OpenRecordset(Sql)
   Do While Not rstResultado.EOF
      Soma = Soma + rstResultado.Fields(rstFormularios.Fields("campodesoma"))
      rstResultado.MoveNext
   Loop
   Me.lblSoma.Caption = "Qtd: " & Me.lstCadastro.ListCount - 1 & "   Soma: " & FormatNumber(Soma, 2)
   rstResultado.Close
   
End If


rstFormularios.Close
rstForm_Campos.Close

Set rstFormularios = Nothing
Set rstForm_Campos = Nothing


End Function


