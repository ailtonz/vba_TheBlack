Option Compare Database
Option Explicit

Public strTabela As String

Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "admPesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function

Public Function RedimencionaControle(frm As Form, ctl As Control)

Dim intAjuste As Integer
On Error Resume Next

intAjuste = frm.Section(acHeader).Height * frm.Section(acHeader).Visible

intAjuste = intAjuste + frm.Section(acFooter).Height * frm.Section(acFooter).Visible

On Error GoTo 0

intAjuste = Abs(intAjuste) + ctl.top

If intAjuste < frm.InsideHeight Then
    ctl.Height = frm.InsideHeight - intAjuste
'    ctl.Width = frm.InsideHeight + (intAjuste + intAjuste)
End If

End Function

Public Function EstaAberto(strName As String) As Boolean
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

Public Function IsFormView(frm As Form) As Boolean
On Error GoTo IsFormView_Err
' Testa se o formulário está aberto em
' modo formulário (form view)

 IsFormView = False
 If frm.CurrentView = 1 Then
    IsFormView = True
 End If

IsFormView_Fim:
  Exit Function
IsFormView_Err:
  Resume IsFormView_Fim
End Function

Public Function pathDesktopAddress() As String
    pathDesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

Public Function ExecutarSQL(strSQL As String, Optional log As Boolean)
'Objetivo: Executar comandos SQL sem mostrar msg's do access.

'Desabilitar menssagens de execução de comando do access
DoCmd.SetWarnings False

If log Then SaidaLog strSQL, "sql.log"

BeginTrans

'Executar a instrução SQL
DoCmd.RunSQL strSQL

CommitTrans

'Abilitar menssagens de execução de comando do access
DoCmd.SetWarnings True

End Function

Public Function SaidaLog(strConteudo As String, Optional strArquivo As String)
On Error GoTo SaidaLog_Err

If strArquivo = "" Then strArquivo = "sql.log"

Open Application.CurrentProject.Path & "\" & strArquivo For Append As #1

Print #1, strConteudo

Close #1

SaidaLog_Fim:
  Exit Function
SaidaLog_Err:
  Resume SaidaLog_Fim

End Function

Public Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset
Dim strSQL As String
Dim strSQL2 As String

strSQL2 = Left(Tabela, Len(Tabela))

strSQL = "SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM (" & strSQL2 & ") as tmp"
'SaidaLog strSQL, "sql.log"

Set rstTabela = CurrentDb.OpenRecordset(strSQL)

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

Public Function CalcularVencimento(Dia As Integer, Optional MES As Integer, Optional Ano As Integer) As Date

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

Public Function CalcularVencimento2(dtInicio As Date, qtdDias As Integer, Optional ForaMes As Boolean) As Date

Dim MyDate

    If ForaMes Then
        MyDate = Format((DateSerial(Year(dtInicio), Month(dtInicio) + 1, qtdDias)), "dd/mm/yyyy")
        CalcularVencimento2 = MyDate
    Else
        MyDate = Format((DateSerial(Year(dtInicio), Month(dtInicio), Day(dtInicio) + qtdDias)), "dd/mm/yyyy")
        
        If Weekday(MyDate) = 1 Then ' Domingo
            CalcularVencimento2 = Format((DateSerial(Year(dtInicio), Month(dtInicio), Day(dtInicio) + qtdDias + 1)), "dd/mm/yyyy")
        ElseIf Weekday(MyDate) = 7 Then ' Sabado
            CalcularVencimento2 = Format((DateSerial(Year(dtInicio), Month(dtInicio), Day(dtInicio) + qtdDias + 2)), "dd/mm/yyyy")
        Else 'Dia da semana
            CalcularVencimento2 = MyDate
        End If
        
    End If

End Function




Function ImportarXLS(strTabela As String, strTitulo As String)
    Dim strArquivo As Variant

    strArquivo = GetOpenFile(, strTitulo)
    
    If Len(strArquivo) > 0 Then
                    
            DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8, strTabela, strArquivo, True
    
    End If
    
End Function


 Function Work_Days(BegDate As Variant, EndDate As Variant) As Integer
   ' Note that this function does not account for holidays.
      Dim WholeWeeks As Variant
      Dim DateCnt As Variant
      Dim EndDays As Integer

      BegDate = DateValue(BegDate)
      EndDate = DateValue(EndDate)
      WholeWeeks = DateDiff("w", BegDate, EndDate)
      DateCnt = DateAdd("ww", WholeWeeks, BegDate)
      EndDays = 0
      Do While DateCnt < EndDate
         If Format(DateCnt, "ddd") <> "Sun" And _
                          Format(DateCnt, "ddd") <> "Sat" Then
            EndDays = EndDays + 1
         End If
         DateCnt = DateAdd("d", 1, DateCnt)
      Loop
      Work_Days = WholeWeeks * 5 + EndDays
   End Function

