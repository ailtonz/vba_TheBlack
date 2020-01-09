Option Compare Database
Option Explicit

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_admPesquisar.lstCadastro.Requery
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

Private Sub txtCategoria_Exit(Cancel As Integer)
    Me.txtCategoria = UCase(Me.txtCategoria)
End Sub


Private Sub Form_BeforeInsert(Cancel As Integer)
Dim strSQL As String
strSQL = "Select * from admCategorias"

    If Me.NewRecord Then Me.codigo = NovoCodigo(strSQL, Me.codigo.ControlSource)
End Sub

Private Sub Form_Load()
'    DoCmd.Maximize
End Sub


Private Sub txtCodigoExterno_Exit(Cancel As Integer)
    Me.txtCodigoExterno = UCase(Me.txtCodigoExterno)
End Sub
