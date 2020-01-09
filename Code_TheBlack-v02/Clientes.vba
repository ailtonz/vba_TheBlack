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

Private Sub Cliente_Exit(Cancel As Integer)
    Me.Cliente = UCase(Me.Cliente)
End Sub



Private Sub Email_Exit(Cancel As Integer)
'    Me.Email = LCase(Me.Email)
End Sub

Private Sub Endereco_Exit(Cancel As Integer)
    Me.Endereco = UCase(Me.Endereco)
End Sub

Private Sub Bairro_Exit(Cancel As Integer)
    Me.Bairro = UCase(Me.Bairro)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
Dim strSQL As String
strSQL = "Select * from Cadastros"
    If Me.NewRecord Then Me.Codigo = NovoCodigo(strSQL, Me.Codigo.ControlSource)
End Sub

Private Sub Form_Load()
'    DoCmd.Maximize
End Sub

Private Sub Municipio_Exit(Cancel As Integer)
    Me.Municipio = UCase(Me.Municipio)
End Sub

Private Sub Estado_Exit(Cancel As Integer)
    Me.Estado = UCase(Me.Estado)
End Sub

Private Sub OBS_Exit(Cancel As Integer)
'    Me.OBS = UCase(Me.OBS)
End Sub
