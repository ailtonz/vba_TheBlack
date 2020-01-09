Option Compare Database
Option Explicit

'***********************************************
'Autor: Luiz Cláudio C. V. Rocha - luiz@mvps.org
'***********************************************

Sub AtivaBotao()
'Ativa ou desativa botão de atualizar vínculos
If Me.lstTabelas.ItemsSelected.Count > 0 Then
    Me.cmdAtualizar.Enabled = True
Else
    Me.cmdVerificar.SetFocus
    Me.cmdAtualizar.Enabled = False
End If
End Sub

Sub PreencheLista()
Dim cnn As ADODB.Connection
Dim cat As Object
Dim tbl As Object

Dim strLista As String 'origem de linha
Dim strSource As String 'path do BD
Dim strStatus As String 'OK ou !

On Error GoTo ErrHandler

'Banco de dados atual
Set cnn = CurrentProject.Connection
Set cat = CreateObject("ADOX.Catalog")
cat.ActiveConnection = cnn

'Percorre todas as tabelas
For Each tbl In cat.Tables
    'Se for vinculada, inclui na listbox
    If tbl.Type = "LINK" Then
        'Armazena o data source
        strSource = _
        tbl.Properties("Jet OLEDB:Link Datasource")
        'Verifica se o link está OK
        On Error Resume Next
        tbl.Properties("Jet OLEDB:Link Datasource") = _
        strSource
        'Verifica se ocorreu erro
        If Err = 0 Then
            'Se não ocorreu erro, OK
            strStatus = "OK"
        Else
            'Se ocorreu erro, exclamação "!"
            strStatus = "!"
            Err.Clear
        End If
        'Três colunas: status, nome da tabela, endereço
        strLista = strLista & strStatus & ";" & tbl.Name _
        & ";" & strSource & ";"
    End If
Next tbl

'Origem da listbox
Me.lstTabelas.RowSource = strLista

Call AtivaBotao

ExitHere:
'Libera a memória
Set tbl = Nothing
Set cat = Nothing
Set cnn = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description
Resume ExitHere
End Sub

Private Sub cmdAtualizar_Click()
'Atualiza o vínculo das tabelas
Dim fd As Office.FileDialog
Dim strArq As String
Dim varItem As Variant
Dim strTabela As String

Dim cnn As ADODB.Connection
Dim cat As Object
Dim tbl As Object

On Error GoTo ErrHandler

'Diálogo de selecionar arquivo - Office
Set fd = Application.FileDialog(msoFileDialogFilePicker)

fd.Filters.Add "BDs do Access", "*.MDB;*.MDE"
fd.Title = "Localize a fonte de dados"
fd.AllowMultiSelect = False
If fd.Show = -1 Then
    strArq = fd.SelectedItems(1)
End If
'Se selecionou arquivo, atualiza os vínculos
If strArq <> "" Then

    'Banco de dados atual
    Set cnn = CurrentProject.Connection
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = cnn

    'Percorre os itens da listbox
    For Each varItem In Me.lstTabelas.ItemsSelected
        strTabela = Me.lstTabelas.Column(1, varItem)
        On Error Resume Next
        'Define o novo vínculo
        Set tbl = cat.Tables(strTabela)
        tbl.Properties("Jet OLEDB:Link Datasource") = _
        strArq
        'Se houver erro, avisa
        If Not Err = 0 Then
            MsgBox "Erro ao vincular " & tbl.Name
            Err.Clear
        End If
    Next varItem
    'Atualiza a listbox
    Call PreencheLista
End If

ExitHere:
'Libera a memória
Set tbl = Nothing
Set cat = Nothing
Set cnn = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description
Resume ExitHere
End Sub

Private Sub cmdLimpar_Click()
'Limpa a seleção
Dim I As Integer

    For I = 0 To Me.lstTabelas.ListCount
        Me.lstTabelas.Selected(I) = False
    Next I
    Call AtivaBotao

End Sub

Private Sub cmdSelecionar_Click()
'Seleciona todos os itens da listbox
Dim I As Integer

    For I = 0 To Me.lstTabelas.ListCount
        Me.lstTabelas.Selected(I) = True
    Next I
    Call AtivaBotao

End Sub

Private Sub cmdVerificar_Click()
    Call PreencheLista
End Sub

Private Sub Form_Open(Cancel As Integer)
    Call PreencheLista
End Sub

Private Sub lstTabelas_Click()
    Call AtivaBotao
End Sub

