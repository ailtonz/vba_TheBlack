Option Compare Database
Option Explicit

Private Sub cmdCancelar_Click()
On Error GoTo Err_cmdCancelar_Click

    DoCmd.Close

Exit_cmdCancelar_Click:
    Exit Sub

Err_cmdCancelar_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancelar_Click
End Sub

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

Dim stDocName As String
Dim stLinkCriteria As String
Dim Campo As String
Dim Filtro As Boolean

Dim RelatoriosCriterios As DAO.Recordset
Dim RelatoriosPeriodos As DAO.Recordset

If IsNull(Me.lstRelatorios.Value) Then
    MsgBox "Por favor selecione um dos relatórios ao lado!", vbCritical
    Me.lstRelatorios.SetFocus
    GoTo Exit_cmdVisualizar_Click
End If

stDocName = Me.lstRelatorios.Column(1)

Set RelatoriosCriterios = CurrentDb.OpenRecordset _
        ("Select * from RelatoriosCriterios where " & _
        "Criterio = True and codRelatorio = " & _
        Me.lstRelatorios.Column(2))

If Not RelatoriosCriterios.EOF Then

    stLinkCriteria = ""
    
    RelatoriosCriterios.MoveFirst
    
    While Not RelatoriosCriterios.EOF
    
         If RelatoriosCriterios.Fields![TipoCriterio] = 3 Then    ' Texto
            
            If RelatoriosCriterios.Fields("Valor") > 0 Then
               stLinkCriteria = stLinkCriteria & RelatoriosCriterios.Fields("Campo") & " = '" & RelatoriosCriterios.Fields("Valor") & "' and "
            End If
         
         ElseIf RelatoriosCriterios.Fields![TipoCriterio] = 2 Then    ' Valor
        
            If RelatoriosCriterios.Fields("Valor") > 0 Then
               stLinkCriteria = stLinkCriteria & RelatoriosCriterios.Fields("Campo") & " = " & RelatoriosCriterios.Fields("Valor") & " and "
            End If
         
         End If
    
       RelatoriosCriterios.MoveNext

    Wend
    
    If stLinkCriteria <> "" Then
        stLinkCriteria = Left(stLinkCriteria, Len(stLinkCriteria) - 5)
    End If

End If


Set RelatoriosPeriodos = CurrentDb.OpenRecordset _
        ("Select * from RelatoriosCriterios where " & _
        "Criterio = False and codRelatorio = " & _
        Me.lstRelatorios.Column(2))

If Not RelatoriosPeriodos.EOF Then

    RelatoriosPeriodos.MoveFirst
    While Not RelatoriosPeriodos.EOF
        
        If RelatoriosPeriodos.Fields![TipoCriterio] = 1 Then    ' Datas
        
            If Not IsNull(RelatoriosPeriodos.Fields![Inicio]) Then
                stLinkCriteria = stLinkCriteria & IIf(stLinkCriteria <> "", " and ", "") & _
                                "[" & RelatoriosPeriodos.Fields![Campo] & "] Between #" & _
                                Format(RelatoriosPeriodos.Fields![Inicio], "mm/dd/yyyy") & _
                                "# AND #" & _
                                Format(RelatoriosPeriodos.Fields![Terminio], "mm/dd/yyyy") & "#"
            End If
        
        
        ElseIf RelatoriosPeriodos.Fields![TipoCriterio] = 2 Then    ' Valor
        
            If Not IsNull(RelatoriosPeriodos.Fields![Inicio]) Then
                stLinkCriteria = stLinkCriteria & IIf(stLinkCriteria <> "", " and ", "") & _
                                "[" & RelatoriosPeriodos.Fields![Campo] & "] Between " & _
                                RelatoriosPeriodos.Fields![Inicio] & _
                                " AND " & _
                                RelatoriosPeriodos.Fields![Terminio] & ""
            End If
            
        ElseIf RelatoriosPeriodos.Fields![TipoCriterio] = 3 Then    ' Texto
        
            If Not IsNull(RelatoriosPeriodos.Fields![Inicio]) Then
                stLinkCriteria = stLinkCriteria & IIf(stLinkCriteria <> "", " and ", "") & _
                                "[" & RelatoriosPeriodos.Fields![Campo] & "] Between '" & _
                                RelatoriosPeriodos.Fields![Inicio] & _
                                "' AND '" & _
                                RelatoriosPeriodos.Fields![Terminio] & "'"
            End If
            
        End If
    
        RelatoriosPeriodos.MoveNext
    Wend

End If

RelatoriosCriterios.Close
RelatoriosPeriodos.Close

DoCmd.OpenReport stDocName, acPreview, , stLinkCriteria


Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click

End Sub

Private Sub Form_Load()
    Me.lstRelatorios.Selected(0) = True
End Sub

Private Sub lstRelatorios_Click()

Dim SQL_Limpar As String

SQL_Limpar = "UPDATE RelatoriosCriterios SET " & _
             "Descricao = Null," & _
             "Inicio = Null, " & _
             "Terminio = Null," & _
             "RelatoriosCriterios.Valor = 0"

DoCmd.SetWarnings False
DoCmd.RunSQL SQL_Limpar
DoCmd.SetWarnings True

Me.Filter = "codRelatorio = " & Me.lstRelatorios.Column(2)
Me.FilterOn = True

End Sub

Private Sub lstRelatorios_DblClick(Cancel As Integer)

    Call cmdVisualizar_Click
 
End Sub

