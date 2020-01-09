Option Compare Database
Option Explicit

'Private Sub cmdParcelamento_Click()
'
'    If Not IsNull(Me.Pagamento.Column(1)) Then
'        GerarParcelamento Me.codPedido, Format(Me.Emissao, "dd/mm/yy"), Me.TotalGeral, Me.Pagamento.Column(1)
'        PedidosPagamentos.Requery
'    End If
'
'End Sub
'
'Private Sub Descricao_Click()
'    Me.codcli = Me.Descricao.Column(1)
'End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    
    If Me.NewRecord Then
        Me.codigo = NovoCodigo(Me.RecordSource, Me.codigo.ControlSource)
        Me.Emissao = Now
    End If
    
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_Pesquisar.lstCadastro.Requery
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

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String

    stDocName = "Pedidos"
    DoCmd.OpenReport stDocName, acPreview, , "codPedido = " & Me.codigo

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
    
End Sub

'Private Sub TipoDePedido_Click()
'    Me.TipoDeMovimento = Me.TipoDePedido.Column(1)
'End Sub

'Private Sub cmdBaixarEstoque_Click()
'    Call MovimentoDeEstoque
'End Sub

'Private Sub MovimentoDeEstoque()
'
'If Me.MovimentarEstoque = True Then
'
'    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
'
'    Dim rItens As DAO.Recordset
'    Dim rProdutos As DAO.Recordset
'
'    Set rItens = CurrentDb.OpenRecordset("Select * from PedidosItens where codPedido = " & Me.codigo)
'    Set rProdutos = CurrentDb.OpenRecordset("Select * from Produtos")
'
'
'    While Not rItens.EOF
'        Dim strcodProduto As String
'        strcodProduto = rItens.Fields("codProduto")
'        rProdutos.FindFirst "codProduto = " & strcodProduto & ""
'
'        If Me.TipoDeMovimento = "Entrada" Then
'            If rItens.Fields("Selecao") = "S/NF" Then
'                rProdutos.Edit
'                rProdutos.Fields("QTD_S_NF") = rProdutos.Fields("QTD_S_NF") + rItens.Fields("Quantidade")
'                rProdutos.Update
'            Else
'                rProdutos.Edit
'                rProdutos.Fields("QTD_C_NF") = rProdutos.Fields("QTD_C_NF") + rItens.Fields("Quantidade")
'                rProdutos.Update
'            End If
'        Else
'            If rItens.Fields("Selecao") = "S/NF" Then
'                rProdutos.Edit
'                rProdutos.Fields("QTD_S_NF") = rProdutos.Fields("QTD_S_NF") - rItens.Fields("Quantidade")
'                rProdutos.Update
'            Else
'                rProdutos.Edit
'                rProdutos.Fields("QTD_C_NF") = rProdutos.Fields("QTD_C_NF") - rItens.Fields("Quantidade")
'                rProdutos.Update
'            End If
'
'        End If
'
'        rItens.MoveNext
'
'    Wend
'
'    rItens.Close
'    rProdutos.Close
'    Me.MovimentarEstoque = False
'    MsgBox "Estoque baixado com sucesso!", vbInformation + vbOKOnly
'
'End If
'
'End Sub
