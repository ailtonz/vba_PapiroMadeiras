Option Compare Database


Private Sub Produto_Click()
    Me.TipoDeMadeira = Me.Produto.Column(1)
    Me.codProduto = Me.Produto.Column(2)
End Sub

Private Sub Quantidade_Exit(Cancel As Integer)
    Call CalcularItem
End Sub
Private Sub Valor_Exit(Cancel As Integer)
    Call CalcularItem
End Sub

Private Sub CalcularItem()

    If Not IsNull(Me.Quantidade) Then
        If Not IsNull(Me.Valor) Then
            Me.Total = Me.Quantidade * Me.Valor
        End If
    End If
        

End Sub


