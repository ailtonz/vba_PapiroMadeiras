Option Compare Database

Private Sub Cheque_Click()
    Me.Emitente = Me.Cheque.Column(1)
    Me.Valor = Me.Cheque.Column(2)
    Me.DataCompensar = Me.Cheque.Column(3)
    Me.codCheque = Me.Cheque.Column(4)
End Sub

Private Sub Emitente_Click()
    Me.Cheque = Me.Emitente.Column(1)
    Me.Valor = Me.Cheque.Column(2)
    Me.DataCompensar = Me.Cheque.Column(3)
    Me.codCheque = Me.Cheque.Column(4)
End Sub
