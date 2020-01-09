Option Compare Database

Private Sub cmdArquivo_Click()
Dim ssql As String

If Not IsNull(Me.cboANO) Then

    ssql = "UPDATE Cheques SET Cheques.Arquivo = Yes WHERE (((Format([DataCompensar],'yy'))=" & Me.cboANO.Column(0) & "))"
    
    ExecutarSQL ssql
    
    Me.lstCheques.Requery

End If

End Sub

Private Sub cmdRetirar_Click()
Dim ssql As String

If Not IsNull(Me.cboANO) Then

    ssql = "UPDATE Cheques SET Cheques.Arquivo = no WHERE (((Format([DataCompensar],'yy'))=" & Me.cboANO.Column(0) & "))"
    
    ExecutarSQL ssql
    
    Me.lstCheques.Requery

End If

End Sub
