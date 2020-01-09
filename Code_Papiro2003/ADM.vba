Option Compare Database
Option Explicit

Public strTabela As String

Public Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")
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

Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
    
    If strTabela = "Cheques" Then Baixadecheques
           
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function

Sub Baixadecheques()

Dim Soma As Integer
Dim rs As DAO.Recordset

Set rs = CurrentDb.OpenRecordset("cheques")

If Not rs.EOF Then

   Do While Not rs.EOF
      
      If rs.Fields("datacompensar") <= Date And rs.Fields("compensou") = "Não" And rs.Fields("deubaixa") = False Then
      
         rs.Edit
         rs.Fields("compensou") = "Sim"
         rs.Fields("DeuBaixa") = True
         rs.Update
         Soma = Soma + 1
         
      End If
            
      rs.MoveNext
   Loop

End If

rs.Close

If Soma <> 0 Then MsgBox "Foram compensado(s) " & Soma & " cheque(s)!"

End Sub

