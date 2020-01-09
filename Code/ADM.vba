Option Compare Database
Option Explicit

Public strTabela As String

Public Function NovoCodigo(tabela, Campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & tabela & ";")
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

Public Function Pesquisar(tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = tabela
    
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


Function GerarParcelamento(codPedido As Long, dtEmissao As Date, ValParcelado As Currency, Parcelamento As String)

'Dim Valor As String
'Valor = "30"

Dim matriz As Variant
Dim x As Integer
Dim Parcelas As DAO.Recordset

Set Parcelas = CurrentDb.OpenRecordset("Select * from PedidosPagamentos")

matriz = Array()
matriz = Split(Parcelamento, ";")

BeginTrans

For x = 0 To UBound(matriz)
    Parcelas.AddNew
    Parcelas.Fields("codPedido") = codPedido
    Parcelas.Fields("Vencimento") = CalcularVencimento(dtEmissao, CInt(matriz(x)))
    Parcelas.Fields("Valor") = ValParcelado / (UBound(matriz) + 1)
    Parcelas.Update
Next

CommitTrans

Parcelas.Close

End Function

Public Function CalcularVencimento(dtInicio As Date, qtdDias As Integer, Optional ForaMes As Boolean) As Date

    If ForaMes Then
        CalcularVencimento = Format((DateSerial(Year(dtInicio), Month(dtInicio) + 1, qtdDias)), "dd/mm/yyyy")
    Else
        CalcularVencimento = Format((DateSerial(Year(dtInicio), Month(dtInicio), Day(dtInicio) + qtdDias)), "dd/mm/yyyy")
    End If

End Function

Public Function ExecutarSQL(strSQL As String)

    'Desabilitar menssagens de execução de comando do access
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL strSQL
    
    'Abilitar menssagens de execução de comando do access
    DoCmd.SetWarnings True

End Function
Function EnviarEmail(strEmail As String, strAssunto As String, strMensagem As String, strAnexo As String)
' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, Outlook 2010.
' This example sends the last saved version of the Activeworkbook object .
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
        .To = strEmail
        .CC = ""
        .BCC = ""
        .Subject = strAssunto 'ActiveSheet.Name
        .Body = strMensagem
'        .Attachments.Add ActiveWorkbook.FullName
        ' You can add other files by uncommenting the following line.
        .Attachments.Add (strAnexo)
        ' In place of the following statement, you can use ".Display" to
        ' display the mail.
        .Send
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Function
