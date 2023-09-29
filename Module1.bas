Attribute VB_Name = "MsbBoxs"
Public Sub MensagemErro(Optional ByVal mensagem As String = "Mensagem de erro sem valor", Optional ByVal titulo As String = "Erro")
    MsgBox mensagem, vbExclamation, titulo
End Sub

Public Sub MensagemSucesso(ByVal mensagem As String, Optional ByVal titulo As String)
     MsgBox mensagem, vbInformation, titulo
End Sub
