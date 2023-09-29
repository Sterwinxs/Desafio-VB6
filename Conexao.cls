VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conn As ADODB.Connection
Private cmd As ADODB.Command
Private rs As ADODB.Recordset ' Adicione um campo para o objeto Recordset

Public Sub Conectar()
   On Error GoTo TratarErro
   Set conn = New ADODB.Connection ' Criando um objeto de conexão ADODB
   
   conn.ConnectionString = "DSN=PostgreSQL;" ' Definindo a string de conexão
   conn.Open ' Abrindo a conexão com o banco de dados
   Exit Sub

TratarErro:
    MsgBox "Erro ao conectar ao banco de dados: " & Err.Description
End Sub

Public Sub ExecutarSQL(ByVal sql As String)
   On Error GoTo TratarErro
   
   Set cmd = New ADODB.Command ' Criando um novo objeto de comando ADODB
   
   With cmd
      .ActiveConnection = conn ' Define a conexão
      .CommandText = sql  ' Atribuindo o comando com a variável sql
      .CommandType = adCmdText ' Especifica que é um comando de texto SQL
      .Execute ' Executa o comando SQL
   End With
   
   Exit Sub

TratarErro:
    ' Lidando com erros - exibindo uma mensagem de erro em uma caixa de diálogo
    MsgBox "Erro ao executar consulta SQL: " & Err.Description
End Sub
Public Function GetRecordset(ByVal sql As String) As ADODB.Recordset
   On Error GoTo TratarErro
   Set result = New ADODB.Recordset ' Cria um novo objeto de Recordset ADODB
   
   result.Open sql, conn ' Abre o Recordset usando a consulta SQL e a conexão 'conn'
   Set GetRecordset = result ' Retornando o Recordset resultante
   Exit Function

TratarErro:
   MsgBox "Erro ao obter o Recordset: " & Err.Description
   Set GetRecordset = Nothing 'Retornar Nothing (nenhum Recordset válido)
End Function

Public Sub Desconectar()
   On Error Resume Next
   conn.Close ' Fechar a conexão
   
   'Liberando recursos
   Set conn = Nothing
   Set cmd = Nothing
   Set rs = Nothing
   
   On Error GoTo 0
End Sub

