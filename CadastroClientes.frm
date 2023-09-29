VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPesquisa 
   Caption         =   "Pesquisar e Editar Clientes"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   8319.761
   ScaleMode       =   0  'User
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   5685
      Picture         =   "CadastroClientes.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   630
      Width           =   495
   End
   Begin VB.TextBox txtBoxPesquisa 
      Height          =   345
      Left            =   2445
      TabIndex        =   6
      Top             =   780
      Width           =   3195
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   6525
      Top             =   495
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=PostgreSQL;"
      OLEDBString     =   "DSN=PostgreSQL;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGridVisualizar 
      Height          =   3825
      Left            =   210
      TabIndex        =   5
      Top             =   1170
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   6747
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton btnCadastrarClientes 
      Caption         =   "Cadastrar clientes"
      Height          =   375
      Left            =   135
      TabIndex        =   4
      Top             =   60
      Width           =   1665
   End
   Begin VB.CommandButton btnExcluir 
      Caption         =   "Excluir"
      Height          =   465
      Left            =   6615
      TabIndex        =   2
      Top             =   5220
      Width           =   1095
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "Editar"
      Height          =   465
      Left            =   5325
      TabIndex        =   1
      Top             =   5235
      Width           =   1095
   End
   Begin VB.CommandButton btnVisualizar 
      Caption         =   "Visualizar Lista"
      Height          =   420
      Left            =   345
      TabIndex        =   0
      Top             =   645
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   3615
      Left            =   2055
      TabIndex        =   8
      Top             =   1215
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTituloTela 
      Caption         =   "Pesquisar Clientes"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3195
      TabIndex        =   3
      Top             =   195
      Width           =   2865
   End
End
Attribute VB_Name = "FrmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEditar_Click()
 EditarRegistros
End Sub

Private Sub btnVisualizar_Click()
   CarregarLista
End Sub

Private Sub btnCadastrarClientes_Click()
   FrmCadastro.Show ' Abrir FrmCadastro
   Unload Me ' Fechar este form
End Sub


Private Sub FlexGridVisualizar_Click()

   Dim nome As String
   Dim sobrenome As String
   Dim sql As String
   Dim clickedRow As Integer
   Dim clickedCol As Integer
   
   Set DataGrid.DataSource = Nothing ' Limpando DataGrid
   
   ' Obtendo a linha e a coluna onde ocorreu o clique no FlexGridVisualizar
   clickedRow = FlexGridVisualizar.Row
   clickedCol = FlexGridVisualizar.Col
   
   ' Verificando se o clique ocorreu em uma célula válida (não no cabeçalho)
   If clickedRow > 0 Then
      ' Atribuindo o valor da célula clicada no FlexGridVisualizar
      Dim cellValue As String
      cellValue = FlexGridVisualizar.TextMatrix(clickedRow, clickedCol)
      
      ' Separando o nome completo em nome e sobrenome
      Dim separar() As String
      separar = Split(cellValue, " ")
      
      If UBound(separar) >= 1 Then
         nome = separar(0) ' O primeiro elemento é o nome
         sobrenome = separar(UBound(separar)) ' O último elemento é o sobrenome
      Else
         nome = cellValue ' Se não houver espaço, utiliza apenas o primeiro nome
         sobrenome = ""
      End If
      
      ' Atribuindo consulta a variavel sql
      sql = "SELECT * FROM contato WHERE nome = '" & nome & "' AND sobrenome = '" & sobrenome & "';"
      
      ' Atribuindo select
      Adodc.RecordSource = sql
      
      ' Executando consulta e atualizando dados
      Adodc.Refresh
      
      ' Preenchendo DataGrid
      Set DataGrid.DataSource = Adodc
   Else
        MsgBox "Clique em uma célula válida."
   End If
   
End Sub

Private Sub excluirContato(ByVal id As String)

   On Error GoTo TratarErro
   
   Dim conn As New Conexao
   conn.Conectar ' Conectando ao banco de dados

   Dim sql As String ' Atribuindo script
   sql = "DELETE FROM contato WHERE id = " & id & ";"
   conn.ExecutarSQL sql ' Executando sql

   conn.Desconectar ' Fechando conexão
   Exit Sub

TratarErro:
    MsgBox "Erro ao excluir contato: " & Err.Description
    RegistrarErro Err.Description
End Sub

Private Sub CarregarLista()

    On Error GoTo TratarErro

    Dim conn As New Conexao
    
    conn.Conectar ' Conectando ao banco de dados

    ' Executar a consulta SQL e obter o resultado (Recordset)
    Dim sql As String
    sql = "SELECT nome, sobrenome FROM contato;"
    Dim result As ADODB.Recordset
    Set result = conn.GetRecordset(sql)

    ' Preenchendo o FlexGridVisualizar com os resultados da consulta
    FlexGridVisualizar.Rows = 1
    FlexGridVisualizar.Cols = 1
    FlexGridVisualizar.FixedCols = 0 ' Remove cabeçalho vertical
    FlexGridVisualizar.FixedRows = 0 ' Remove cabeçalho horizontal
    FlexGridVisualizar.SelectionMode = flexSelectionByCell ' Habilitar seleção de células

    FlexGridVisualizar.TextMatrix(0, 0) = "Nome Completo" ' Adiciona cabeçalho de coluna ao FlexGridVisualizar
    
    Dim colIndex As Integer
    colIndex = 0 ' Atribuindo coluna a 0 para começarmos da primeira coluna

    ' Preenchendo o FlexGridVisualizar com os resultados da consulta
    Do While Not result.EOF
        colIndex = colIndex + 1 ' Avançando para proxima coluna
        FlexGridVisualizar.Rows = FlexGridVisualizar.Rows + 1 ' ' Percorre as linhas no FlexGridVisualizar para preencher os resultados
        FlexGridVisualizar.TextMatrix(FlexGridVisualizar.Rows - 1, 0) = result.Fields("nome").Value & " " & result.Fields("sobrenome").Value
        result.MoveNext
    Loop
    
    FlexGridVisualizar.ColWidth(0) = 2000
    ' Fechando o objeto resultado (Recordset) e desconectando do banco de dados
    result.Close
    conn.Desconectar
    Exit Sub

TratarErro:
   MsgBox "Erro ao carregar lista de contatos: " & Err.Description
   RegistrarErro Err.Description
   
End Sub
Private Sub btnExcluir_Click()
ExcluirRegistro
End Sub



Private Sub Form_Load()

End Sub

Private Sub Picture1_Click()

   Set DataGrid.DataSource = Nothing ' Limpar o controle DataGrid
   
   Dim pesquisa As String
   pesquisa = Trim(txtBoxPesquisa.Text) ' Atribuindo valor da txtBox a pesquisa
   
   Dim sql As String ' Criando vairavel da consulta sql
   sql = "SELECT * FROM contato WHERE nome LIKE '%" & pesquisa & "%' OR sobrenome LIKE '%" & pesquisa & "%';"
   
   Adodc.RecordSource = sql ' Atribuindo consulta
   
   Adodc.Refresh ' Atualizando com informações da consulta
   
   If Adodc.Recordset.EOF Then ' Verifica se o Recordset está vazio
      MsgBox "Contato não encontrado."
      txtBoxPesquisa.Text = "" ' Limpa o campo de pesquisa
   Else
      Set DataGrid.DataSource = Adodc.Recordset ' Inserindo informações da busca
   End If
   
End Sub


Private Sub RegistrarErro(ByVal Mensagem As String)

    Dim fileName As String
    Dim fileNumber As Integer

    fileName = "C:\Users\Gabrielly Castro\Desktop\Desafio VB6\Log.txt"
    fileNumber = FreeFile

    Open fileName For Append As fileNumber
    Print #fileNumber, Now & " - " & Mensagem
    Close fileNumber
    
End Sub
Private Sub EditarRegistros()

   Dim linha As Integer
   Dim coluna As Integer
   Dim dadosDoDataGrid As String
   
   If Adodc.Recordset.EOF Then
      MsgBox "O DataGrid está vazio."
      Exit Sub
   End If
   
   Dim totalDeColunas As Integer
   totalDeColunas = 5 ' Definindo quantidade de colunas
   
   Do While Not Adodc.Recordset.EOF
      ' Loop para percorrer cada coluna
      For coluna = 0 To totalDeColunas - 1
         Dim valor As String
         valor = Adodc.Recordset.Fields(coluna).Value ' Obter o valor da célula atual
         dadosDoDataGrid = dadosDoDataGrid & valor & vbCrLf
      
         ' Construindo o select
         Dim sqlUpdate As String
         sqlUpdate = "UPDATE contato SET "
         sqlUpdate = sqlUpdate & "nome = '" & valor & "', "
         sqlUpdate = sqlUpdate & "sobrenome = '" & valor & "', "
         sqlUpdate = sqlUpdate & "email = '" & valor & "', "
         sqlUpdate = sqlUpdate & "telefone = '" & valor & "';"
         sqlUpdate = sqlUpdate & "WHERE id = " & idContato & ";"
      Next coluna
      
      Adodc.Recordset.MoveNext 'Executando o SQL e movendo para o próximo registro
   Loop

    If Not erro Then ' Se não houve erro durante a edição
      MsgBox "Cadastro editado com sucesso."
   Else
      MsgBox "Erro ao editar cadastro."
   End If

End Sub

Private Sub ExcluirRegistro()

  If Not Adodc.Recordset.EOF Then
      Dim id As Integer
      id = Adodc.Recordset.Fields("id").Value ' Obtém o ID do registro selecionado no DataGrid
      
      ' Chama a função para excluir o contato com base no ID
      excluirContato id
      
      ' Atualiza o DataGrid após a exclusão
      Adodc.Refresh
   Else
      MsgBox "Nenhum registro selecionado para exclusão."
   End If
   
End Sub


