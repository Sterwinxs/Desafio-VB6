VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCadastro 
   Caption         =   "Cadastro"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Dados"
      Height          =   2790
      Left            =   1305
      TabIndex        =   11
      Top             =   930
      Width           =   4710
      Begin VB.TextBox txtBoxEmail 
         Height          =   345
         Left            =   135
         TabIndex        =   2
         Top             =   1590
         Width           =   4335
      End
      Begin VB.TextBox txtBoxNome 
         Height          =   345
         Left            =   135
         TabIndex        =   0
         Top             =   420
         Width           =   4230
      End
      Begin VB.TextBox txtBoxSobrenome 
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   1035
         Width           =   4350
      End
      Begin VB.TextBox txtBoxTelefone 
         Height          =   345
         Left            =   105
         TabIndex        =   3
         Top             =   2040
         Width           =   4425
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   165
      Top             =   690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnPesquisarClientes 
      Caption         =   "Pesquisar clientes"
      Height          =   375
      Left            =   75
      TabIndex        =   10
      Top             =   75
      Width           =   1665
   End
   Begin VB.CommandButton btnAdicionar 
      Caption         =   "Adicionar"
      Height          =   450
      Left            =   2025
      TabIndex        =   4
      Top             =   3915
      Width           =   3090
   End
   Begin VB.Label lblTelefone 
      Caption         =   "Telefone"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   270
      TabIndex        =   9
      Top             =   2940
      Width           =   855
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   300
      TabIndex        =   8
      Top             =   2505
      Width           =   795
   End
   Begin VB.Label lblSobrenome 
      Caption         =   "Sobrenome"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   135
      TabIndex        =   7
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   270
      TabIndex        =   6
      Top             =   1395
      Width           =   780
   End
   Begin VB.Label lblTituloTela 
      Caption         =   "Tela Cadastro"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   2655
      TabIndex        =   5
      Top             =   285
      Width           =   2385
   End
End
Attribute VB_Name = "FrmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdicionar_Click()


 ' Validação do campo Nome
   Dim nome As String
   nome = Trim(txtBoxNome.Text)
   
   ' Validação do campo Sobrenome
   Dim sobrenome As String
   sobrenome = Trim(txtBoxSobrenome.Text)
   
   ' Validação do campo Telefone
   Dim telefone As String
   telefone = Trim(txtBoxTelefone.Text)
   
   ' Validação do campo Email
   Dim email As String
   email = Trim(txtBoxEmail.Text)
   
   ' Verificar se todos os campos obrigatórios estão preenchidos
   If nome <> "" And sobrenome <> "" And telefone <> "" And email <> "" Then
      ' Validação do comprimento do telefone
      If (Len(telefone) = 10 Or Len(telefone) = 11) And IsNumeric(telefone) Then
         ' Validação do email
         If IsValidEmail(email) Then
            AdicionarContato
         Else
            MsgBox "O email não é válido. Por favor, insira um email válido."
         End If
      Else
         MsgBox "O telefone deve ter 10 ou 11 caracteres, sendo apenas números."
      End If
   Else
      MsgBox "Por favor, preencha todos os campos obrigatórios."
   End If
   

End Sub

Private Sub btnPesquisarClientes_Click()

   FrmPesquisa.Show ' Abrir Form de pesquisa
   Unload Me ' Fechar Form de Cadastro
   
End Sub

Private Function IsValidEmail(ByVal email As String) As Boolean

   Dim regEx As Object
   Set regEx = CreateObject("VBScript.RegExp")
   
   ' Deve começar com um ou mais caracteres, seguido por um símbolo "@"
   ' Seguido por um ou mais caracteres, pontos
   ' E após os pontos com caracteres também
   regEx.Pattern = "^[\w\.-]+@[\w\.-]+\.\w+$"
   
   IsValidEmail = regEx.Test(email) ' Valida o email
   
End Function

Private Sub AdicionarContato()

   On Error GoTo TratarErro
   
   Dim conn As New Conexao
   Dim email As String
   Dim nome As String
   Dim sobrenome As String
   Dim telefone As String
   Dim sql As String
   Dim rs As Object
   Dim verificaSql As String
   
   email = Trim(txtBoxEmail.Text)
   nome = Trim(txtBoxNome.Text)
   sobrenome = Trim(txtBoxSobrenome.Text)
   telefone = Trim(txtBoxTelefone.Text)
   
   
   
   conn.Conectar ' Conectando ao banco de dados
   
   ' Criando tabela caso ela não exista
   sql = "CREATE TABLE IF NOT EXISTS Contato (" & _
          "id SERIAL PRIMARY KEY," & _
          "nome VARCHAR(255) NOT NULL, " & _
          "sobrenome VARCHAR(255) NOT NULL, " & _
          "email VARCHAR(255) , " & _
          "telefone VarChar(15));"
          
    conn.ExecutarSQL sql
    
    'Inserindo informações das variáveis na tabela Contato
    sql = "INSERT INTO Contato (nome, sobrenome, email, telefone) VALUES ('" & nome & "','" & sobrenome & "','" & email & "','" & telefone & "');"
    conn.ExecutarSQL sql
    
    MsgBox "Contato adicionado com sucesso."
       
    conn.Desconectar 'Fechando conexão
    Exit Sub

TratarErro:
    MsgBox "Erro ao adicionar contato: " & Err.Description
    RegistrarErro Err.Description

End Sub

Private Sub RegistrarErro(ByVal Mensagem As String)

    Dim fileName As String
    Dim fileNumber As Integer

    fileName = "C:\Users\Gabrielly Castro\Desktop\bkp projeto\VB6\Case\Log.txt" ' Substitua pelo caminho do seu arquivo de log
    fileNumber = FreeFile

    Open fileName For Append As fileNumber
    Print #fileNumber, Now & " - " & Mensagem
    Close fileNumber
    
End Sub



Private Sub Form_Load()

End Sub
