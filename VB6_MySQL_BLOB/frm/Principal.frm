VERSION 5.00
Begin VB.Form Principal 
   AutoRedraw      =   -1  'True
   Caption         =   "Selecione sua IMAGEM"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAtualizarImagem 
      Caption         =   "Atualizar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtNumeroImagem 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton btnProxima 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton btnAnterior 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton btnSalvarImagem 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnCarregaImagem 
      Caption         =   "Carrega"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image Imagem 
      Height          =   5055
      Left            =   120
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintContador As Integer
Dim mstrCaminhoImagem As String
Private Sub btnAnterior_Click()
On Error GoTo btnAnterior_Click_Error
Dim lrstRecordSet As New ADODB.Recordset
    
    If mintContador = 0 Then
        MsgBox "Impossível pesquisar com o contador Zerado.", vbCritical, "Atenção!"
        btnAtualizarImagem.Enabled = False
        GoTo Fim_Error
    Else
       mintContador = mintContador - 1
       txtNumeroImagem.Text = mintContador
       If mintContador = 0 Then GoTo Fim_Error
       btnAtualizarImagem.Enabled = True
       If Not CarregaImagenDoBanco Then
            MsgBox "Falha ao carregar imagem do banco.", vbCritical, "Atenção!"
       End If
    End If

Fim_Error:
    Exit Sub

btnAnterior_Click_Error:
    MsgBox "Erro " & Err.Number & " (" & Err.Description & ") na função btnAnterior_Click do Formulário Principal", vbCritical, "ERRO"
    GoTo Fim_Error
End Sub

Private Sub btnAtualizarImagem_Click()
On Error GoTo btnAtualizarImagem_Click_Error

Dim lrstMyStream As New ADODB.Stream
Dim lrstRecordSet As New ADODB.Recordset
    If mintContador > 0 Then
        'Buscando o Registro
        Set lrstRecordSet = SQLRecordSet("SELECT * FROM Album Where ID = " & mintContador)
        'Criando uma nova instancia de ADODB.Stream
        Set lrstMyStream = New ADODB.Stream
        'Definindo o tipo do ADODB.Stream
        lrstMyStream.Type = adTypeBinary
    
        lrstMyStream.Open
        lrstMyStream.LoadFromFile FileName
        lrstRecordSet!Imagem = lrstMyStream.Read
        lrstRecordSet.Update
        lrstMyStream.Close
    
         Set lrstRecordSet = SQLRecordSet("SELECT max(Id) as ID_Max FROM Album")
         txtNumeroImagem.Text = lrstRecordSet!ID_Max
         mintContador = lrstRecordSet!ID_Max
         btnSalvarImagem.Enabled = False
         MsgBox "Imagem atualizada com Sucesso!", vbInformation, "Sucesso!"
    End If


Fim_Error:
    Exit Sub

btnAtualizarImagem_Click_Error:
    MsgBox "Erro " & Err.Number & " (" & Err.Description & ") na função btnAtualizarImagem_Click do Formulário Principal", vbCritical, "ERRO"
    GoTo Fim_Error
End Sub

Private Sub btnCarregaImagem_Click()
Dim sFiltro As String
    'Carrega a Imagem
    Title = "Selecione "
    sFiltro = "Arquivos JPG" & vbNullChar & "*.jpg;" & vbNullChar
    sFiltro = sFiltro & "All Files" & vbNullChar & "*.*" & String$(2, 0)
    PesquisaWindows.Filter = sFiltro
    PesquisaWindows.FileName = ""
    PesquisaWindows.Windows_Show
    mstrCaminhoImagem = FileName
    Set Imagem.Picture = LoadPicture(FileName)
    'Redimensionar
    Imagem.Height = 5055
    Imagem.Width = 6495
    Imagem.Stretch = True
    btnSalvarImagem.Enabled = True
End Sub

Private Function CarregaImagenDoBanco() As Boolean
On Error GoTo CarregaImagenDoBanco_Error
Dim lrstMyStream As New ADODB.Stream
Dim lrstRecordSet As New ADODB.Recordset

    'Buscando o Registro
    Set lrstRecordSet = SQLRecordSet("SELECT * FROM Album WHERE ID = " & mintContador)
    'Criando uma nova instancia de ADODB.Stream
    Set lrstMyStream = New ADODB.Stream
    'Definindo o tipo do ADODB.Stream
    lrstMyStream.Type = adTypeBinary

    If Not IsNull(lrstRecordSet!Imagem) And lrstRecordSet.EOF <> True Then
        lrstMyStream.Open
        'Carregando a imagem do bando para o ADODB.Stream
        lrstMyStream.Write lrstRecordSet!Imagem
        'Salvando a imagem em um arquivo para poder carrega-lá no campo Imagem
        lrstMyStream.SaveToFile App.Path & "\Imagem.jpg", adSaveCreateOverWrite
        'Carregando a imagendo no campo imagem.
        Imagem.Picture = LoadPicture(App.Path & "\Imagem.jpg")
        'Redimensionando
        Imagem.Height = 5055
        Imagem.Width = 6495
        Imagem.Stretch = True
        'Apagando a imagem depois de carregada
        Kill App.Path & "\Imagem.jpg"
        lrstMyStream.Close
    Else
        MsgBox "Falha ao carregar a imagem.", vbCritical, "Atenção!"
        mintContador = 0
        btnAtualizarImagem.Enabled = False
        txtNumeroImagem.Text = mintContador
        Imagem.Picture = LoadPicture("")
    End If
    
    CarregaImagenDoBanco = True
Fim_Error:
    Set lrstRecordSet = Nothing
    Exit Function
CarregaImagenDoBanco_Error:
    MsgBox "Erro " & Err.Number & " (" & Err.Description & ") na função CarregaImagenDoBanco do Formulário Principal", vbCritical, "ERRO"
    GoTo Fim_Error
End Function

Private Sub btnProxima_Click()
       mintContador = mintContador + 1
       txtNumeroImagem.Text = mintContador
       If mintContador > 0 Then
            btnAtualizarImagem.Enabled = True
       Else
            btnAtualizarImagem.Enabled = False
       End If
       If Not CarregaImagenDoBanco Then
            MsgBox "Falha ao carregar imagem do banco.", vbCritical, "Atenção!"
       End If
End Sub

Private Sub btnSalvarImagem_Click()
On Error GoTo btnSalvarImagem_Click_Error

Dim lrstMyStream As New ADODB.Stream
Dim lrstRecordSet As New ADODB.Recordset

    'Buscando o Registro
    Set lrstRecordSet = SQLRecordSet("SELECT * FROM Album")
    'Criando uma nova instancia de ADODB.Stream
    Set lrstMyStream = New ADODB.Stream
    'Definindo o tipo do ADODB.Stream
    lrstMyStream.Type = adTypeBinary

    lrstMyStream.Open
    lrstRecordSet.AddNew
    lrstMyStream.LoadFromFile FileName
    lrstRecordSet!Imagem = lrstMyStream.Read
    lrstRecordSet.Update
    lrstMyStream.Close

     Set lrstRecordSet = SQLRecordSet("SELECT max(Id) as ID_Max FROM Album")
     txtNumeroImagem.Text = lrstRecordSet!ID_Max
     mintContador = lrstRecordSet!ID_Max
     btnSalvarImagem.Enabled = False
     MsgBox "Imagem adicionada ao banco com Sucesso!", vbInformation, "Sucesso!"
Fim_Error:
    Set lrstMyStream = Nothing
    Set lrstRecordSet = Nothing
    Exit Sub
btnSalvarImagem_Click_Error:
    MsgBox "Erro " & Err.Number & " (" & Err.Description & ") na função btnSalvarImagem_Click do Formulário Principal", vbCritical, "ERRO"
    GoTo Fim_Error
End Sub



Private Sub Form_Load()
    If Not Conexao_Open_DBConexao Then
        MsgBox "Falha ao iniciar conexão com o banco" & vbEnter & Err.Number & " - " & Err.Description, vbCritical, "Atenção"
        Unload Me
    End If
End Sub
