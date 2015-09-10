VERSION 5.00
Begin VB.Form frmPostWebService 
   AutoRedraw      =   -1  'True
   Caption         =   "Post Web Service"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmbCarrega 
         Caption         =   "Carrega"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         TabIndex        =   2
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CommandButton cmdEnviaSOAP 
         Caption         =   "Envia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   1
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblTextoProjeto 
         Alignment       =   2  'Center
         Caption         =   $"frmPostWebService.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   4515
      End
   End
End
Attribute VB_Name = "frmPostWebService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCarrega_Click()
    PegaInformacoes
End Sub

Private Sub cmdEnviaSOAP_Click()
    EnviaInformacoes
End Sub

Private Sub Form_Load()
End Sub

Private Sub EnviaInformacoes()
On Error GoTo EnviaInformacoes_Error
Dim strSoapAction As String
Dim strUrl As String
Dim strXml As String
Dim strParam As String
    
    'TAG's do utilizadas pelo Web Services
    strParam = "<nome>QualquerNome</nome>" & _
               "<descricao>Qualquer Descrição</descricao>"
    
    'Endereço do Web Service
    strUrl = "http://meu.domain.com/servico.asmx"
    
    'Ação que está sendo utilizada pelo Web Service.
    strSoapAction = "http://tempuri.org/Add"
    
    'Cria o SOAP.
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<Add xmlns=""http://tempuri.org/"">" & strParam & "</Add>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"

    ' Envia os Requisitos para o WEB Service
    Debug.Print ConexaoWebService(strUrl, strSoapAction, strXml)
    
Fim_Erro:
   Exit Sub

EnviaInformacoes_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EnviaInformacoes of Formulário frmPostWebService"
    GoTo Fim_Erro
End Sub

Private Sub PegaInformacoes()
On Error GoTo PegaInformacoes_Error
Dim strSoapAction As String
Dim strUrl As String
Dim strXml As String
Dim strParam As String
    
    'TAG's do utilizadas pelo Web Services
    strParam = "<id>214</id>"
    
    'Endereço do Web Service
    strUrl = "http://meu.domain.com/servico.asmx"
    
    'Ação que está sendo utilizada pelo Web Service.
    strSoapAction = "http://tempuri.org/Get"
    
    'Cria o SOAP.
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<Get xmlns=""http://tempuri.org/"">" & strParam & "</Get>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"

    ' Envia os Requisitos para o WEB Service
    Debug.Print ConexaoWebService(strUrl, strSoapAction, strXml)
  

Fim_Erro:
   Exit Sub

PegaInformacoes_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PegaInformacoes of Formulário frmPostWebService"
    GoTo Fim_Erro
End Sub

Private Function ConexaoWebService(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String) As String
On Error GoTo GetWebService_Error
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    
    On Error GoTo GetWebService_Error
    
    ' Cria Objetos do DOMDocument e XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Carrega XML
    objDom.async = False
    objDom.loadXML XmlBody

    ' Abre conexão com o serviço
    objXmlHttp.open "POST", AsmxUrl, False
    
    ' Cria headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    
    ' Envia o comando XML
    objXmlHttp.send objDom.xml

    ' Recebe a Resposta do WEB Service
    strRet = objXmlHttp.responseText
    

Fim_Erro:
   
   GetWebService = strRet
   
   Set objXmlHttp = Nothing
   Set objDom = Nothing
   
   Exit Function

GetWebService_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetWebService of Formulário frmPostWebService"
    GoTo Fim_Erro
End Function


Private Function TrataRetornoWebService(ByVal pstrRetorno As String) As String
On Error GoTo TrataRetornoWebService_Error
Dim lstrRetorno As String

    intPos1 = InStr(pstrRetorno, "GetResult>") + 10
    intPos2 = InStr(pstrRetorno, "</GetResult")
    If intPos1 > 7 And intPos2 > 0 Then
        lstrRetorno = Mid(pstrRetorno, intPos1, intPos2 - intPos1)
    End If
    

Fim_Erro:
   Exit Function

TrataRetornoWebService_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TrataRetornoWebService of Formulário frmPostWebService"
    GoTo Fim_Erro
End Function
