Attribute VB_Name = "Conexao"
Global DBConexao As ADODB.Connection

Function SQLRecordSet(VarTexto As String) As Object
Dim lrdsTbRs As ADODB.Recordset
    Set lrdsTbRs = New ADODB.Recordset
    lrdsTbRs.CursorLocation = adUseServer
    lrdsTbRs.Open VarTexto, DBConexao, adOpenStatic, adLockOptimistic
    Set SQLRecordSet = lrdsTbRs
End Function

Public Function Conexao_Open_DBConexao() As Boolean
   Set DBConexao = New ADODB.Connection
   DBConexao.CursorLocation = adUseClient
   DBConexao.Open "Driver={MySQL ODBC 5.1 Driver};Dsn=vb_mysql_blob;Server=localhost;Database=vb_mysql_blob;User=root;Password=;"
   Conexao_Open_DBConexao = True
End Function

'CREATE TABLE `Album` (
'    `Id` INT NOT NULL AUTO_INCREMENT,
'    `Imagem` LONGBLOB NULL DEFAULT NULL,
'    PRIMARY KEY (`Id`)
')
'COLLATE='latin1_swedish_ci'
'ENGINE = MyISAM;

