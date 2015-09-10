Attribute VB_Name = "PesquisaWindows"
Option Explicit

Private Type OPENFILENAME
    lStructSize     As Long
    hWndOwner       As Long
    hInstance       As Long
    lpstrFilter     As String
    lpstrCusFilter  As String
    nMaxCustFilter  As Long
    nFilterIndex    As Long
    lpstrFile       As String
    nMaxFile        As Long
    lpstrFileTitle  As String
    nMaxFileTitle   As Long
    lpstrInitialDir As String
    lpstrTitle      As String
    Flags           As Long
    nFileOffset     As Integer
    nFileExtension  As Integer
    lpstrDefExt     As String
    lCustData       As Long
    lpfnHook        As Long
    lpTemplateName  As String
End Type

Private OFN As OPENFILENAME
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpOFN As OPENFILENAME) As Long
Private Declare Function VarPtr Lib "VB40032.DLL" (lpVar As Any) As Long

Public hWndOwner As Long
Public FileName As String
Public Filter As String
Public Title As String

Sub Windows_Show()

    Dim iNull As Integer
    Dim sFilter As String
    Dim sNull As String
    Dim sTitle As String
    Dim sFileName As String * 1024
    
    LSet sFileName = FileName & vbNullChar
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = hWndOwner
        .lpstrFilter = Filter
        .lpstrFile = sFileName
        .nMaxFile = Len(sFileName)
        .lpstrTitle = Title
        .Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    End With

    If GetOpenFileName(OFN) Then
        iNull = InStr(OFN.lpstrFile, vbNullChar)
        If iNull Then
            FileName = Left$(OFN.lpstrFile, iNull - 1)
        Else
            FileName = OFN.lpstrFile
        End If
    Else
        FileName = ""
    End If

End Sub

