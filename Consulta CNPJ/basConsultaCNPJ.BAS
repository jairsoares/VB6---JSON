Attribute VB_Name = "ConsultaCNPJ"
' Jair Soares
' jairsoares.css@gmail.com
' CSS-Sistemas
'------------------------------------


Option Explicit

Const cWebserviceCNPJ = "https://www.receitaws.com.br/v1/cnpj/"

Public JB As JsonBag



Public Function FazConsultaCNPJ(CNPJ As String) As String
    Dim objXML As Object
    Set objXML = CreateObject("Microsoft.XMLHTTP")
    objXML.Open "GET", cWebserviceCNPJ + CNPJ, False
    objXML.send
    If (objXML.Status = 404) Then
        FazConsultaCNPJ = "404 Error"
    Else
        FazConsultaCNPJ = objXML.responseText
    End If
    Set objXML = Nothing
End Function



Public Function ObtenDados(JSON, Chave As String) As String
Dim i   As Integer
Dim ret As String

ret = Mid(JSON, InStr(JSON, Chave) + Len(Chave) + 4, Len(JSON))
ret = Mid(ret, 1, InStr(ret, Chr(34)) - 1)
ObtenDados = ret

End Function

