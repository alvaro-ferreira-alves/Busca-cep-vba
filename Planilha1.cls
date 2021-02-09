VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub BuscarCep()
    Dim json As String
    Dim jsonObject As Object
    Dim i As Long
    Dim columnCep As Long
    Dim ws As Worksheet
    Dim objHTTP As Object

    columnCep = Planilha1.Cells(Rows.Count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    For i = 2 To columnCep
        Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
        Url = "https://buscacepinter.correios.com.br/app/endereco/carrega-cep-endereco.php?pagina=%2Fapp%2Fendereco%2Findex.php&cepaux=&mensagem_alerta=&endereco=" & Range("A" & i).Value & "&tipoCEP=ALL"
        objHTTP.Open "GET", Url, False
        objHTTP.Send
        strResult = objHTTP.responseText
        json = strResult
        
        Set objetoJson = JsonConverter.ParseJson(json)
        
        If objetoJson("total") = 0 Then
             Range("B" & i).Value = objetoJson("mensagem")
        Else
            Range("B" & i).Value = objetoJson("dados")(1)("localidade")
        End If
    Next
    
    Application.ScreenUpdating = True
    
End Sub

