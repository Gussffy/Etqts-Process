Attribute VB_Name = "M�dulo2"
Sub ImprimirEtiquetas()
    Dim wsLista As Worksheet
    Dim wsEtiqueta As Worksheet
    Dim qrText As String
    Dim validade As String
    Dim quantidade As String
    Dim codigoProduto As String
    Dim pos17 As Long
    Dim pos30 As Long
    Dim i As Long
    Dim LastRow As Long
    
    Set wsLista = ThisWorkbook.Sheets("ListaQR")
    Set wsEtiqueta = ThisWorkbook.Sheets("etiqueta")

    ' Encontrar a �ltima linha com QR codes na planilha "ListaQR"
    LastRow = wsLista.Cells(wsLista.Rows.Count, "A").End(xlUp).Row

    For i = 1 To LastRow
        qrText = wsLista.Cells(i, 1).Value
        
        ' Inicializar vari�veis
        validade = ""
        quantidade = ""
        codigoProduto = ""

        ' Encontrar a posi��o do n�mero 17
        pos17 = InStr(qrText, "17")
        
        If pos17 > 0 And Len(qrText) > pos17 + 6 Then
            ' Extrair a data de validade
            validade = Mid(qrText, pos17 + 6, 2) & "/" & Mid(qrText, pos17 + 4, 2) & "/20" & Mid(qrText, pos17 + 2, 2)
            
            ' Encontrar a posi��o do n�mero 30 ap�s a data de validade
            pos30 = InStr(pos17 + 6, qrText, "30")
            
            If pos30 > 0 And Len(qrText) > pos30 + 2 Then
                ' Extrair a quantidade
                quantidade = Mid(qrText, pos30 + 2, 2)
                
                ' Extrair o c�digo do produto (os �ltimos 5 d�gitos)
                codigoProduto = Right(qrText, 5)
            End If
        End If
        
        ' Preenchendo campos na planilha "etiqueta"
        wsEtiqueta.Range("C5").Value = validade  ' Preenche a c�lula C5 com a data de validade
        wsEtiqueta.Range("C6").Value = quantidade  ' Preenche a c�lula C6 com a quantidade
        wsEtiqueta.Range("C1").Value = codigoProduto  ' Preenche a c�lula C1 com o c�digo do produto
        
        ' Imprimir planilha "etiqueta"
        wsEtiqueta.PrintOut
    Next i
End Sub
