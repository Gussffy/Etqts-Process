Attribute VB_Name = "Módulo2"
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

    ' Encontrar a última linha com QR codes na planilha "ListaQR"
    LastRow = wsLista.Cells(wsLista.Rows.Count, "A").End(xlUp).Row

    For i = 1 To LastRow
        qrText = wsLista.Cells(i, 1).Value
        
        ' Inicializar variáveis
        validade = ""
        quantidade = ""
        codigoProduto = ""

        ' Encontrar a posição do número 17
        pos17 = InStr(qrText, "17")
        
        If pos17 > 0 And Len(qrText) > pos17 + 6 Then
            ' Extrair a data de validade
            validade = Mid(qrText, pos17 + 6, 2) & "/" & Mid(qrText, pos17 + 4, 2) & "/20" & Mid(qrText, pos17 + 2, 2)
            
            ' Encontrar a posição do número 30 após a data de validade
            pos30 = InStr(pos17 + 6, qrText, "30")
            
            If pos30 > 0 And Len(qrText) > pos30 + 2 Then
                ' Extrair a quantidade
                quantidade = Mid(qrText, pos30 + 2, 2)
                
                ' Extrair o código do produto (os últimos 5 dígitos)
                codigoProduto = Right(qrText, 5)
            End If
        End If
        
        ' Preenchendo campos na planilha "etiqueta"
        wsEtiqueta.Range("C5").Value = validade  ' Preenche a célula C5 com a data de validade
        wsEtiqueta.Range("C6").Value = quantidade  ' Preenche a célula C6 com a quantidade
        wsEtiqueta.Range("C1").Value = codigoProduto  ' Preenche a célula C1 com o código do produto
        
        ' Imprimir planilha "etiqueta"
        wsEtiqueta.PrintOut
    Next i
End Sub
