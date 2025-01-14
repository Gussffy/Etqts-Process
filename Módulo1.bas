Attribute VB_Name = "Módulo1"
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim LastRow As Long

    If Not Intersect(Target, Me.Range("A1")) Is Nothing Then
        Application.EnableEvents = False

        ' Encontrar a última linha com dados na coluna A
        LastRow = ThisWorkbook.Sheets("ListaQR").Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Copiar o QR code da célula A1 para a próxima linha na planilha "ListaQR"
        ThisWorkbook.Sheets("ListaQR").Cells(LastRow, "A").Value = Me.Range("A1").Value

        ' Limpar a célula A1 após copiar
        Me.Range("A1").ClearContents

        Application.EnableEvents = True
    End If
End Sub
