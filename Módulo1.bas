Attribute VB_Name = "M�dulo1"
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim LastRow As Long

    If Not Intersect(Target, Me.Range("A1")) Is Nothing Then
        Application.EnableEvents = False

        ' Encontrar a �ltima linha com dados na coluna A
        LastRow = ThisWorkbook.Sheets("ListaQR").Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Copiar o QR code da c�lula A1 para a pr�xima linha na planilha "ListaQR"
        ThisWorkbook.Sheets("ListaQR").Cells(LastRow, "A").Value = Me.Range("A1").Value

        ' Limpar a c�lula A1 ap�s copiar
        Me.Range("A1").ClearContents

        Application.EnableEvents = True
    End If
End Sub
