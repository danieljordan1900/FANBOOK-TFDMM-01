Public Sub CalcularSumaInputs()
    Dim shp As Visio.Shape
    Dim txtBox As Object
    Dim sumaTotalH As Double
    Dim categoria As String
    sumaTotalH = 0 ' Inicializar la suma

    ' Obtener la página activa
    Set pagActual = Visio.ActivePage

    ' Recorrer los shapes de la página activa
    For Each shp In pagActual.Shapes
        On Error Resume Next
        Set txtBox = shp.Object  ' Intentar obtener el objeto

        ' Verificar si el objeto es un TextBox de ActiveX
        If Not txtBox Is Nothing Then
            If TypeName(txtBox) = "TextBox" Then
                ' Leer la categoría desde las Propiedades Personalizadas
                categoria = ""
                If shp.CellExists("Prop.Categoria", False) Then
                    categoria = shp.Cells("Prop.Categoria").ResultStr("")
                End If
                
                ' Si el TextBox pertenece a la categoría "H", sumarlo
                If categoria = "H" Then
                    sumaTotalH = sumaTotalH + Val(txtBox.Value)
                End If
            End If
        End If
        On Error GoTo 0
    Next shp

    MsgBox sumaTotalH
    
End Sub