Public Sub CalcularSumaInputs()
    Dim shp As Visio.shape
    Dim txtBox As Object
    Dim sumaTotal As Double
    Dim categoria As String
    sumaTotal = 0 ' Inicializar la suma

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
                
                ' Si el TextBox pertenece a la categoría "X", sumarlo
                If categoria = "X" Then
                    sumaTotal = sumaTotal + Val(txtBox.Value)
                End If
            End If
        End If
        On Error GoTo 0
    Next shp

    ' Buscar el TextBox donde mostrar el resultado
    For Each shp In pagActual.Shapes
        On Error Resume Next
        Set txtBox = shp.Object  ' Intentar obtener el objeto
        
        If Not txtBox Is Nothing Then
            If TypeName(txtBox) = "TextBox" Then
                ' Leer el TAG (Categoría)
                categoria = ""
                If shp.CellExists("Prop.Categoria", False) Then
                    categoria = shp.Cells("Prop.Categoria").ResultStr("")
                End If
                
                ' Si el TextBox tiene la categoría "TotalX", mostrar la suma
                If categoria = "TotalX" Then
                    txtBox.Value = sumaTotal
                End If
            End If
        End If
        On Error GoTo 0
    Next shp    
End Sub