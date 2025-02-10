Public Sub CalcularSumaInputs()
    Dim pagActual As Visio.Page
    Dim shp As Visio.shape
    Dim txtBox As Object
    Dim sumaTotal As Double
    Dim categoria As String

    ' Obtener la página activa
    Set pagActual = Visio.ActivePage
    sumaTotal = 0 ' Inicializar la suma

    ' Recorrer los shapes de la página actual
    For Each shp In pagActual.Shapes
        ' Verificar si el Shape contiene un Control ActiveX (TextBox)
        If Not shp.Object Is Nothing Then
            If TypeOf shp.Object Is MSForms.TextBox Then
                Set txtBox = shp.Object
                
                ' Leer el TAG (Categoría) desde las Propiedades Personalizadas
                On Error Resume Next
                categoria = shp.Cells("Prop.Categoria").ResultStr("")
                On Error GoTo 0
                
                ' Si el TextBox pertenece a la categoría "Input", sumarlo
                If categoria = "X" Then
                    sumaTotal = sumaTotal + Val(txtBox.Value)
                End If
            End If
        End If
    Next shp

    ' Buscar el TextBox donde mostrar el resultado
    For Each shp In pagActual.Shapes
        If Not shp.Object Is Nothing Then
            If TypeOf shp.Object Is MSForms.TextBox Then
                Set txtBox = shp.Object
                
                ' Si el TextBox tiene la categoría "Resultado", mostrar la suma
                On Error Resume Next
                categoria = shp.Cells("Prop.Categoria").ResultStr("")
                On Error GoTo 0
                
                If categoria = "TotalX" Then
                    txtBox.Value = sumaTotal
                End If
            End If
        End If
    Next shp
End Sub