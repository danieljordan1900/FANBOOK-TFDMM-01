Private Sub Document_PageChanged(ByVal Page As IVPage)
    Dim shp As Visio.Shape
    Dim button As Object
    Dim categoria As String
    
    Dim pagActual As Visio.Page
    Set pagActual = Visio.ActivePage
    
    ' Recorremos todas las formas de la página
    For Each shp In pagActual.Shapes
        ' Intentamos obtener el objeto del botón ActiveX
        On Error Resume Next
        Set button = shp.Object
        On Error GoTo 0
        
        ' Verificamos si es un botón ActiveX
        If Not button Is Nothing Then
            If TypeName(button) = "CommandButton" Then ' Verifica que sea un botón
                ' Leer la propiedad "Categoría"
                If shp.CellExists("Prop.Categoria", False) Then
                    categoria = shp.Cells("Prop.Categoria").ResultStr("")
                    
                    ' Si la categoría es "activo", darle focus real y salir del bucle
                    If LCase(categoria) = "activo" Then
                        DoEvents
                        button.SetFocus ' Intenta dar focus real al botón
                        Exit For ' Salimos tras encontrar el primer botón válido
                    End If
                End If
            End If
        End If
    Next shp
End Sub

Public Sub CalcularSumaInputs()
    Dim sumaTotalH As Double
    Dim sumaTotalW As Double
    Dim maxA As Double
    Dim maxP As Double
    Dim pagActual As Visio.Page
    Set pagActual = Visio.ActivePage

    ' Inicializar sumas
    sumaTotalH = 0
    sumaTotalW = 0
    
    ' Inicializar maximos
    maxA = 0
    maxP = 0

    ' Recorrer los shapes de la página activa
    Dim shp As Visio.Shape
    For Each shp In pagActual.Shapes
        Call ProcesarShape(shp, sumaTotalH, sumaTotalW, maxA, maxP)
    Next shp
    
    For Each shp In pagActual.Shapes
        Call ProcesarShapeTotal(shp, sumaTotalH, sumaTotalW, maxA, maxP)
    Next shp

End Sub

Private Sub ProcesarShape(shp As Visio.Shape, ByRef sumaH As Double, ByRef sumaW As Double, ByRef mA As Double, ByRef mP As Double)
    Dim subShp As Visio.Shape
    Dim txtBox As Object
    Dim categoria As String
    Dim valor As Double

    ' Si el shape es un grupo, recorrer sus sub-shapes
    If shp.Type = visTypeGroup Then
        For Each subShp In shp.Shapes
            Call ProcesarShape(subShp, sumaH, sumaW, mA, mP)
        Next subShp
    Else
        ' Intentar obtener un objeto ActiveX
        On Error Resume Next
        Set txtBox = shp.Object
        On Error GoTo 0

        ' Verificar si es un TextBox válido
        If Not txtBox Is Nothing Then
            If TypeName(txtBox) = "TextBox" Then
                ' Leer la categoría
                categoria = ""
                If shp.CellExists("Prop.Categoria", False) Then
                    categoria = shp.Cells("Prop.Categoria").ResultStr("")
                End If

                ' Intentar convertir el valor a número
                If IsNumeric(txtBox.Value) Then
                    valor = CDbl(txtBox.Value)
                Else
                    valor = 0
                End If

                ' Sumar según la categoría
                Select Case categoria
                    Case "H"
                        sumaH = sumaH + valor
                    Case "W"
                        sumaW = sumaW + valor
                End Select
                
                ' Comparar según la categoría
                Select Case categoria
                    Case "A"
                        If (valor > mA) Then
                            mA = valor
                        End If
                    Case "P"
                        If (valor > mP) Then
                            mP = valor
                        End If
                End Select
            End If
        End If
    End If
End Sub

Private Sub ProcesarShapeTotal(shp As Visio.Shape, ByRef sumaH As Double, ByRef sumaW As Double, ByRef mA As Double, ByRef mP As Double)
    Dim subShp As Visio.Shape
    Dim txtBox As Object
    Dim categoria As String
    Dim valor As Double

    ' Si el shape es un grupo, recorrer sus sub-shapes
    If shp.Type = visTypeGroup Then
        For Each subShp In shp.Shapes
            Call ProcesarShapeTotal(subShp, sumaH, sumaW, mA, mP)
        Next subShp
    Else
        ' Intentar obtener un objeto ActiveX
        On Error Resume Next
        Set txtBox = shp.Object
        On Error GoTo 0

        ' Verificar si es un TextBox válido
        If Not txtBox Is Nothing Then
            If TypeName(txtBox) = "TextBox" Then
                ' Leer la categoría
                categoria = ""
                If shp.CellExists("Prop.Categoria", False) Then
                    categoria = shp.Cells("Prop.Categoria").ResultStr("")
                End If

                ' Sumar según la categoría
                Select Case categoria
                    Case "TotalH"
                        txtBox.Value = sumaH
                    Case "TotalW"
                        txtBox.Value = sumaW
                    Case "MaxA"
                        txtBox.Value = mA
                    Case "MaxP"
                        txtBox.Value = mP
                End Select
            End If
        End If
    End If
End Sub


Private Sub CommandButton1_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton2_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton3_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton4_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton5_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton6_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton7_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton8_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton9_Click()
    CalcularSumaInputs
End Sub




Private Sub CommandButton10_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton11_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton12_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton13_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton14_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton15_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton16_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton17_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton18_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton19_Click()
    CalcularSumaInputs
End Sub



Private Sub CommandButton20_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton21_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton22_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton23_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton24_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton25_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton26_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton27_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton28_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton29_Click()
    CalcularSumaInputs
End Sub




Private Sub CommandButton30_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton31_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton32_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton33_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton34_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton35_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton36_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton37_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton38_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton39_Click()
    CalcularSumaInputs
End Sub




Private Sub CommandButton40_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton41_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton42_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton43_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton44_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton45_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton46_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton47_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton48_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton49_Click()
    CalcularSumaInputs
End Sub




Private Sub CommandButton50_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton51_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton52_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton53_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton54_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton55_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton56_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton57_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton58_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton59_Click()
    CalcularSumaInputs
End Sub




Private Sub CommandButton60_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton61_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton62_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton63_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton64_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton65_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton66_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton67_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton68_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton69_Click()
    CalcularSumaInputs
End Sub



Private Sub CommandButton70_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton71_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton72_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton73_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton74_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton75_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton76_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton77_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton78_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton79_Click()
    CalcularSumaInputs
End Sub



Private Sub CommandButton80_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton81_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton82_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton83_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton84_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton85_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton86_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton87_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton88_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton89_Click()
    CalcularSumaInputs
End Sub



Private Sub CommandButton90_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton91_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton92_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton93_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton94_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton95_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton96_Click()
    CalcularSumaInputs
End Sub

Private Sub CommandButton97_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton98_Click()
    CalcularSumaInputs
End Sub

private Sub CommandButton99_Click()
    CalcularSumaInputs
End Sub


private Sub CommandButton100_Click()
    CalcularSumaInputs
End Sub