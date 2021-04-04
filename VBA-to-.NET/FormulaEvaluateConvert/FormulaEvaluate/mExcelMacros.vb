Option Explicit Off
Imports Microsoft.Office.Interop.Excel

Public Module mExcelMacros

    Sub Pascal()
        '***************************************************************************
        '*Macro: Pascal                                                            *
        '*Utilidad: Construye el triangulo de Pascal según la cantidad niveles que *
        '*se indique en la celda E1.                                               *
        '*Construido por: Alvin Correa.                                            *
        '*Fecha:                                                                   *
        '*Observaciones: Esta rutina consiste solamente es el uso de la estrructura*
        '*For-Next, la cual se utiliza dos veces. la primera para llenar los unos  *
        '*del triangulo y la segunda es un doble fornext anidado que genera el tri-*
        '*angulo como tal.                                                         *
        '***************************************************************************
        Application.ScreenUpdating = False
        'Lee la cantidad de niveles:
        n = Cells(1, 5).Value
        ' Llena unos "1":
        For i = 1 To n
            Cells(i, 1).Value = 1
            Cells(i, i).Value = 1
        Next i
        ' Llena el resto:
        If n > 2 Then
            For i = 3 To n
                For j = 2 To i - 1
                    Cells(i, j).Value = Cells(i - 1, j).Value + Cells(i - 1, j - 1).Value
                Next j
            Next i
        End If
    End Sub
    Sub Borrar()
        '***************************************************************************
        '* Rutina de Borrado.                                                      *
        '***************************************************************************
        Application.ScreenUpdating = False
        n = Cells(1, 5).Value
        For i = 1 To n
            For j = 1 To i
                Cells(i, j).Value = Nothing
            Next j
        Next i
    End Sub
    Sub G_2D()
        '***************************************************************************
        '*Macro: G_2D                                                              *
        '*Utilidad: Grafica una función en R^2 dada la regla de correspondencia es_*
        '*crita en la celda C2.                                                    *
        '*Construido por: Alvin Correa.                                            *
        '*Fecha:                                                                   *
        '*Observaciones: Esta rutina consiste en graficar en dos dimensiones una   *
        '*funcion dada.                                                             *
        '*                                                                         *
        '*                                                                         *
        '***************************************************************************
        Application.ScreenUpdating = False
        Dim n As Integer
        Dim h As Double
        Dim formula As String
        Dim graf As Chart
        Dim chartsTemp As ChartObjects 'Contador de charts (gráficos) para eliminar el anterior
        Dim OK As Boolean
        Dim Fun As New clsMathParser
        n = Cells(6, 5).Value
        a = Cells(6, 3).Value
        b = Cells(6, 4).Value
        h = (b - a) / n
        formula = Cells(2, 3).Value
        OK = Fun.StoreExpression(formula) 'Lectura de la fórmula
        If Not OK Then GoTo Error_Handler
        For i = 0 To n
            Cells(6 + i, 1).Value = a + i * h
            Cells(6 + i, 2).Value = Fun.Eval1(a + i * h)
        Next i
        '----------------------- Eliminar gráficos anteriores-------------
        chartsTemp = ActiveSheet.ChartObjects
        If chartsTemp.Count > 0 Then
            chartsTemp.Item(chartsTemp.Count).Delete()  ' VBA to .NET: Not sure why we need .Item here - looks like a bug in the COM definitions
        End If
        '-----------------------------------------------------------------
        datos = Range(Cells(6, 1), Cells(6 + n, 2)).Address
        'rango a graficar
        graf = Charts.Add() 'gráfico y sus caraterísticas
        With graf
            .Name = "Gráfico"
            .ChartType = XlChartType.xlXYScatterSmoothNoMarkers
            .SetSourceData(Source:=Sheets("Graficas en 2D").Range(datos), PlotBy:=XlRowCol.xlColumns)
            .Location(Where:=XlChartLocation.xlLocationAsObject, Name:="Graficas en 2D")
        End With
        With ActiveChart()
            .HasTitle = False
            .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).HasTitle = False
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).HasTitle = False
        End With
        'Mostramos el procedimiento
        Application.ScreenUpdating = True
        ActiveChart.HasLegend = False
        '---------------------------------------------------------------
        If Err.Number <> 0 Then GoTo Error_Handler
Error_Handler: Cells(1, 1).Value = Fun.ErrorDescription 'imprimir mensaje error
        '---------------------------------------------------------------
    End Sub
    Sub G3D()
        '***************************************************************************
        '*Macro: G3D                                                               *
        '*Utilidad: Grafica una función en R^3 dada la regla de correspondencia es_*
        '*crita en la celda B2.                                                    *
        '*Construido por: Alvin Correa.                                            *
        '*Fecha:                                                                   *
        '*Observaciones: Esta rutina consiste en graficar en dos dimensiones una   *
        '*funcion dada.                                                            *
        '*                                                                         *
        '*                                                                         *
        '***************************************************************************
        ' Grafica en 3D
        Application.ScreenUpdating = False
        Dim xmin, xmax, ymin, ymax, hx, hy, xi, yi As Double
        Dim n As Integer
        Dim fxy As String 'función f(x,y)
        Dim graf As Chart
        Dim OK As Boolean
        Dim Fun As New clsMathParser ' Se utiliza para llamar al móduloque interpreta las funciones³
        fxy = Cells(2, 2).Value
        xmin = Cells(5, 3).Value
        xmax = Cells(5, 4).Value
        ymin = Cells(5, 5).Value
        ymax = Cells(5, 6).Value
        n = Cells(3, 2).Value ' núumero de puntos n x n
        hx = (xmax - xmin) / n
        hy = (ymax - ymin) / n
        If hx > 0 And hy > 0 And n > 0 Then
            For i = 0 To n
                xi = xmin + i * hx
                Cells(7, 2 + i).Value = xi
                For j = 0 To n
                    yi = ymin + j * hy
                    Cells(8 + j, 1).Value = yi
                    OK = Fun.StoreExpression(fxy) 'formula actual es 'f(x,y)'
                    If Not OK Then GoTo Error_Handler
                    Fun.Variable("x") = xi
                    Fun.Variable("y") = yi
                    Cells(8 + j, 2 + i).Value = Fun.Eval() 'retorna f(xa,ya)
                Next j
            Next i
        End If
        '----------------------- eliminar gráficos anteriores-------------
        chartsTemp = ActiveSheet.ChartObjects
        If chartsTemp.Count > 0 Then
            chartsTemp(chartsTemp.Count).Delete
        End If
        '-----------------------------------------------------------------
        datos = Range(Cells(7, 1), Cells(7 + n, n + 2)).Address 'rango a graficar
        Range(datos).Select()
        Selection.NumberFormat = ";;;" 'ocular celdas
        Charts.Add()
        ActiveChart.ChartType = XlChartType.xlSurface
        ActiveChart.SetSourceData(Source:=Sheets("Graficas en 3D").Range(datos), PlotBy:=XlRowCol.xlColumns)
        ActiveChart.Location(Where:=XlChartLocation.xlLocationAsObject, Name:="Graficas en 3D")
        '---------------------------------------------------------------
        If Err.Number <> 0 Then GoTo Error_Handler
Error_Handler: Cells(1, 1).Value = Fun.ErrorDescription 'enviar un mensaje de error
        '---------------------------------------------------------------
        ActiveChart.HasLegend = False
    End Sub
    Sub Romberg()
        Application.ScreenUpdating = False
        ' Integración de Romber
        Dim R(,) As Double
        Dim a, b, h, suma As Double
        Dim n As Integer
        Dim formula As String
        Dim OK As Boolean
        Dim Fun As New clsMathParser ' así se llama el módulo de clase aquí³
        formula = Cells(2, 2).Value
        a = Cells(2, 3).Value
        b = Cells(2, 4).Value
        n = Cells(2, 5).Value
        ReDim R(n, n)
        h = b - a
        OK = Fun.StoreExpression(formula) 'formula actual es 'formula'
        If Not OK Then GoTo Error_Handler
        '-------------------------------------------------------------------
        For i = 1 To 50 'limpiar
            For j = 1 To 50
                Cells(2 + i, j).Value = Nothing
            Next j
        Next i
        '-------------------------------------------------------------------
        R(1, 1) = h / 2 * (Fun.Eval1(a) + Fun.Eval1(b))
        'paso3 de algoritmo de Romberg
        For i = 1 To n
            'paso 4
            suma = 0
            For k = 1 To 2 ^ (i - 1)
                suma = suma + Fun.Eval1(a + h * (k - 0.5)) 'evalúa en la fórmula actual
            Next k
            R(2, 1) = 0.5 * (R(1, 1) + h * suma)
            'paso5
            For j = 2 To i
                R(2, j) = R(2, j - 1) + (R(2, j - 1) - R(1, j - 1)) / (4 ^ (j - 1) - 1)
            Next j
            'paso 6 salida R(2,j)
            For j = 1 To i
                Cells(3 + i - 1, j) = R(2, j) 'columnas 2,3,...n
            Next j
            'paso 7
            h = h / 2
            'paso 8
            For j = 1 To i
                R(1, j) = R(2, j)
            Next j
        Next i
        '---------------------------------------------------------------
        If Err.Number <> 0 Then GoTo Error_Handler
Error_Handler: Cells(1, 1).Value = Fun.ErrorDescription
        '---------------------------------------------------------------
    End Sub

End Module