Attribute VB_Name = "Módulo1"
Sub Pascal()
Attribute Pascal.VB_ProcData.VB_Invoke_Func = " \n14"
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
   n = Cells(1, 5)
   ' Llena unos "1":
   For i = 1 To n
      Cells(i, 1) = 1
      Cells(i, i) = 1
   Next i
   ' Llena el resto:
   If n > 2 Then
      For i = 3 To n
         For j = 2 To i - 1
            Cells(i, j) = Cells(i - 1, j) + Cells(i - 1, j - 1)
         Next j
      Next i
   End If
End Sub
Sub Borrar()
Attribute Borrar.VB_ProcData.VB_Invoke_Func = " \n14"
   '***************************************************************************
   '* Rutina de Borrado.                                                      *
   '***************************************************************************
   Application.ScreenUpdating = False
   n = Cells(1, 5).Value
   For i = 1 To n
      For j = 1 To i
         Cells(i, j).Value = Null
      Next j
   Next i
End Sub
Sub G_2D()
Attribute G_2D.VB_ProcData.VB_Invoke_Func = " \n14"
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
   n = Cells(6, 5)
   a = Cells(6, 3)
   b = Cells(6, 4)
   h = (b - a) / n
   formula = Cells(2, 3)
   OK = Fun.StoreExpression(formula) 'Lectura de la fórmula
   If Not OK Then GoTo Error_Handler
      For i = 0 To n
         Cells(6 + i, 1) = a + i * h
         Cells(6 + i, 2) = Fun.Eval1(a + i * h)
   Next i
   '----------------------- Eliminar gráficos anteriores-------------
   Set chartsTemp = ActiveSheet.ChartObjects
   If chartsTemp.Count > 0 Then
      chartsTemp(chartsTemp.Count).Delete
   End If
   '-----------------------------------------------------------------
   datos = Range(Cells(6, 1), Cells(6 + n, 2)).Address
   'rango a graficar
   Set graf = Charts.Add 'gráfico y sus caraterísticas
   With graf
      .Name = "Gráfico"
      .ChartType = xlXYScatterSmoothNoMarkers
      .SetSourceData Source:=Sheets("Graficas en 2D").Range(datos), PlotBy:=xlColumns
      .Location Where:=xlLocationAsObject, Name:="Graficas en 2D"
   End With
   With ActiveChart
     .HasTitle = False
     .Axes(xlCategory, xlPrimary).HasTitle = False
     .Axes(xlValue, xlPrimary).HasTitle = False
   End With
   'Mostramos el procedimiento
   Application.ScreenUpdating = True
   ActiveChart.HasLegend = False
   '---------------------------------------------------------------
   If Err Then GoTo Error_Handler
Error_Handler:       Cells(1, 1) = Fun.ErrorDescription 'imprimir mensaje error
   '---------------------------------------------------------------
End Sub
Sub G3D()
Attribute G3D.VB_ProcData.VB_Invoke_Func = " \n14"
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
   fxy = Cells(2, 2)
   xmin = Cells(5, 3)
   xmax = Cells(5, 4)
   ymin = Cells(5, 5)
   ymax = Cells(5, 6)
   n = Cells(3, 2) ' núumero de puntos n x n
   hx = (xmax - xmin) / n
   hy = (ymax - ymin) / n
   If hx > 0 And hy > 0 And n > 0 Then
      For i = 0 To n
         xi = xmin + i * hx
         Cells(7, 2 + i) = xi
         For j = 0 To n
            yi = ymin + j * hy
            Cells(8 + j, 1) = yi
            OK = Fun.StoreExpression(fxy) 'formula actual es 'f(x,y)'
            If Not OK Then GoTo Error_Handler
               Fun.Variable("x") = xi
               Fun.Variable("y") = yi
            Cells(8 + j, 2 + i) = Fun.Eval() 'retorna f(xa,ya)
         Next j
      Next i
   End If
'----------------------- eliminar gráficos anteriores-------------
Set chartsTemp = ActiveSheet.ChartObjects
If chartsTemp.Count > 0 Then
chartsTemp(chartsTemp.Count).Delete
End If
'-----------------------------------------------------------------
datos = Range(Cells(7, 1), Cells(7 + n, n + 2)).Address 'rango a graficar
Range(datos).Select
Selection.NumberFormat = ";;;" 'ocular celdas
Charts.Add
ActiveChart.ChartType = xlSurface
ActiveChart.SetSourceData Source:=Sheets("Graficas en 3D").Range(datos), PlotBy:=xlColumns
ActiveChart.Location Where:=xlLocationAsObject, Name:="Graficas en 3D"
'---------------------------------------------------------------
If Err Then GoTo Error_Handler
Error_Handler: Cells(1, 1) = Fun.ErrorDescription 'enviar un mensaje de error
'---------------------------------------------------------------
ActiveChart.HasLegend = False
End Sub
Sub Romberg()
Attribute Romberg.VB_ProcData.VB_Invoke_Func = " \n14"
   Application.ScreenUpdating = False
   ' Integración de Romber
   Dim R() As Double
   Dim a, b, h, suma As Double
   Dim n As Integer
   Dim formula As String
   Dim OK As Boolean
   Dim Fun As New clsMathParser ' así se llama el módulo de clase aquí³
   formula = Cells(2, 2)
   a = Cells(2, 3)
   b = Cells(2, 4)
   n = Cells(2, 5)
   ReDim R(n, n)
   h = b - a
   OK = Fun.StoreExpression(formula) 'formula actual es 'formula'
   If Not OK Then GoTo Error_Handler
   '-------------------------------------------------------------------
   For i = 1 To 50 'limpiar
      For j = 1 To 50
         Cells(2 + i, j) = Null
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
   If Err Then GoTo Error_Handler
Error_Handler:        Cells(1, 1) = Fun.ErrorDescription
   '---------------------------------------------------------------
End Sub
