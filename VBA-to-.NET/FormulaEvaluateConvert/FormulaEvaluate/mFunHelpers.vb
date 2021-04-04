Imports System.Math

Module mFunHelpers
    Function IsMissing(obj As Object) As Boolean
        IsMissing = obj Is Nothing
    End Function

    Function Sqr(d As Double)
        Sqr = d * d
    End Function

    Function Atn(d As Double)
        Atn = Atan(d)
    End Function

    Function Sgn(d As Double)
        Sgn = Sign(d)
    End Function

    Function MakeArray(ParamArray doubles() As Double) As Double()
        MakeArray = doubles
    End Function
End Module
