# Example: Convert Formula Evaluate

## Notes
* New project - Visual Basic "Class Library (.NET Framework)"
* Wrap code in `Public Class` (except the `Option Explicit On`)
* `Private Type` to [`Private Structure`](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/data-types/how-to-declare-a-structure)
  * Add `Dim` / `Public` to fields in the Structure
* `Property Get` to [`Property`](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/property-statement) and [`Get`](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/get-statement)
  *  From
    ```vb
    Public Property Get Expression() As String
        Expression = Expr
    End Property
    ```
  *  To
    ```vb
    Public ReadOnly Property Expression() As String
        Get
            Expression = Expr
        End Get
    End Property
    ```
* `Property Let` to [`Property`](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/property-statement) and [`Set`](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/set-statement) 
  *  From
    ```vb
    Public Property Let VarValue(ByVal Index As Long, ByVal VarVal As Double)
      If Index <= VTtop Then
        VT(Index).Value = VarVal
        VT(Index).Init = True
        iInit = iInit + 1
      End If
    End Property
    ```
  *  To
    ```vb
    Public WriteOnly Property VarValue(ByVal Index As Long) As Double
        Set(ByVal VarVal As Double)
            If Index <= VTtop Then
                VT(Index).Value = VarVal
                VT(Index).Init = True
                iInit = iInit + 1
            End If
        End Set
    End Property
    ```

* Add brackets to function calls where needed
  *  From `Err.Raise 1001, "MathParser", ErrMsg` to `Err.Raise(1001, "MathParser", ErrMsg)`
  *  From `Catch_Sign SubExpr, Sign` to `Catch_Sign(SubExpr, Sign)`

* Add explicit type to parameters when one of the parameters has a type
  * From `Public Function EvalMulti(ByRef VarValue() As Double, Optional ByVal VarName)` to `Public Function EvalMulti(ByRef VarValue() As Double, Optional ByVal VarName As Object)`

* Optional / `IsMissing` pattern
  * Approach 1:
    * From 
      ```vb
      Public Function EvalMulti(ByRef VarValue() As Double, Optional ByVal VarName)
       ...
       If IsMissing(VarName) Then
       ....
      ```
    * To adding the parameter type and assigning a default of Nothing, then checking this
      ```vb
      Public Function EvalMulti(ByRef VarValue() As Double, Optional ByVal VarName As Object = Nothing)
       ...
       If VarName Is Nothing Then
       ....
      ```
  * Approach 2:
    * Define a function `IsMissing()` As
    ```vb
    Private Function IsMissing(obj As Object) As Boolean
        IsMissing = obj Is Nothing
    End Function
    ```
    and still add default value for the optional parameters.

* `If Err() <> 0 Then` to `If Err.Number <> 0 Then`

* Remove Set for assignment

* Escape keywords with [] brackets: `Dim char As String`  to `Dim [char] As String`

* Undo VS formatting for 
  * From `.FunTok = GetFunTok(Of Char)()` to `.FunTok = GetFunTok([char])`

* Don't mix type characters and explicit types
  * From 
    ```vb
    Dim n&, i&, j&, k&, p&, count_iter&, Node_dup As Boolean
    ```
  * To 
    ```vb
    Dim n&, i&, j&, k&, p&, count_iter&
    Dim Node_dup As Boolean
    ```
* Add `Imports System.Math`

* Define `Sqr`

