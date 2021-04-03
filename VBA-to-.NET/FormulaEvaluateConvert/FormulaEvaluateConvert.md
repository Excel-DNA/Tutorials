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
    * Add the right default to the function declaration and remove the IsMissing()

  * Approach 3:
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

* `Variant` to `Object`

* Take care with array dimensions in declarations
  * From
  ```vb
  Sub ET_Dump(ByRef ETable() As Variant)
    ReDim ETable(ETtop, 30)
  ```
  * To
  ```vb
  Sub ET_Dump(ByRef ETable(,) As Object)
    ReDim ETable(ETtop, 30)
  ```
  
* `Option Private Module` to `Friend Module ...`

* Rewrite `Array(...)` function calls with `MakeArray` helper function
  ```vb
    Private Function MakeArray(ParamArray doubles() As Double) As Double()
        MakeArray = doubles
    End Function
  ```
  Could also do
  `MyArray = New Double() { 1.1, 1.2, 1.3}`

* Remove type from ReDim calls
  `ReDim BI(n) As Double, DI(n) As Double, BK(n) As Double, DK(n) As Double`
  to
  `ReDim BI(n) : ReDim DI(n) : ReDim BK(n) : ReDim DK(n)`

* Date processing
  ```vb
  Select Case UCase$(SymbConst)
    Case "DATE"  'or date
        retval = CDbl(Date)
    Case "TIME"  'or time
        retval = CDbl(Time)
    Case "NOW"   'or now
        retval = CDbl(Now)
  ```
  to 
  ```vb
  Select Case UCase$(SymbConst)
    Case "DATE"  'or date
        retval = DateTime.Today.ToOADate()
    Case "TIME"  'or time
        retval = (DateTime.MinValue + (DateTime.Now - DateTime.Today)).ToOADate()
    Case "NOW"   'or now
        retval = DateTime.Now.ToOADate()
  ```
  
*
  ```vb
  ris = DateSerial(.Arg(1).Value, .Arg(2).Value, .Arg(3).Value).ToOADate()
  ```
  to
  ```vb
  ris = DateSerial(.Arg(1).Value, .Arg(2).Value, .Arg(3).Value).ToOADate()
  ```

*
  ```vb
  Case symYear : ris = Year(CDate(a))
  ```
  to 
  ```vb
  Case symYear : ris = Year(DateTime.FromOADate(a))
  ```

* Make sure Optional parameters have a default
  ```vb
  Private Function cvDegree(ByVal DMS As String, ByRef angle As Double, Optional ByRef sMsg As String = Nothing) As Boolean
  ```
  to
  ```vb
  Private Function cvDegree(ByVal DMS As String, ByRef angle As Double, Optional ByRef sMsg As String = Nothing) As Boolean
  ```
