# Argument Type Basics

This discussion is about classic data types like Double, String and Boolean that are used in Excel worksheets, and are passed into and back from VBA or .NET user-defined functions. The new [Linked Data Types](https://support.microsoft.com/en-us/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877) feature of Excel allow extended types like 'Geography' and properties like 'Location' for data exploration in Excel. These 'Linked Data Types' won't be covered in this 'Basics' topic.

I want to explain how different data types are used by Excel for the worksheet data, in VBA for UDF arguments and variable values and correspondingly for UDFs created in the .NET type system. I'll discuss the  basic data types as they appear in these different settings, with a specific focus on user-defined functions (UDFs) created in VBA vs. those created in Excel-DNA with the .NET languages.

Data type related issues that I don't cover in detail in this 'Basics' topic include:
* Linked Data Types in new versions of Excel
* Dynamic arrays and how they work with UDFs.
* Data types when dealing with the COM object model.
* Data transformations in the Excel-DNA UDF implementation - e.g. `Task<>` types for async functions, and 'object handles' that represent references to internal data structures.

> **Argument vs Parameter**
>
> I tend to use these two terms inconsistently, and when I first created Excel-DNA I followed the Excel C API documentation usage, which led me somewhat astray.
> On [StackOverflow](https://stackoverflow.com/questions/156767/whats-the-difference-between-an-argument-and-a-parameter#:~:text=A%20parameter%20is%20a%20variable,pass%20into%20the%20method's%20parameters.&text=Parameter%20is%20variable%20in%20the,that%20gets%20passed%20to%20function.) there is a clear definition of these terms:
> > A parameter is a variable in a method definition. When a method is called, the arguments are the data you pass into the method's parameters. Parameter is variable in the declaration of function. Argument is the actual value of this variable that gets passed to function.

## Worksheet value types

A cell in Excel is empty or it has a value and possibly also a formula associated with it. If it has a formula, the cell has a value that is computed from the formula.
The cell values will have one of the following value types:
* Empty - a cell is never empty if the cell has a formula - the formula or UDF can't result in an 'empty' cell value - though possibly in an empty (0-length) string
* Number (Double) - always represented by 64-bit floating point numbers
* Text (String) - represented as a Unicode string with at most 32767 characters. An empty (0-length) string is not the same as an empty cell.
* Logical (Boolean) - either TRUE or FALSE
* Error - the various possible Excel error values like #VALUE!

For a UDF called from an Excel sheet, there are a few other types involved.
A parameter of the UDF might have additional value types:
* Missing - if no value is specified for the parameter in the formula
* Array of basic values - 1D or 2D array of scalar values, from an array literal, a sheet array reference or from a Dynamic Arrays spill range reference.
* Sheet Reference (Range in VBA) - one area where VBA and Excel-DNA will differ a bit and is discussed in detail below.

> **Notes**
> * There is no date / time value type in the list. For Excel, Date/Time is a formatting options for internal double values that represent the date and time. So date/time display formatting is similar to font selection or colours applied to a cell - it is a display option and not an internal value. But as well soon see, cell values that are formatted as dates are sometimes interpreted as DateTime values in VBA. (Specifically, the `Range.Value` property sometimes returns a COM DateTime or Currency value, in constrast with the `Range.Value2` property which returns a `double` in both these cases.)
> * Similar to date / time, there is no special currency type, with the above caveat about conversion in the `Range.Value` property.
> * A cell itself cannot contain an array of values. With the 'Dynamic Arrays' feature, a cell can become the anchor cell for an spill region.
> * Excel does not store integer (whole number) values in a cell - it uses floating point values to represent all numbers.
> * Values from Defined Names will have the same set of value types as for cells, but can also evaluate to array values.

## Worksheet with various types
In the worksheet 'ArgumentTypes.xlsm' that accompanies this discussion I show all the different data types that cells can have, and how they are described by the built-in Excel `=TYPE(...)` function.

## UDF argument types in VBA

In this section I look at how the worksheet data types relate to VBA types for UDFs.

A single UDF accepts one parameter. In VBA we'd write that as 
```vb
Function MyFunc(input)
End Function
```

If we're using data type declarations, as is required with 'Option Explicit' the equivalent would be:
```vb
Function MyFunc(input as Variant) As Variant
End Function
```

The [VBA Variant data type](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/variant-data-type) is a special type that can contain any type of value.

We can examine the value type if a Variant value using the [`VarType`](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function) function, and some other helpes. The full code is in the workbook that accompanies this tutorial, but the main function looks like this:
```vb
Function ArgumentInfoVba(Optional arg As Variant) As String
    Dim value
    Dim value2

    If IsMissing(arg) Then
        ArgumentInfoVba = "<Missing>"
    ElseIf IsObject(arg) And TypeOf arg Is Range Then
        value = arg.value
        value2 = arg.value2
        If varType(value) = varType(value2) Then
            ArgumentInfoVba = "Range(" & arg.Address & "): " & ArgumentInfoVba(value)
        Else
            ArgumentInfoVba = "Range(" & arg.Address & "): " & ArgumentInfoVba(value) & " (" & ArgumentInfoVba(value2) & ")"
        End If
    ElseIf IsArray(arg) Then
        ArgumentInfoVba = "Array" + ArraySize(arg)
    Else
        ArgumentInfoVba = VarTypeName(varType(arg)) & ": " & CStr(arg)
    End If
        
End Function

' Return a string with the description of the variant type passed in
Function VarTypeName(varType As Integer)
    Select Case varType
        Case vbEmpty: ' 0
            VarTypeName = "vbEmpty" ' Empty (uninitialized)
        ' Omitted here, but there lots more cases like this...
    End Select
End Function

Function ArraySize(value As Variant) As String
    On Error Resume Next ' Nasty hack to allow 2D and 1D arrays
    If varType(value) > vbArray Then
        ArraySize = "(" & LBound(value, 1) & ":" & UBound(value, 1) & ")"
        ArraySize = "(" & LBound(value, 1) & ":" & UBound(value, 1) & "," & LBound(value, 2) & ":" & UBound(value, 2) & ")"
    Else
        ArraySize = "<Not an array>"
    End If
End Function
```

On the 'ArgumentValue' sheet we can now see how different datatype look on the sheet, when a VBA function is called with a reference, and then with a literal value:

| Value Type                   | Example Value                     |       | Excel TYPE | Name            | ArgumentInfoVBA - Reference Argument              | ArgumentInfoVBA - Literal Argument |
|------------------------------|-----------------------------------|-------|------------|-----------------|---------------------------------------------------|------------------------------------|
| Empty Cell                   |                                   |       | 1          | Number          | Range($B$2): vbEmpty:                             | <Missing>                          |
| Number                       | 1.234                             |       | 1          | Number          | Range($B$3): vbDouble: 1.234                      | vbDouble: 1.234                    |
|                              | 42                                |       | 1          | Number          | Range($B$4): vbDouble: 42                         | vbDouble: 42                       |
|                              | 9.87E+201                         |       | 1          | Number          | Range($B$5): vbDouble: 9.87E+201                  | vbDouble: 9.87E+201                |
|     Formatted as Date        | 06-Nov-20                         |       | 1          | Number          | Range($B$6): vbDate: 2020/11/06 (vbDouble: 44141) | vbDouble: 44141                    |
|     Formatted as Currency    | R99.99                            |       | 1          | Number          | Range($B$7): vbCurrency: 99.99 (vbDouble: 99.99)  | vbDouble: 99.99                    |
| String                       | Hello, World!                     |       | 2          | Text            | Range($B$8): vbString: Hello, World!              | vbString: Hello, World!            |
|     Empty (0-length)         |                                   |       | 2          | Text            | Range($B$9): vbString:                            | vbString:                          |
| Boolean                      | TRUE                              |       | 4          | Logical Value   | Range($B$10): vbBoolean: True                     | vbBoolean: True                    |
|                              | FALSE                             |       | 4          | Logical Value   | Range($B$11): vbBoolean: False                    | vbBoolean: False                   |
| Error                        | #DIV/0!                           |       | 16         | Error Value     | Range($B$12): vbError: Error 2007                 | vbError: Error 2007                |
|                              | #VALUE!                           |       | 16         | Error Value     | Range($B$13): vbError: Error 2015                 | vbError: Error 2015                |
|                              | #REF!                             |       | 16         | Error Value     | Range($B$14): vbError: Error 2023                 | vbError: Error 2023                |
|                              | #NAME?                            |       | 16         | Error Value     | Range($B$15): vbError: Error 2029                 | vbError: Error 2029                |
|                              | #NUM!                             |       | 16         | Error Value     | Range($B$16): vbError: Error 2036                 | vbError: Error 2036                |
|                              | #N/A                              |       | 16         | Error Value     | Range($B$17): vbError: Error 2042                 | vbError: Error 2042                |
|                              | #SPILL!                           |       | 16         | Error Value     | Range($B$18): vbError: Error 2045                 | NOTE: =ERROR.TYPE(????)            |
| Ctrl+Shift+Enter 2D Array    | 1                                 | A     | 64         | Array           | Range($B$19:$C$20): Array(1:2,1:2)                | Array(1:2,1:2)                     |
|                              | 0.1                               | FALSE |            |                 |                                                   |                                    |
| Dynamic 2D Array             | 99                                | B     | 64         | Array           | Range($B$21:$C$22): Array(1:2,1:2)                | Array(1:2,1:2)                     |
|                              |                                   | TRUE  |            |                 |                                                   |                                    |
| Single Row Array             | 1                                 | A     | 64         | Array           | Range($B$23:$C$23): Array(1:1,1:2)                | Array(1:2)                         |
| Single Column Array          | 1                                 |       | 64         | Array           | Range($B$24:$B$25): Array(1:2,1:1)                | Array(1:2,1:1)                     |
|                              | A                                 |       |            |                 |                                                   |                                    |
| Array Literal in Single Cell | 1                                 |       | 1          | Number          | Range($B$26): vbDouble: 1                         | Array(1:2)                         |
| Linked Data                  | MICROSOFT CORPORATION (XNAS:MSFT) |       | 128        | Linked Data (?) | Range($B$28): vbError: Error 2015                 |                                    |
|                              | 1993                              |       | 1          | Number          | Range($B$29): vbDouble: 1993                      |                                    |
  
  
The formulas on this sheet look like this:

| Value Type                   | Example Value                     |                    | Excel TYPE     | Name                | ArgumentInfoVBA - Reference Argument | ArgumentInfoVBA - Literal Argument  |
|------------------------------|-----------------------------------|--------------------|----------------|---------------------|--------------------------------------|-------------------------------------|
| Empty Cell                   |                                   |                    | =TYPE(B2)      | =ExcelTypeName(D2)  | =ArgumentInfoVba(B2)                 | =ArgumentInfoVba()                  |
| Number                       | 1.234                             |                    | =TYPE(B3)      | =ExcelTypeName(D3)  | =ArgumentInfoVba(B3)                 | =ArgumentInfoVba(1.234)             |
|                              | 42                                |                    | =TYPE(B4)      | =ExcelTypeName(D4)  | =ArgumentInfoVba(B4)                 | =ArgumentInfoVba(42)                |
|                              | 9.87E+201                         |                    | =TYPE(B5)      | =ExcelTypeName(D5)  | =ArgumentInfoVba(B5)                 | =ArgumentInfoVba(9.87E+201)         |
|     Formatted as Date        | 44141                             |                    | =TYPE(B6)      | =ExcelTypeName(D6)  | =ArgumentInfoVba(B6)                 | =ArgumentInfoVba(44141)             |
|     Formatted as Currency    | 99.99                             |                    | =TYPE(B7)      | =ExcelTypeName(D7)  | =ArgumentInfoVba(B7)                 | =ArgumentInfoVba(99.99)             |
| String                       | Hello, World!                     |                    | =TYPE(B8)      | =ExcelTypeName(D8)  | =ArgumentInfoVba(B8)                 | =ArgumentInfoVba("Hello, World!")   |
|     Empty (0-length)         |                                   |                    | =TYPE(B9)      | =ExcelTypeName(D9)  | =ArgumentInfoVba(B9)                 | =ArgumentInfoVba("")                |
| Boolean                      | TRUE                              |                    | =TYPE(B10)     | =ExcelTypeName(D10) | =ArgumentInfoVba(B10)                | =ArgumentInfoVba(TRUE)              |
|                              | FALSE                             |                    | =TYPE(B11)     | =ExcelTypeName(D11) | =ArgumentInfoVba(B11)                | =ArgumentInfoVba(FALSE)             |
| Error                        | =1/0                              |                    | =TYPE(B12)     | =ExcelTypeName(D12) | =ArgumentInfoVba(B12)                | =ArgumentInfoVba(#DIV/0!)           |
|                              | =SUM("A")                         |                    | =TYPE(B13)     | =ExcelTypeName(D13) | =ArgumentInfoVba(B13)                | =ArgumentInfoVba(#VALUE!)           |
|                              | =OFFSET($A$2, -1, -1)             |                    | =TYPE(B14)     | =ExcelTypeName(D14) | =ArgumentInfoVba(B14)                | =ArgumentInfoVba(#REF!)             |
|                              | =MadeUpName                       |                    | =TYPE(B15)     | =ExcelTypeName(D15) | =ArgumentInfoVba(B15)                | =ArgumentInfoVba(#NAME?)            |
|                              | =1E+200 ^ 2                       |                    | =TYPE(B16)     | =ExcelTypeName(D16) | =ArgumentInfoVba(B16)                | =ArgumentInfoVba(#NUM!)             |
|                              | =VLOOKUP("A", {"B"}, 0,FALSE )    |                    | =TYPE(B17)     | =ExcelTypeName(D17) | =ArgumentInfoVba(B17)                | =ArgumentInfoVba(#N/A)              |
|                              | ={1;2}                            |                    | =TYPE(B18)     | =ExcelTypeName(D18) | =ArgumentInfoVba(B18)                | NOTE: =ERROR.TYPE(????)             |
| Ctrl+Shift+Enter 2D Array    | ={1,"A";0.1,FALSE}                | ={1,"A";0.1,FALSE} | =TYPE(B19:C20) | =ExcelTypeName(D19) | =ArgumentInfoVba(B19:C20)            | =ArgumentInfoVba({1,"A";0.1,FALSE}) |
|                              | ={1,"A";0.1,FALSE}                | ={1,"A";0.1,FALSE} |                | =ExcelTypeName(D20) |                                      |                                     |
| Dynamic 2D Array             | ={99,"B";"",TRUE}                 | ={99,"B";"",TRUE}  | =TYPE(B21#)    | =ExcelTypeName(D21) | =ArgumentInfoVba(B21#)               | =ArgumentInfoVba({99,"B";"",TRUE})  |
|                              | ={99,"B";"",TRUE}                 | ={99,"B";"",TRUE}  |                | =ExcelTypeName(D22) |                                      |                                     |
| Single Row Array             | ={1,"A"}                          | ={1,"A"}           | =TYPE(B23#)    | =ExcelTypeName(D23) | =ArgumentInfoVba(B23#)               | =ArgumentInfoVba({1,"A"})           |
| Single Column Array          | ={1;"A"}                          |                    | =TYPE(B24#)    | =ExcelTypeName(D24) | =ArgumentInfoVba(B24#)               | =ArgumentInfoVba({1;"A"})           |
|                              | ={1;"A"}                          |                    |                |                     |                                      |                                     |
| Array Literal in Single Cell | =@{1,"A"}                         |                    | =TYPE(B26#)    | =ExcelTypeName(D26) | =ArgumentInfoVba(B26#)               | =@ArgumentInfoVba({1,"A"})          |
| Linked Data                  | MICROSOFT CORPORATION (XNAS:MSFT) |                    | =TYPE(B28)     | =ExcelTypeName(D28) | =ArgumentInfoVba(B28)                |                                     |
|                              | =B28.[Year incorporated]          |                    | =TYPE(B29)     | =ExcelTypeName(D29) | =ArgumentInfoVba(B29)                |                                     |

Another situation to consider is when the parameter type of the VBA function is not `Variant` but somethin glike `String` ro `Double`:
```vb
Function ArgumentIntegerVBA(value As Integer)
    ArgumentIntegerVBA = value
End Function
```
In this case Excel will attempt to do a conversion of the input value into the parameter type. That conversion might fail (e.g. if a string if passed to `ArgumentIntegerVBA`) or might round or otherwise change the input to convert it to the parameter type (e.g. when passing a double here, it will be rounded).

### Arrays and Range parameters

Consider a function like `=SumEvenValues` where we want to take an array (or range reference of values and then add the even numbers.
A simple implementation (borrowed from the lovely [Excel-Easy](https://www.excel-easy.com/vba/examples/user-defined-function.html) website) might look like this:
```vb
Function vbaSumEvenNumbers(rng As Range)
    Dim cell As Range
    
    vbaSumEvenNumbers = 0
    For Each cell In rng
        If cell.value Mod 2 = 0 Then
            vbaSumEvenNumbers = vbaSumEvenNumbers + cell.value
        End If
    Next
End Function
```

While simple, this approach to accepting array inputs in function raises a problem, and it will be instructive to try to address it.

Because we are declaring the parameter as 'Range', the function can't be called with the input from a literal array or other function. So these calls will fail: `=vbaSumEvenNumbers({1,2,3,4})` and `=vbaSumEvenNumbers(SEQUENCE(100))`.
(Here the SEQUENCE function is only available under Dynamic Array versions of Excel, but `=SEQUENCE(n)` returns an array of numbers `1, 2, 3, ... n`.)
Another problem with the function is that the performance when called with a large input is likely to be bad, because we are eumerating through the cells in the input, and not reading all the values at once. That's a bit of an implementation detail in the algorithm, so I'll focus on the data type issues for now.

An alternative that takes 'Variant' arguments would have to analyze the input more carefully. It is longer, but more explicit about how the data types are handled.

```vb
Function vbaSumEvenNumbersV(arg As Variant)
    Dim i As Long
    Dim j As Long
    Dim rng As Range
    Dim area As Range
    
    If IsObject(arg) And TypeOf arg Is Range Then
        ' Get the dereferenced values from each Area of the Range and call ourselves with the actual values - then we'll process as below
        Set rng = arg
        For Each area In rng.Areas
            vbaSumEvenNumbersV = vbaSumEvenNumbersV(area.value)
        Next
    ElseIf IsArray(arg) Then
        If ArrayDim(arg) = 1 Then
            ' 1D array
            For i = LBound(arg, 1) To UBound(arg, 1)
                If arg(i) Mod 2 = 0 Then
                    vbaSumEvenNumbersV = vbaSumEvenNumbersV + arg(i)
                End If
            Next
            
        Else
            ' 2D array
            For i = LBound(arg, 1) To UBound(arg, 1)
                For j = LBound(arg, 2) To UBound(arg, 2)
                    If arg(i, j) Mod 2 = 0 Then
                        vbaSumEvenNumbersV = vbaSumEvenNumbersV + arg(i, j)
                    End If
                Next
            Next
        End If
    Else
        ' We have a single value - maybe string, bool etc.
        ' We'll process it using whatever type conversion VBA does when evaluating the 'Mod' operator
        ' Alternatively we could examine further
        If arg Mod 2 = 0 Then
            vbaSumEvenNumbersV = arg
        Else
            vbaSumEvenNumbersV = 0
        End If
    End If
    
End Function
```

But now we can deal with 1D and 2D literal inputs, in addition to the `Range` input.

> Note that there is a little quirk with `Range`s with multiple areas, where calling `Range.Value` will return the values from the first area only. So to really replicate the first function which does `For Each cell in theRange` we need to iterate through all the areas.

On the sheet the different options look like this

| `1` | `=vbaSumEvenNumbers(A1:A5)`          |
|-----|--------------------------------------|
| `2` | `=vbaSumEvenNumbers({1,2,3,4,5})`    |
| `3` | `=vbaSumEvenNumbersV(A1:A5)`         |
| `4` | `=vbaSumEvenNumbersV({1,2,3,4,5})`   |
| `5` | `=vbaSumEvenNumbersV(SEQUENCE(100))` |

With calculated results as discussed

| 1 | 6       |
|---|---------|
| 2 | #VALUE! |
| 3 | 6       |
| 4 | 6       |
| 5 | 2550    |


## UDF Arguments in .NET

We now consider the data type situation for UDFs defined with Excel-DNA in .NET.

Our simple argument description function would look like this:

```vb
Function ArgumentInfoDna(arg As Object) As String

End Function
```

