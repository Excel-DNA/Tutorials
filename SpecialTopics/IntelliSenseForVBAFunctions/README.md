# IntelliSense for VBA functions

This tutorial shows how to enable in-sheet IntelliSense help for VBA Functions.

> TODO: Pic

When you edit a formula on a worksheet, Excel has two mechanisms to assist you in entering the right function call information.

### The 'Insert Function' dialog

This helps you to to select a function and enter or select the right arguments for the function.

> TODO: pic

### The in-sheet 'Formula AutoComplete' features

While entering the formula in the cell, the function list and function details help is displayed as an IntelliSense-style popup.

> TODO: pic

### Descriptions and help information for user-defined functions (UDFs)

UDFs that are declared in .xll add-ins (like those made in .NET with Excel-DNA) can be registered with function and argument descriptions, and these will display in the 'Insert Function' dialog. For Functions in VBA code (either in an .xlam add-in or in the workbook itself) you can is possible to use the `Application.MacroOptions` method to add function and argument descriptions.

However, information added about UDFs with these mechanism only display in the 'Insert Function' dialog. Excel exposes no built-in mechanism that lets UDFs defined in .xll add-ins or VBA to also participate in the in-sheet 'Formula AutoComplete' feature.

## Excel-DNA IntelliSense extension

To enable in-sheet help for UDFs, I developed the Excel-DNA IntelliSense extension. This extension allows functions defined in .xll add-ins as well as VBA functions to register and show IntelliSense information.

I will use these VBA functions as an example:

```vb
Function TempCelcius(tempInFahrenheit As Double) As Double
    Celcius = (tempInFahrenheit - 32.0) * 5.0 / 9.0
End Function

Function TempFahrenheit(tempInCelcius As Double) As Double
    Fahrenheit = (tempInCelcius /  5.0 * 9.0) + 32.0
End Function

```

To enable the Excel-DNA IntelliSense requires the ExcelDna.IntelliSense.xll (or ExcelDna.IntelliSense64.xll) add-in to be loaded, and then function descriptions provided as:
* A special (possibly hidden) worksheet,
* An extra file next to the workbook or add-in with the descriptions in xml format, or
* The same xml format information saved in the 'CustomXML' properties of the workbook.





