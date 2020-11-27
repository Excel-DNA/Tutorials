# IntelliSense for VBA functions

This tutorial shows you how to enable in-sheet IntelliSense help for VBA Functions.

> TODO: Pic

When you edit a formula on a worksheet, Excel has two mechanisms to assist you in entering the right function call information.

### The 'Function Arguments' dialog

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

You can find more details about the extension at the [Excel-DNA IntelliSense GitHub site](https://github.com/Excel-DNA/IntelliSense), but here I will focus on the use for VBA functions.

I will use these VBA functions as an example:

```vb
Function TempCelsius(tempInFahrenheit As Double) As Double
    TempCelsius = (tempInFahrenheit - 32#) * 5# / 9#
End Function

Function TempFahrenheit(tempInCelsius As Double) As Double
    TempFahrenheit = (tempInCelsius / 5# * 9#) + 32#
End Function

Function TempDeltaFromHeat(heatDelta As Double, mass As Double, specificHeatCapacity As Double)
    TempDeltaFromHeat = heatDelta / mass / specificHeatCapacity
End Function
```

To enable the Excel-DNA IntelliSense requires the ExcelDna.IntelliSense.xll (or ExcelDna.IntelliSense64.xll) add-in to be loaded, and then function descriptions provided as:
* A special (possibly hidden) worksheet,
* An extra file next to the workbook or add-in with the descriptions in xml format, or
* The same xml format information saved in the 'CustomXML' properties of the workbook.

#### Create function descriptions worksheet

| FunctionInfo      | 1              |                |                |                  |      |                |                      |                   |
|-------------------|----------------|----------------|----------------|------------------|------|----------------|----------------------|-------------------|
| TempCelsius       | converts ...   |                | tempFahrenheit | is the temper... |      |                |                      |                   |
| TempFahrenheit    | converts ...   |                | tempCelsius    | is the temper... |      |                |                      |                   |
| TempDeltaFromHeat | calculates ... | https://www... | heatDelta      | is the amount... | mass | is the mass... | specificHeatCapacity | is the specific.. |

Details of the sheet format are:
* The name of the sheet must be '\_IntelliSense\_'; it may be a hidden sheet
* The first cell (A1) must contain the string 'FunctionInfo'
* The next cell across (B1) must contain the value 1
* From the second row down, each row contains the information for a single function
  * Function name
  * Function description
  * Function help link
  * Argument1 name
  * Argument1 description
  * Argument2 name
  * Argument2 description
  * etc.
 
### Function descriptions xml file
An alternate way to provide the function information to the IntelliSense add-in is with an xml file next to the workbook or add-in file.
For a workbook with the name 'MyWorkbook.xlsm' the IntelliSense file must be named 'MyWorkbook.intellisense.xml'.

The contents of the xml file, matching the above example worksheet, would be

```xml
<IntelliSense xmlns="http://schemas.excel-dna.net/intellisense/1.0">
  <FunctionInfo>
    <Function Name="TempCelsius" 
              Description="Converts the temperature from degrees Fahrenheit to degrees Celsius" >
      <Argument Name="tempInFahrenheit" 
                Description="is the temperature in degrees Fahrenheit" />
    </Function>
    <Function Name="TempFahrenheit" 
              Description="Converts the temperature from degrees Celsius to degrees Fahrenheit " >
      <Argument Name="tempInCelsius" 
                Description="is the temperature in degrees Celsius" />
    </Function>
    <Function Name="TempDeltaFromHeat"
              Description="Calculates the temperature change for a body, given the amount of heat absorbed or released, in K (or equivalantly degrees C)" 
              HelpTopic="https://www.softschools.com/formulas/physics/temperature_formula/640/">
      <Argument Name="heatDelta"
                Description="is the amount of heat absorbed or released (in J)" />
      <Argument Name="mass"
                Description="is the mass of the body (in kg)" />
      <Argument Name="specificHeatCapacity"
                Description="is the specific heat capacity of the substance (in J/Kg/°C)" />
    </Function>
  </FunctionInfo>
</IntelliSense>
```


 
#### Download the ExcelDna.IntelliSense(64).xll add-in

The newest release of the IntelliSense add-in can be found here:
https://github.com/Excel-DNA/IntelliSense/releases

You need to check whether your version of Excel is a 32-bit or 64-bit version, then download the matching .xll file.
To test you can just follow File -> Open and select the .xll file. Installing the add-in so that it opens automatically can be done in the `Alt+t, i` Add-Ins dialog - also at File -> Options -> Add-Ins, Manage: Excel add-ins.


### Application.MacroOptions


