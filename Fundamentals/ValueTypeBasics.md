# Value Type Basics

This discussion is about classic value types like Double, String and Boolean that are used in Excel worksheets, and are passed into and back from VBA or .NET user-defined functions. The new [Linked Data Types](https://support.microsoft.com/en-us/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877) feature of Excel allow extended types like 'Geography' and properties like 'Location' for data exploration in Excel. These 'Linked Data Types' won't be covered in this 'Basics' topic.

I want to explain how different value types are used by Excel for the worksheet data, in VBA for parameter and variable values and in the .NET type system. I'll discuss the  basic data types as they appear in these different settings, with a specific focus on user-defined functions (UDFs) created in VBA vs. those created in Excel-DNA with the .NET languages.

UDFs that are created in VBA or .NET will receive (and return) values from (and to) the worksheet.

* How are the data types and sepcial values related?
* What differs when code is moved from VBA to VB.NET? 

Data type related issues that I don't cover in this 'Basics' topic include:
* Linked Data Types in new versions of Excel
* Data types when dealing with the COM object model.
* Dynamic arrays and how they work with UDFs.
* Data transformations in the UDF implementation - e.g. `Task<>` types for async functions, and 'object handles' that represent references to internal data structures.

## Worksheet value types

A cell in Excel has a value, and possibly a formula associated with it. If it has a formula, the cell has a value that is computed from the formula.
The cell values will have one of the following types:
* Empty - note - never empty if the cell has a formula - the formula can't result in an 'empty' value, only possibly in an empty string
* Double - 64-bit floating point numbers
* String - represented as a Unicode string with at most 32767 characters. An empty (0-length) string is not the same as an empty cell.
* Boolean - either TRUE or FALSE
* Error - various possible errors values like #VALUE!

> **Notes**
> * There is no date / time value type in the list. For Excel, Date/Time is a formatting options for internal double values that represent the date and time. So date/time display formatting is similar to font selection or colours applied to a cell - it is a diaply option and not an internal value.
> * Similar to date / time, there is no special currency type. The currency formatting is a display option for number values.
> * A cell itself cannot contain an array of values. With the 'Dynamic Arrays' feature, a cell can become the anchor cell for an split region.
> * Excel does not store integer (whole number) values in a cell - it uses floating point values to represent all numbers.

For a UDF called from an Excel sheet, there are a few other types involved.
A parameter of the UDF might have additional value types:
* Missing - if no value is specified for the parameter in the formula
* 2D Array of basic values - from an array literal, a sheet array reference or from a Dynamic Arrays spill range reference.
* Sheet Reference - This is one area where VBA and Excel-DNA will differ a bit, and is discussed below.

> **Notes**
> * Values from named ranges will have same set of value types

## UDF value types in VBA

In this section I look at how the worksheet data types relate to VBA data types.

Suppose the UDF accepts one parameter. In VBA we'd write that as 
```vb
Function MyFunc(input)
End Function
```

If we're using data type declaraionts, as is required with 'Option Explicit' the equivalent would be:
```vb
Function MyFunc(input as Variant) As Variant
End Function
```




