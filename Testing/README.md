# Excel-DNA Testing Helper

`ExcelDna.Testing` is a NuGet package and library that lets you develop automatic tests for Excel models and add-ins, including add-ins developed with Excel-DNA and VBA. Test code is written in C# or Visual Basic and is hosted by the popular [xUnit](https://xunit.net/) test framework, allowing automated tests to run from Visual Stuio or other standard test runners.

Tests developed with the testing helper will run with the add-in loaded inside an Excel instance, and so allow you to test the interaction between an add-in and a real instance of Excel. This type of 'integration testing' can augment 'unit testing' where individual library features are tested in isolation. It is often in the interaction with Excel where the problematic aspects of an add-in are revealed, and developing automated testing for this environment has been difficult.

The testing helper allows flexibility and power in designing automated Excel tests:
* The test code can either run in a separate process that drives Excel through the COM object model, or can be loaded inside the Excel process itself, allowing use of both the COM object model and the C API from the test code.
* Functions, macros and even ribbon commands can be tested.
* Test projects can include pre-populated test workbooks containing spreadsheet models to test or test data.

Running automated tests against Excel does introduce complications:
* Testing requires a copy of Excel to be installed on the machine where the tests are run, so don't work well as automated test for 'continuous integration' environments.
* Test outcomes can depend on the exact version of Excel the is used. This is both an advantage in identifying some 
* Integration tests with Excel can be quite slow to run compared to direct unit testing of functions.

This tutorial will introduce the Excel-DNA testing helper, and show you how to create a test project for your Excel model or add-in.

## Background and prerequisites
* Visual Studio and Excel


* xUnit
[xUnit](https://xunit.net/) is a unit testing tool for the .NET Framework. 

If you are not familiar with unit test frameworks, or with xUnit in particular, you might want to look at or work through the XUnit Getting Started instructions for 
[Using .NET Framework with Visual Studio](https://xunit.net/docs/getting-started/netfx/visual-studio).




## Creating a test project
The test project 

## Testing example

## Technical notes
* Error values - COM vs C API

## Solution layout suggestions

For supporting both functional unit testing and Excel integration testing.

## Reference
