# Ribbon Basics

In this tutorial we will add a ribbon extension to an Excel-DNA add-in.
The Office ribbon extensions are defined in an .xml file which we will add to the project.
Then a class is defined to process the ribbon callback methods, either to react to commands like a button press or to update the ribbon interface.
I will also show how to request a ribbon interface update.

For this tutorial I will use a Visual Basic add-in.

Some more advanced topics not covered in this tutorial:
* Comparing the native ribbon interface we use here, with the high-level wrapper provided by VSTO.
* Internals of how the ribbon implementation works in Excel-DNA.

## Preparing the project

Our starting point is a simple Excel-DNA add-in that declares a single UDF as a test.
To prepare the environment and our project for the ribbon extensions, we add two steps.

1. Install XML schemas (optional)

To help get IntelliSense help for the ribbon extension, we can either 
* install the `Excel-DNA XML Schemas` extension to Visual Studio, or 
* install the `ExcelDna.XmlSchemas` package in our add-in.

The first approach requires admin permissions on the machine, but has the advantage of now adding any extra files to the project.

2. Reference the Excel COM interop assemblies

Next we need to add a reference to the COM interop assemblies to our add-in. This will allow us to easily use the Excel COM object model from our add-in.
To do this, I add the NuGet package 'ExcelDna.Interop' to the project. It would also be possible to reference the COM libraries directly.

## Adding the Ribbon xml file

## Adding the Ribbon handler code

